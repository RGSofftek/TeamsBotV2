import os
import json
from io import BytesIO
from aiohttp import ClientSession
from azure.storage.fileshare import ShareServiceClient
import pandas as pd
from datetime import datetime
from botbuilder.core import ActivityHandler, TurnContext, MessageFactory, ConversationState
from botbuilder.schema import SuggestedActions, CardAction, ActionTypes

class ReportBot(ActivityHandler):
    """
    Un bot que saluda al usuario y ofrece opciones iniciales basadas en la variable de entorno JUNTA.
    Si JUNTA=TRUE, solo muestra 'Revisar contenido de la sesión'. Si JUNTA=FALSE, muestra 'Generar la presentación' y 'Revisar agenda'.
    Permite modificar la agenda o el contenido de la sesión y genera reportes con la agenda actualizada.
    """

    def __init__(self, conversation_state: ConversationState):
        super(ReportBot, self).__init__()
        self.conversation_state = conversation_state
        self.conversation_data = self.conversation_state.create_property("conversation_data")
        
        # Configuración de Azure File Storage
        self.storage_account_name = os.getenv("STORAGE_ACCOUNT_NAME")
        self.sas_token = os.getenv("SAS_TOKEN")
        self.file_share_name = os.getenv("FILE_SHARE_NAME")
        self.directory_name = os.getenv("DIRECTORY_NAME", "reports")
        self.inputs_directory = f"{self.directory_name}/inputs"
        self.meeting_transcripts_directory = f"{self.directory_name}/meeting_transcripts"
        self.service_url = f"https://{self.storage_account_name}.file.core.windows.net"
        self.service_client = ShareServiceClient(account_url=self.service_url, credential=self.sas_token)
        self.share_client = self.service_client.get_share_client(self.file_share_name)

        # Obtener la matrícula del líder desde la variable de entorno
        self.leader_id = os.getenv("LEADER_ID")
        if not self.leader_id:
            raise ValueError("La variable de entorno LEADER_ID no está configurada.")

        # Leer la variable de entorno JUNTA
        junta_value = os.getenv("JUNTA")
        if junta_value not in ["TRUE", "FALSE"]:
            raise ValueError("La variable de entorno JUNTA debe ser 'TRUE' o 'FALSE'.")
        self.junta_enabled = (junta_value == "TRUE")

    async def download_file_from_share(self, filename: str, directory: str) -> str:
        """
        Descarga un archivo desde Azure File Storage y devuelve su contenido como string.
        
        Args:
            filename (str): Nombre del archivo a descargar.
            directory (str): Directorio donde se encuentra el archivo.
        
        Returns:
            str: Contenido del archivo.
        """
        try:
            directory_client = self.share_client.get_directory_client(directory)
            file_client = directory_client.get_file_client(filename)
            stream = file_client.download_file()
            content = stream.readall().decode('utf-8')
            return content
        except Exception as e:
            raise Exception(f"Error downloading file {filename}: {str(e)}")

    async def download_json_from_share(self, filename: str) -> dict:
        """
        Descarga un archivo JSON desde Azure File Storage y lo parsea.
        
        Args:
            filename (str): Nombre del archivo JSON.
        
        Returns:
            dict: Contenido del archivo parseado.
        """
        try:
            content = await self.download_file_from_share(filename, self.inputs_directory)
            return json.loads(content)
        except Exception as e:
            raise Exception(f"Error parsing JSON file {filename}: {str(e)}")

    async def get_next_meeting_agenda(self, turn_context: TurnContext) -> list:
        """
        Obtiene la agenda de la próxima reunión desde agenda.json.
        
        Args:
            turn_context (TurnContext): Contexto para enviar mensajes.
        
        Returns:
            list: Puntos de la agenda.
        """
        try:
            agenda_data = await self.download_json_from_share("agenda.json")
            meetings = agenda_data.get("meetings", [])
            if not meetings:
                await turn_context.send_activity("No hay reuniones programadas en la agenda.")
                return []
            current_time = datetime.utcnow()
            future_meetings = [
                m for m in meetings
                if datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ") > current_time
            ]
            if not future_meetings:
                await turn_context.send_activity("No hay reuniones futuras en la agenda.")
                return []
            next_meeting = min(
                future_meetings,
                key=lambda m: datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ")
            )
            return next_meeting.get("body", {}).get("agenda", [])
        except Exception as e:
            error_msg = str(e)
            if "ResourceNotFound" in error_msg:
                await turn_context.send_activity("No se encontró el archivo de agenda en Azure File Storage. Contacta al administrador.")
            else:
                await turn_context.send_activity(f"Error al leer la agenda: {error_msg}.")
            return []

    async def generate_agenda_points(self, user_input: str) -> list:
        """
        Convierte la entrada del usuario en una lista de puntos de agenda.
        Valida que cada punto tenga entre 1 y 20 palabras y capitaliza la primera letra.
        
        Args:
            user_input (str): Texto con puntos de agenda separados por líneas o comas.
        
        Returns:
            list: Lista de puntos válidos o mensajes de error.
        """
        if not user_input.strip():
            return ["La entrada no puede estar vacía."]
        
        points = []
        if "\n" in user_input:
            points = [p.strip() for p in user_input.split("\n") if p.strip()]
        else:
            points = [p.strip() for p in user_input.split(",") if p.strip()]
        
        if len(points) > 10:
            return ["Demasiados puntos. Máximo 10 permitidos."]
        
        valid_points = []
        for point in points:
            word_count = len(point.split())
            has_valid_content = any(c.isalpha() for c in point)
            if 1 <= word_count <= 20 and has_valid_content:
                capitalized_point = point[0].upper() + point[1:] if point else point
                valid_points.append(capitalized_point)
            else:
                if not has_valid_content:
                    valid_points.append(f"Punto inválido (debe contener letras): {point}")
                else:
                    valid_points.append(f"Punto inválido (requiere 1-20 palabras): {point}")
        
        if not valid_points:
            return ["No se proporcionaron puntos de agenda válidos."]
        
        return valid_points

    async def save_agenda_to_share(self, agenda: list, turn_context: TurnContext):
        """
        Guarda la agenda actualizada en agenda.json en Azure File Storage.
        
        Args:
            agenda (list): Lista de puntos de la agenda a guardar.
            turn_context (TurnContext): Contexto para enviar mensajes de error.
        """
        try:
            agenda_data = await self.download_json_from_share("agenda.json")
            meetings = agenda_data.get("meetings", [])
            if not meetings:
                await turn_context.send_activity("No hay reuniones programadas para actualizar.")
                return
            
            current_time = datetime.utcnow()
            future_meetings = [
                m for m in meetings
                if datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ") > current_time
            ]
            if not future_meetings:
                await turn_context.send_activity("No hay reuniones futuras para actualizar.")
                return
            
            next_meeting = min(
                future_meetings,
                key=lambda m: datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ")
            )
            
            next_meeting["body"] = next_meeting.get("body", {})
            next_meeting["body"]["agenda"] = agenda
            
            directory_client = self.share_client.get_directory_client(self.inputs_directory)
            file_client = directory_client.get_file_client("agenda.json")
            file_client.upload_file(json.dumps(agenda_data, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            error_msg = str(e)
            await turn_context.send_activity(f"Error al guardar la agenda: {error_msg}.")

    async def get_new_team_members(self, leader_id: str) -> list:
        """
        Obtiene la lista de nuevos miembros del equipo desde un archivo Excel.
        
        Args:
            leader_id (str): ID del líder.
        
        Returns:
            list: Nombres de nuevos miembros.
        """
        try:
            users_df = await self.download_file_from_share("Tabla_de_Usuarios_Actualizada.xlsx", self.inputs_directory)
            team_members = users_df[users_df['Matrícula Líder'].astype(str) == leader_id]
            new_members = team_members[team_members['Nuevo miembro'] == True]
            new_member_names = new_members['Nombre Completo'].tolist()
            return new_member_names
        except Exception as e:
            error_msg = str(e)
            return []

    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        """
        Maneja la adición de nuevos miembros a la conversación.
        
        Args:
            members_added: Lista de miembros añadidos.
            turn_context (TurnContext): Contexto de la conversación.
        """
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("¡Hola! ¿Qué te gustaría hacer hoy?")
                if self.junta_enabled:
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Revisar contenido de la sesión", value="Revisar contenido de la sesión")
                    ]
                else:
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                    ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                )

    async def on_message_activity(self, turn_context: TurnContext):
        """
        Maneja los mensajes del usuario según el estado conversacional.
        
        Args:
            turn_context (TurnContext): Contexto de la conversación.
        """
        text = turn_context.activity.text.strip()
        conv_data = await self.conversation_data.get(turn_context, {
            "quarter": None,
            "state": "initial",
            "next_meeting_agenda": None,
            "session_content": None,
            "agenda_change_history": []
        })

        if conv_data["state"] == "initial":
            if self.junta_enabled and text == "Revisar contenido de la sesión":
                try:
                    content = await self.download_file_from_share("test_processed.txt", self.meeting_transcripts_directory)
                    conv_data["session_content"] = content
                    await self.conversation_data.set(turn_context, conv_data)
                    await turn_context.send_activity(f"Este es el contenido generado a partir de la transcripción de la reunión:\n\n{content}")
                    await self.ask_for_changes(turn_context)
                except Exception as e:
                    error_msg = str(e)
                    if "ResourceNotFound" in error_msg:
                        await turn_context.send_activity("No se encontró el archivo test_processed.txt en Azure File Storage. Contacta al administrador.")
                    else:
                        await turn_context.send_activity(f"Error al leer el contenido: {error_msg}.")
            elif not self.junta_enabled:
                if text == "Generar la presentación":
                    conv_data["state"] = "selecting_quarter"
                    await self.conversation_data.set(turn_context, conv_data)
                    await turn_context.send_activity("¡Genial! Para el reporte, ¿qué trimestre desea usar?")
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Q1", value="Q1"),
                        CardAction(type=ActionTypes.im_back, title="Q2", value="Q2"),
                        CardAction(type=ActionTypes.im_back, title="Q3", value="Q3"),
                        CardAction(type=ActionTypes.im_back, title="Q4", value="Q4"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona un trimestre:")
                    )
                elif text == "Revisar agenda":
                    agenda_items = await self.get_next_meeting_agenda(turn_context)
                    if agenda_items:
                        agenda_message = "Agenda actual de la próxima reunión:\n\n"
                        for item in agenda_items:
                            agenda_message += f"- {item}\n"
                        await turn_context.send_activity(agenda_message)
                        await turn_context.send_activity("¿Deseas hacer modificaciones a esta agenda?")
                        actions = [
                            CardAction(type=ActionTypes.im_back, title="Realizar modificaciones", value="Realizar modificaciones"),
                            CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        ]
                        conv_data["state"] = "awaiting_modification_choice"
                        conv_data["next_meeting_agenda"] = agenda_items
                        await self.conversation_data.set(turn_context, conv_data)
                        await turn_context.send_activity(
                            MessageFactory.suggested_actions(actions, "Selecciona una opción:")
                        )
                else:
                    await turn_context.send_activity("Por favor, selecciona una opción válida: 'Generar la presentación' o 'Revisar agenda'.")
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida.")
            return

        if conv_data["state"] == "awaiting_modification_choice":
            if text == "Generar la presentación":
                conv_data["state"] = "selecting_quarter"
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("¡Genial! Para el reporte, ¿qué trimestre desea usar?")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Q1", value="Q1"),
                    CardAction(type=ActionTypes.im_back, title="Q2", value="Q2"),
                    CardAction(type=ActionTypes.im_back, title="Q3", value="Q3"),
                    CardAction(type=ActionTypes.im_back, title="Q4", value="Q4"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona un trimestre:")
                )
            elif text == "Realizar modificaciones":
                conv_data["state"] = "modifying_agenda"
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("Qué cambios quieres hacer? Selecciona una opción:")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Reescribir toda la agenda", value="Reescribir toda la agenda"),
                    CardAction(type=ActionTypes.im_back, title="Mantener los puntos anteriores y agregar nuevos", value="Mantener los puntos anteriores y agregar nuevos"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Selecciona una opción:")
                )
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida: 'Realizar modificaciones' o 'Generar la presentación'.")
            return

        if conv_data["state"] == "modifying_agenda":
            if text in ["Reescribir toda la agenda", "Mantener los puntos anteriores y agregar nuevos"]:
                conv_data["modification_type"] = text
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity(
                    "Por favor, ingresa los puntos de la agenda, uno por línea o separados por comas. "
                    "Cada punto debe tener entre 1 y 20 palabras y contener letras. "
                    "Ejemplo: 'Discutir metas del equipo, Revisar presupuesto trimestral'."
                )
                conv_data["state"] = "awaiting_agenda_input"
                await self.conversation_data.set(turn_context, conv_data)
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida: 'Reescribir toda la agenda' o 'Mantener los puntos anteriores y agregar nuevos'.")
            return

        if conv_data["state"] == "awaiting_agenda_input":
            user_input = text
            new_points = await self.generate_agenda_points(user_input)
            invalid_points = [p for p in new_points if p.startswith("Punto inválido") or p.startswith("Demasiados puntos") or p.startswith("La entrada no puede")]
            if invalid_points:
                await turn_context.send_activity(
                    "Algunos puntos no cumplen los requisitos:\n" +
                    "\n".join(invalid_points) +
                    "\nPor favor, ingresa puntos con 1-20 palabras y que contengan letras."
                )
                return
            
            # Registrar el cambio en el historial
            old_agenda = conv_data.get("next_meeting_agenda", []).copy()
            modification_type = "rewrite" if conv_data["modification_type"] == "Reescribir toda la agenda" else "append"
            new_agenda = new_points if modification_type == "rewrite" else old_agenda + new_points
            change_entry = {
                "timestamp": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                "modification_type": modification_type,
                "old_agenda": old_agenda,
                "new_agenda": new_agenda,
                "added_points": new_points if modification_type == "append" else []
            }
            conv_data["agenda_change_history"].append(change_entry)
            
            # Actualizar la agenda
            if conv_data["modification_type"] == "Reescribir toda la agenda":
                conv_data["next_meeting_agenda"] = new_points
            else:
                conv_data["next_meeting_agenda"].extend(new_points)
            agenda_message = "Así luciría la nueva agenda:\n\n"
            for point in conv_data["next_meeting_agenda"]:
                agenda_message += f"- {point}\n"
            await turn_context.send_activity(agenda_message)
            await turn_context.send_activity("¿Confirmas estos cambios?")
            actions = [
                CardAction(type=ActionTypes.im_back, title="Confirmar y Generar Presentación", value="Confirmar y Generar Presentación"),
                CardAction(type=ActionTypes.im_back, title="Descartar y volver al inicio", value="Descartar y volver al inicio"),
            ]
            conv_data["state"] = "confirming_agenda_changes"
            await self.conversation_data.set(turn_context, conv_data)
            await turn_context.send_activity(
                MessageFactory.suggested_actions(actions, "Selecciona una opción:")
            )
            return

        if conv_data["state"] == "confirming_agenda_changes":
            if text == "Confirmar y Generar Presentación":
                await turn_context.send_activity("Cambios confirmados.")
                conv_data["state"] = "selecting_quarter"
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("¡Genial! Para el reporte, ¿qué trimestre desea usar?")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Q1", value="Q1"),
                    CardAction(type=ActionTypes.im_back, title="Q2", value="Q2"),
                    CardAction(type=ActionTypes.im_back, title="Q3", value="Q3"),
                    CardAction(type=ActionTypes.im_back, title="Q4", value="Q4"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona un trimestre:")
                )
            elif text == "Descartar y volver al inicio":
                conv_data["state"] = "initial"
                conv_data["next_meeting_agenda"] = None
                conv_data["agenda_change_history"] = []
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("Cambios descartados. Volviendo al menú inicial.")
                if self.junta_enabled:
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Revisar contenido de la sesión", value="Revisar contenido de la sesión")
                    ]
                else:
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                    ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                )
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida: 'Confirmar y Generar Presentación' o 'Descartar y volver al inicio'.")
            return

        if conv_data["state"] == "selecting_quarter":
            text_upper = text.upper()
            if text_upper == "REINTENTAR":
                if not conv_data.get("next_meeting_agenda"):
                    await turn_context.send_activity("No hay cambios de agenda para reintentar. Por favor, revisa la agenda nuevamente.")
                    conv_data["state"] = "initial"
                    await self.conversation_data.set(turn_context, conv_data)
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                    )
                    return
                await turn_context.send_activity(f"Reintentando generar el reporte para {conv_data['quarter']}...")
                await turn_context.send_activity("Procesando tu reporte, por favor espera...")
                success = await self.call_azure_function(
                    turn_context,
                    conv_data["quarter"],
                    self.leader_id,
                    conv_data.get("next_meeting_agenda", []),
                    await self.get_new_team_members(self.leader_id)
                )
                if success:
                    if conv_data.get("next_meeting_agenda"):
                        await self.save_agenda_to_share(conv_data["next_meeting_agenda"], turn_context)
                        await turn_context.send_activity("Reporte generado y agenda actualizada.")
                    else:
                        await turn_context.send_activity("Reporte generado.")
                    conv_data["quarter"] = None
                    conv_data["state"] = "initial"
                    conv_data["next_meeting_agenda"] = None
                    conv_data["agenda_change_history"] = []
                    await self.conversation_data.set(turn_context, conv_data)
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                    )
                else:
                    await turn_context.send_activity("Error al generar la presentación. La agenda no se guardó.")
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Reintentar", value="Reintentar"),
                        CardAction(type=ActionTypes.im_back, title="Volver al inicio", value="Volver al inicio"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                    )
                return
            elif text_upper == "VOLVER AL INICIO":
                conv_data["state"] = "initial"
                conv_data["next_meeting_agenda"] = None
                conv_data["agenda_change_history"] = []
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("Volviendo al menú inicial.")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                    CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                )
                return
            elif text_upper in ["Q1", "Q2", "Q3", "Q4"]:
                conv_data["quarter"] = text_upper
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity(f"¡Genial! Generando tu reporte para {text_upper} con el ID de líder {self.leader_id}...")
                await turn_context.send_activity("Procesando tu reporte, por favor espera...")
                new_members = await self.get_new_team_members(self.leader_id)
                if not conv_data.get("next_meeting_agenda"):
                    conv_data["next_meeting_agenda"] = await self.get_next_meeting_agenda(turn_context)
                success = await self.call_azure_function(
                    turn_context,
                    conv_data["quarter"],
                    self.leader_id,
                    conv_data.get("next_meeting_agenda", []),
                    new_members
                )
                if success:
                    if conv_data.get("next_meeting_agenda"):
                        await self.save_agenda_to_share(conv_data["next_meeting_agenda"], turn_context)
                        await turn_context.send_activity("Reporte generado y agenda actualizada.")
                    else:
                        await turn_context.send_activity("Reporte generado.")
                    conv_data["quarter"] = None
                    conv_data["state"] = "initial"
                    conv_data["next_meeting_agenda"] = None
                    conv_data["agenda_change_history"] = []
                    await self.conversation_data.set(turn_context, conv_data)
                    if self.junta_enabled:
                        actions = [
                            CardAction(type=ActionTypes.im_back, title="Revisar contenido de la sesión", value="Revisar contenido de la sesión")
                        ]
                    else:
                        actions = [
                            CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                            CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                        ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "¿Qué más puedo hacer por ti?:")
                    )
                else:
                    await turn_context.send_activity("Error al generar la presentación. La agenda no se guardó.")
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Reintentar", value="Reintentar"),
                        CardAction(type=ActionTypes.im_back, title="Volver al inicio", value="Volver al inicio"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                    )
                return
            else:
                await turn_context.send_activity("Por favor, selecciona un trimestre válido (Q1, Q2, Q3 o Q4).")
            return

        await turn_context.send_activity("No entendí eso. Por favor selecciona una opción o sigue el flujo.")

    async def ask_for_changes(self, turn_context: TurnContext):
        """
        Pregunta al usuario si desea hacer cambios al contenido de la sesión.
        
        Args:
            turn_context (TurnContext): Contexto de la conversación.
        """
        actions = [
            CardAction(type=ActionTypes.im_back, title="Sí", value="Sí"),
            CardAction(type=ActionTypes.im_back, title="Proceder al envío", value="Proceder al envío"),
        ]
        await turn_context.send_activity(
            MessageFactory.suggested_actions(actions, "¿Quieres hacer algún cambio?")
        )
        conv_data = await self.conversation_data.get(turn_context)
        conv_data["state"] = "reviewing_session_content"
        await self.conversation_data.set(turn_context, conv_data)

    async def call_azure_function(self, turn_context: TurnContext, quarter: str, leader_id: str, agenda: list, new_members: list) -> bool:
        """
        Llama a una Azure Function para generar un reporte y devuelve si fue exitoso.
        
        Args:
            turn_context (TurnContext): Contexto para enviar mensajes.
            quarter (str): Trimestre del reporte.
            leader_id (str): ID del líder.
            agenda (list): Puntos de la agenda.
            new_members (list): Nuevos miembros del equipo.
        
        Returns:
            bool: True si la generación fue exitosa, False en caso contrario.
        """
        azure_function_url = os.getenv("AZURE_FUNCTION_URL", "https://<your-function-app>.azurewebsites.net/api/generate_presentation")
        auth_token = os.getenv("AZURE_FUNCTION_AUTH_TOKEN", "<your-auth-token>")

        payload = {
            "q": quarter,
            "matricula_lider": leader_id,
            "agenda": agenda if agenda else ["Punto 1 de la agenda", "Punto 2 de la agenda"],
            "nuevos_miembros": new_members
        }

        try:
            async with ClientSession() as session:
                async with session.post(
                    azure_function_url,
                    headers={"Authorization": auth_token, "Content-Type": "application/json"},
                    json=payload
                ) as response:
                    if response.status == 200:
                        data = await response.json()
                        report_url = data.get("public_url")
                        hero_card = {
                            "contentType": "application/vnd.microsoft.card.hero",
                            "content": {
                                "title": "¡Tu reporte está listo!",
                                "subtitle": f"Reporte para {quarter} con ID de líder {leader_id}",
                                "buttons": [
                                    {
                                        "type": "openUrl",
                                        "title": "Descargar Reporte",
                                        "value": report_url
                                    }
                                ]
                            }
                        }
                        await turn_context.send_activity(MessageFactory.attachment(hero_card))
                        return True
                    else:
                        await turn_context.send_activity("Lo siento, ocurrió un error al generar tu reporte.")
                        return False
        except Exception as e:
            await turn_context.send_activity(f"Ocurrió un error inesperado: {str(e)}")
            return False