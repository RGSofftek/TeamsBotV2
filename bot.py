import os
import json
from io import BytesIO
from aiohttp import ClientSession
from azure.storage.fileshare import ShareServiceClient
import pandas as pd
from datetime import datetime
from azure.ai.inference import ChatCompletionsClient
from azure.core.credentials import AzureKeyCredential
from botbuilder.core import ActivityHandler, TurnContext, MessageFactory, ConversationState
from botbuilder.schema import SuggestedActions, CardAction, ActionTypes

class ReportBot(ActivityHandler):
    """
    Un bot que saluda al usuario y ofrece opciones iniciales: 'Generar la presentación' y 'Revisar agenda'.
    Permite modificar la agenda de la próxima reunión usando Azure Open AI y genera reportes con la agenda actualizada.
    Incluye los nombres de nuevos miembros del equipo en el payload para la Azure Function.
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
        self.service_url = f"https://{self.storage_account_name}.file.core.windows.net"
        self.service_client = ShareServiceClient(account_url=self.service_url, credential=self.sas_token)
        self.share_client = self.service_client.get_share_client(self.file_share_name)

        # Configuración de Azure Open AI
        api_key = os.getenv("AZURE_OPENAI_KEY")
        base_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT").rstrip("/")
        deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4")
        self.openai_client = ChatCompletionsClient(
            endpoint=f"{base_endpoint}/openai/deployments/{deployment_name}",
            credential=AzureKeyCredential(api_key)
        )

    async def download_file_from_share(self, filename: str) -> pd.DataFrame:
        directory_client = self.share_client.get_directory_client(self.inputs_directory)
        file_client = directory_client.get_file_client(filename)
        stream = file_client.download_file()
        with BytesIO(stream.readall()) as f:
            df = pd.read_excel(f, engine="openpyxl")
        return df

    async def download_json_from_share(self, filename: str) -> dict:
        directory_client = self.share_client.get_directory_client(self.inputs_directory)
        file_client = directory_client.get_file_client(filename)
        stream = file_client.download_file()
        with BytesIO(stream.readall()) as f:
            data = json.load(f)
        return data

    async def get_next_meeting_agenda(self, turn_context: TurnContext) -> list:
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
        prompt = f"""
        Eres un asistente que convierte texto en puntos concisos de agenda.
        Toma este texto y conviértelo en una lista de puntos breves:
        "{user_input}"
        Cada punto debe tener entre 5 y 10 palabras.
        Responde solo con la lista, un punto por línea, sin numeración ni etiquetas.
        """
        payload = {
            "messages": [
                {"role": "system", "content": "Convierte el texto en puntos de agenda concisos."},
                {"role": "user", "content": prompt}
            ],
            "max_tokens": 100,
            "temperature": 0.7,
            "top_p": 1.0
        }
        try:
            response = self.openai_client.complete(payload)
            text = response.choices[0].message.content.strip()
            points = [point.strip() for point in text.split("\n") if point.strip()]
            return points
        except Exception as e:
            return ["Error al generar puntos de agenda."]

    async def get_new_team_members(self, leader_id: str) -> list:
        """
        Obtiene los nombres completos de los nuevos miembros del equipo de un líder.

        Args:
            leader_id (str): Matrícula del líder.

        Returns:
            list: Lista de nombres completos de los nuevos miembros.
        """
        try:
            users_df = await self.download_file_from_share("Tabla_de_Usuarios_Actualizada.xlsx")
            # Filtrar miembros del equipo del líder
            team_members = users_df[users_df['Matricula Lider'].astype(str) == leader_id]
            # Filtrar nuevos miembros (TRUE en "Nuevo miembro")
            new_members = team_members[team_members['Nuevo miembro'] == True]  # noqa: E712
            # Obtener nombres completos
            new_member_names = new_members['Nombre Completo'].tolist()
            return new_member_names
        except Exception as e:
            error_msg = str(e)
            return []

    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("¡Hola! ¿Qué te gustaría hacer hoy?")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                    CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                )

    async def on_message_activity(self, turn_context: TurnContext):
        text = turn_context.activity.text.strip()
        conv_data = await self.conversation_data.get(turn_context, {"quarter": None, "leader_id": None, "state": "initial", "next_meeting_agenda": None})

        if conv_data["state"] == "initial":
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
                await turn_context.send_activity("Por favor, describe los cambios que deseas hacer a la agenda.")
                conv_data["state"] = "awaiting_agenda_input"
                await self.conversation_data.set(turn_context, conv_data)
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida: 'Reescribir toda la agenda' o 'Mantener los puntos anteriores y agregar nuevos'.")
            return

        if conv_data["state"] == "awaiting_agenda_input":
            user_input = text
            new_points = await self.generate_agenda_points(user_input)
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
                await turn_context.send_activity("Cambios confirmados. Ahora generaremos la presentación.")
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
                conv_data.pop("next_meeting_agenda", None)
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity("Cambios descartados. Volviendo al menú inicial.")
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
            if text_upper in ["Q1", "Q2", "Q3", "Q4"]:
                conv_data["quarter"] = text_upper
                conv_data["state"] = "selecting_leader_id"
                if "next_meeting_agenda" not in conv_data or not conv_data["next_meeting_agenda"]:
                    conv_data["next_meeting_agenda"] = await self.get_next_meeting_agenda(turn_context)
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity(f"Seleccionaste {text_upper}. Ahora, por favor ingresa tu ID de líder.")
            else:
                await turn_context.send_activity("Por favor, selecciona un trimestre válido (Q1, Q2, Q3 o Q4).")
            return

        if conv_data["state"] == "selecting_leader_id":
            try:
                users_df = await self.download_file_from_share("Tabla_de_Usuarios_Actualizada.xlsx")
                valid_leaders = users_df['Matricula Lider'].astype(str).tolist()
                if text in valid_leaders:
                    conv_data["leader_id"] = text
                    # Obtener los nombres de los nuevos miembros
                    new_members = await self.get_new_team_members(text)
                    await turn_context.send_activity(f"¡Genial! Generando tu reporte para {conv_data['quarter']} con el ID de líder {text}...")
                    await turn_context.send_activity("Procesando tu reporte, por favor espera...")
                    await self.call_azure_function(turn_context, conv_data["quarter"], conv_data["leader_id"], conv_data.get("next_meeting_agenda", []), new_members)
                    conv_data["quarter"] = None
                    conv_data["leader_id"] = None
                    conv_data["state"] = "initial"
                    conv_data.pop("next_meeting_agenda", None)
                    await self.conversation_data.set(turn_context, conv_data)
                    await turn_context.send_activity("Reporte generado. ¿Qué más puedo hacer por ti?")
                    actions = [
                        CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                        CardAction(type=ActionTypes.im_back, title="Revisar agenda", value="Revisar agenda"),
                    ]
                    await turn_context.send_activity(
                        MessageFactory.suggested_actions(actions, "Por favor, selecciona una opción:")
                    )
                else:
                    await turn_context.send_activity(f"El ID de líder {text} no se encontró en la tabla de usuarios. Intenta de nuevo.")
            except Exception as e:
                error_msg = str(e)
                if "ResourceNotFound" in error_msg:
                    await turn_context.send_activity("No se pudo encontrar el archivo de usuarios en Azure File Storage. Contacta al administrador.")
                else:
                    await turn_context.send_activity(f"Error al acceder a los datos de usuarios: {error_msg}. Intenta de nuevo.")
            return

        await turn_context.send_activity("No entendí eso. Por favor selecciona una opción o sigue el flujo.")

    async def call_azure_function(self, turn_context: TurnContext, quarter: str, leader_id: str, agenda: list, new_members: list):
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
                    else:
                        await turn_context.send_activity("Lo siento, ocurrió un error al generar tu reporte. Intenta de nuevo más tarde.")
        except Exception as e:
            await turn_context.send_activity("Ocurrió un error inesperado al procesar tu solicitud. Por favor, intenta de nuevo más tarde.")