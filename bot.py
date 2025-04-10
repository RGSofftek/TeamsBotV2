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
    Un bot que saluda al usuario y ofrece opciones iniciales: 'Generar la presentación' y 'Revisar agenda'.
    Genera reportes usando una Azure Function con la agenda de la próxima reunión desde un archivo JSON.
    """

    def __init__(self, conversation_state: ConversationState):
        super(ReportBot, self).__init__()
        self.conversation_state = conversation_state
        self.conversation_data = self.conversation_state.create_property("conversation_data")
        
        self.storage_account_name = os.getenv("STORAGE_ACCOUNT_NAME")
        self.sas_token = os.getenv("SAS_TOKEN")
        self.file_share_name = os.getenv("FILE_SHARE_NAME")
        self.directory_name = os.getenv("DIRECTORY_NAME", "reports")
        self.inputs_directory = f"{self.directory_name}/inputs"
        self.service_url = f"https://{self.storage_account_name}.file.core.windows.net"
        self.service_client = ShareServiceClient(account_url=self.service_url, credential=self.sas_token)
        self.share_client = self.service_client.get_share_client(self.file_share_name)

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
        """
        Obtiene la agenda de la próxima reunión desde agenda.json.
        
        Args:
            turn_context (TurnContext): Contexto del turno actual.
        
        Returns:
            list: Lista de puntos de la agenda de la próxima reunión.
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
        conv_data = await self.conversation_data.get(turn_context, {"quarter": None, "leader_id": None, "state": "initial"})
        
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
                try:
                    agenda_data = await self.download_json_from_share("agenda.json")
                    meetings = agenda_data.get("meetings", [])
                    
                    if not meetings:
                        await turn_context.send_activity("No hay reuniones programadas en la agenda.")
                    else:
                        current_time = datetime.utcnow()
                        future_meetings = [
                            m for m in meetings
                            if datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ") > current_time
                        ]
                        if not future_meetings:
                            await turn_context.send_activity("No hay reuniones futuras en la agenda.")
                        else:
                            next_meeting = min(
                                future_meetings,
                                key=lambda m: datetime.strptime(m["start"], "%Y-%m-%dT%H:%M:%SZ")
                            )
                            subject = next_meeting.get("subject", "Sin título")
                            start = next_meeting.get("start", "Sin hora")
                            agenda_items = next_meeting.get("body", {}).get("agenda", [])
                            agenda_message = f"Agenda de la próxima reunión: **{subject}** ({start}):\n\n"
                            for item in agenda_items:
                                agenda_message += f"- {item}\n"
                            await turn_context.send_activity(agenda_message)
                            await turn_context.send_activity("¿Deseas hacer modificaciones a esta agenda?")
                            actions = [
                                CardAction(type=ActionTypes.im_back, title="Realizar modificaciones", value="Realizar modificaciones"),
                                CardAction(type=ActionTypes.im_back, title="Generar la presentación", value="Generar la presentación"),
                            ]
                            conv_data["state"] = "awaiting_modification_choice"
                            conv_data["next_meeting_agenda"] = agenda_items  # Guardar la agenda para usarla en el payload
                            await self.conversation_data.set(turn_context, conv_data)
                            await turn_context.send_activity(
                                MessageFactory.suggested_actions(actions, "Selecciona una opción:")
                            )
                except Exception as e:
                    error_msg = str(e)
                    if "ResourceNotFound" in error_msg:
                        await turn_context.send_activity("No se encontró el archivo de agenda en Azure File Storage. Contacta al administrador.")
                    else:
                        await turn_context.send_activity(f"Error al leer la agenda: {error_msg}. Intenta de nuevo.")
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
                await turn_context.send_activity("Funcionalidad de modificación en desarrollo. Por favor, dime cómo quieres proceder.")
            else:
                await turn_context.send_activity("Por favor, selecciona una opción válida: 'Realizar modificaciones' o 'Generar la presentación'.")
            return

        if conv_data["state"] == "selecting_quarter":
            text_upper = text.upper()
            if text_upper in ["Q1", "Q2", "Q3", "Q4"]:
                conv_data["quarter"] = text_upper
                conv_data["state"] = "selecting_leader_id"
                # Si no hay agenda guardada, obtenerla del JSON
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
                    await turn_context.send_activity(f"¡Genial! Generando tu reporte para {conv_data['quarter']} con el ID de líder {text}...")
                    await turn_context.send_activity("Procesando tu reporte, por favor espera...")
                    await self.call_azure_function(turn_context, conv_data["quarter"], conv_data["leader_id"], conv_data.get("next_meeting_agenda", []))
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
                elif "engine" in error_msg.lower():
                    await turn_context.send_activity("El formato del archivo de usuarios no es válido o falta una dependencia. Contacta al administrador.")
                elif "Matricula Lider" not in users_df.columns:
                    await turn_context.send_activity("El archivo de usuarios no tiene la columna 'Matricula Lider'. Contacta al administrador.")
                else:
                    await turn_context.send_activity(f"Error al acceder a los datos de usuarios: {error_msg}. Intenta de nuevo.")
            return
        
        if conv_data["state"] == "modifying_agenda":
            await turn_context.send_activity("Esperando instrucciones para modificar la agenda. ¿Qué quieres hacer?")
            return
        
        await turn_context.send_activity("No entendí eso. Por favor selecciona una opción o sigue el flujo.")

    async def call_azure_function(self, turn_context: TurnContext, quarter: str, leader_id: str, agenda: list):
        """
        Llama a la Azure Function para generar el reporte con la agenda de la próxima reunión.
        
        Args:
            turn_context (TurnContext): Contexto del turno actual.
            quarter (str): Trimestre seleccionado (ej. 'Q2').
            leader_id (str): ID de líder ingresado.
            agenda (list): Lista de puntos de la agenda de la próxima reunión.
        """
        azure_function_url = os.getenv("AZURE_FUNCTION_URL", "https://<your-function-app>.azurewebsites.net/api/generate_presentation")
        auth_token = os.getenv("AZURE_FUNCTION_AUTH_TOKEN", "<your-auth-token>")
        
        payload = {
            "q": quarter,
            "matricula_lider": leader_id,
            "agenda": agenda if agenda else [
                "Punto 1 de la agenda",
                "Punto 2 de la agenda",
                "Punto 3 de la agenda",
                "Punto 8 de la agenda",
                "Punto 9 de la agenda",
                "Punto 10 de la agenda"
            ]  # Fallback si no hay agenda
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
            print(f"Error al llamar a la Azure Function: {e}")
            await turn_context.send_activity("Ocurrió un error inesperado al procesar tu solicitud. Por favor, intenta de nuevo más tarde.")