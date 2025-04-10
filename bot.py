import os
from azure.storage.fileshare import ShareServiceClient
import pandas as pd
from botbuilder.core import ActivityHandler, TurnContext, MessageFactory, ConversationState
from botbuilder.schema import SuggestedActions, CardAction, ActionTypes
from aiohttp import ClientSession
from io import BytesIO

class ReportBot(ActivityHandler):
    """
    A simplified bot that prompts for a quarter (Q1-Q4) and a leader ID, then calls an Azure Function.
    It uses conversation state to store 'quarter' and 'leader_id' only during the current conversation.
    Additionally, it reads an Excel file from Azure File Storage to validate the leader ID.
    """

    def __init__(self, conversation_state: ConversationState):
        """
        Initializes the ReportBot with conversation state and Azure File Storage configuration.
        
        Args:
            conversation_state (ConversationState): Used for tracking conversation-level data.
        """
        super(ReportBot, self).__init__()
        self.conversation_state = conversation_state
        # Conversation property to hold { "quarter": str, "leader_id": str }
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

    async def download_file_from_share(self, filename: str) -> pd.DataFrame:
        """
        Downloads and reads an Excel file from Azure File Storage.
        
        Args:
            filename (str): The name of the file to download.
        
        Returns:
            pd.DataFrame: The DataFrame containing the Excel data.
        """
        directory_client = self.share_client.get_directory_client(self.inputs_directory)
        file_client = directory_client.get_file_client(filename)
        stream = file_client.download_file()
        with BytesIO(stream.readall()) as f:
            df = pd.read_excel(f)
        return df

    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        """
        Greets new users and initiates the quarter selection.
        
        Args:
            members_added (list): List of members added to the conversation.
            turn_context (TurnContext): Context for the current turn.
        """
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Hi there! For today's report, which quarter would you like to use?")
                actions = [
                    CardAction(type=ActionTypes.im_back, title="Q1", value="Q1"),
                    CardAction(type=ActionTypes.im_back, title="Q2", value="Q2"),
                    CardAction(type=ActionTypes.im_back, title="Q3", value="Q3"),
                    CardAction(type=ActionTypes.im_back, title="Q4", value="Q4"),
                ]
                await turn_context.send_activity(
                    MessageFactory.suggested_actions(actions, "Please select a quarter:")
                )

    async def on_message_activity(self, turn_context: TurnContext):
        """
        Processes user messages. First expects a quarter, then expects a leader ID.
        Once both are provided, calls the Azure Function to generate the report.
        
        Args:
            turn_context (TurnContext): Context for the current turn.
        """
        text = turn_context.activity.text.strip().upper()
        conv_data = await self.conversation_data.get(turn_context, {"quarter": None, "leader_id": None})
        
        # Si no se ha seleccionado un quarter aún:
        if conv_data["quarter"] is None:
            if text in ["Q1", "Q2", "Q3", "Q4"]:
                conv_data["quarter"] = text
                await self.conversation_data.set(turn_context, conv_data)
                await turn_context.send_activity(f"You selected {text}. Now, please type your leader ID.")
            else:
                await turn_context.send_activity("Please select a valid quarter (Q1, Q2, Q3, or Q4).")
            return
        
        # Si el quarter está seleccionado pero el leader_id no:
        if conv_data["leader_id"] is None:
            # Descargar y leer el archivo Excel para validar el leader_id
            try:
                users_df = await self.download_file_from_share("Tabla_de_Usuarios_Actualizada.xlsx")
                # Asumiendo que la columna se llama 'Matricula Lider'
                valid_leaders = users_df['Matricula Lider'].astype(str).tolist()
                if text in valid_leaders:
                    conv_data["leader_id"] = text
                    await turn_context.send_activity(f"Great! Generating your report for {conv_data['quarter']} using leader ID {text}...")
                    await turn_context.send_activity("Processing your report, please wait...")
                    await self.call_azure_function(turn_context, conv_data["quarter"], conv_data["leader_id"])
                    
                    # Reiniciar para la próxima conversación
                    conv_data["quarter"] = None
                    conv_data["leader_id"] = None
                    await self.conversation_data.set(turn_context, conv_data)
                else:
                    await turn_context.send_activity(f"Leader ID {text} not found in user table. Please try again.")
            except Exception as e:
                await turn_context.send_activity(f"Error accessing user data: {str(e)}. Please try again.")
            return
        
        await turn_context.send_activity("I didn't understand that. Please select a quarter or type a leader ID.")

    async def call_azure_function(self, turn_context: TurnContext, quarter: str, leader_id: str):
        """
        Calls the Azure Function to generate the report and sends the result back to the user.
        
        Args:
            turn_context (TurnContext): The current turn context.
            quarter (str): The selected quarter (e.g., 'Q2').
            leader_id (str): The typed leader ID.
        """
        azure_function_url = os.getenv("AZURE_FUNCTION_URL", "https://<your-function-app>.azurewebsites.net/api/generate_presentation")
        auth_token = os.getenv("AZURE_FUNCTION_AUTH_TOKEN", "<your-auth-token>")
        
        payload = {
            "q": quarter,
            "matricula_lider": leader_id,
             "agenda": [
                "Punto 1 de la agenda",
                "Punto 2 de la agenda",
                "Punto 3 de la agenda",
                "Punto 8 de la agenda",
                "Punto 9 de la agenda",
                "Punto 10 de la agenda"
            ]
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
                                "title": "Your report is ready!",
                                "subtitle": f"Report for {quarter} using leader ID {leader_id}",
                                "buttons": [
                                    {
                                        "type": "openUrl",
                                        "title": "Download Report",
                                        "value": report_url
                                    }
                                ]
                            }
                        }
                        await turn_context.send_activity(MessageFactory.attachment(hero_card))
                    else:
                        await turn_context.send_activity("Sorry, an error occurred while generating your report. Please try again later.")
        except Exception as e:
            print(f"Error calling Azure Function: {e}")
            await turn_context.send_activity("An unexpected error occurred while processing your request. Please try again later.")