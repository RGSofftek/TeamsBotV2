import os
import json
from aiohttp import web
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    ConversationState,
    MemoryStorage,
)
from botbuilder.schema import Activity
from bot import ReportBot

# Load environment variables from variables.env if it exists (for local development)
if os.path.exists("variables.env"):
    from dotenv import load_dotenv
    load_dotenv("variables.env", override=False)

# Environment Variables:
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")

adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# MemoryStorage for local development; replace with persistent storage in production
memory_storage = MemoryStorage()
conversation_state = ConversationState(memory_storage)

# Create an instance of our simplified ReportBot.
bot = ReportBot(conversation_state)

async def _logic(turn_context):
    """
    The callback that handles each incoming activity.
    """
    await bot.on_turn(turn_context)
    # Save conversation state changes.
    await conversation_state.save_changes(turn_context, False)

async def messages(req: web.Request) -> web.Response:
    """
    Handles incoming requests from the Bot Framework (Emulator, Teams, etc.).
    """
    body = await req.text()
    try:
        activity = Activity().deserialize(json.loads(body))
    except Exception as e:
        return web.Response(status=400, text=f"Invalid request body: {e}")

    auth_header = req.headers.get("Authorization", "")
    response = await adapter.process_activity(activity, auth_header, _logic)
    if response:
        return web.json_response(data=response.body, status=response.status)
    return web.Response(status=201)

app = web.Application()
app.router.add_post("/api/messages", messages)

if __name__ == "__main__":
    web.run_app(app, host="localhost", port=3978)
