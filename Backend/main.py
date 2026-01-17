from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import json
import os
from dotenv import load_dotenv

# Load .env (your GEMINI_API_KEY should be set here)
load_dotenv()

# Configure Gemini API (temporarily disabled due to deprecation)
# genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Load business data
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE_DIR, "business.json"), "r", encoding="utf-8") as f:
    business_data = json.load(f)
    
    
app = FastAPI()

@app.get("/")
def home():
    return {"message": "Svaadh Kitchen Backend is Live 🍲"}

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or restrict to your website domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temporarily disabled model due to library deprecation
# model = genai.GenerativeModel("gemini-1.5-flash")

def get_fallback_response(user_message):
    """Fallback response system for testing"""
    message_lower = user_message.lower()
    
    if "menu" in message_lower or "food" in message_lower:
        return "🍛 For today's menu and regular updates, please join our WhatsApp group: https://chat.whatsapp.com/EpLv7mtYipm61ScKjbOiuk. You'll get daily menu updates and can place orders directly there!"
    
    elif "timing" in message_lower or "time" in message_lower or "when" in message_lower:
        return "⏰ Order timings: Breakfast (order before 7:00 AM), Lunch (order before 9:45 AM), Dinner (order before 5:15 PM). We're closed on Sundays. Delivery: Breakfast 8-9 AM, Lunch 11 AM-1 PM, Dinner 7-9 PM."
    
    elif "delivery" in message_lower or "area" in message_lower or "location" in message_lower:
        return "📍 We deliver to Magarpatta, Amanora Township, Bhosale Garden, and Hadapsar areas within 3km radius. We offer both self-delivery and Zomato delivery!"
    
    elif "order" in message_lower:
        return "📞 To place an order, you can: 1) Join our WhatsApp group for daily menu and quick ordering: https://chat.whatsapp.com/EpLv7mtYipm61ScKjbOiuk, 2) Call us at 9930748908, or 3) Click the '📞 Place Order' button for our Tally form. WhatsApp group is fastest for daily menu!"
    
    else:
        return f"👋 Hello! Welcome to Svaadh Kitchen! We're a homemade vegetarian cloud kitchen in Hadapsar, Pune. For today's menu and regular updates, please join our WhatsApp group: https://chat.whatsapp.com/EpLv7mtYipm61ScKjbOiuk. How can I help you today?"

@app.post("/chat")
async def chat(request: Request):
    data = await request.json()
    user_message = data.get("message", "")

    # Use fallback response system for now
    bot_reply = get_fallback_response(user_message)
    
    return JSONResponse({"reply": bot_reply})
