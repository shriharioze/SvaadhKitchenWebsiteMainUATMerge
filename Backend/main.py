from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

# Load .env (your GEMINI_API_KEY should be set here)
load_dotenv()

# Configure Gemini API
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Load business data
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE_DIR, "business.json"), "r", encoding="utf-8") as f:
    business_data = json.load(f)
    
    
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or restrict to your website domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

model = genai.GenerativeModel("gemini-2.0-flash")

@app.post("/chat")
async def chat(request: Request):
    data = await request.json()
    user_message = data.get("message", "")

    # Construct context prompt using business data
    system_prompt = f"""
You are an assistant for {business_data['name']}, a homemade vegetarian cloud kitchen in Pune.
About: {business_data['about']['summary']}
Details:
- About (description): {business_data['about']['description']}
- Order Timings: Breakfast {business_data['order timings']['breakfast']}, Lunch {business_data['order timings']['lunch']}, Dinner {business_data['order timings']['dinner']}, Closed on {business_data['order timings']['closed_on']}
- Delivery: Self Delivery ({business_data['delivery']['self_delivery']}), Zomato ({business_data['delivery']['zomato']}), within {business_data['delivery']['radius_km']} km radius
- Menu: {', '.join([item['name'] + ' (₹' + str(item['price']) + ')' 
                      for meal in business_data['menu']['meal_types'] 
                      for item in meal['items']])}
- Customization options: {', '.join(business_data['menu']['customization'])}
- Contact: {business_data['contact']['phone']['primary']}, Email: {business_data['contact']['email']}, WhatsApp Group: {business_data['contact']['whatsapp_group']}
- Locations Served: {', '.join(business_data['locations_served'])}

Behave like a friendly local kitchen staff replying to customers’ messages.
Keep it short, natural, and helpful.
"""

    # Generate response from Gemini
    response = model.generate_content(f"{system_prompt}\nCustomer: {user_message}\nKitchen Assistant:")

    bot_reply = response.text.strip()

    return JSONResponse({"reply": bot_reply})
