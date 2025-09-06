# app_asgi.py

# Step 1: Import your existing Flask app
from app import app  # <-- your Flask app from app.py

# Step 2: Import the WsgiToAsgi adapter
from asgiref.wsgi import WsgiToAsgi

# Step 3: Wrap your Flask app for ASGI
asgi_app = WsgiToAsgi(app)
