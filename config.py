import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

PORT = int(os.environ.get("PORT", 5001))
DEBUG = os.environ.get("RENDER") is None  # debug only when running locally
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-re-underwriting-key")
