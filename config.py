import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

PORT = 5001
DEBUG = True
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-re-underwriting-key")
