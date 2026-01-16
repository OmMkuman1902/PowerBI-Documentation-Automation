from dotenv import load_dotenv
import os

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
WORKSPACE_ID = os.getenv("WORKSPACE_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
SCOPE = ["https://api.fabric.microsoft.com/.default"]

BASE = "https://api.fabric.microsoft.com/v1"


#MODEL_ID = "2ba8e970-d12e-448a-8d35-a6fd966e2b56" #SCAR

