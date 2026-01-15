
from dotenv import load_dotenv
import os

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET  = os.getenv("CLIENT_SECRET")
TENANT_ID      = os.getenv("TENANT_ID")



AUTHORITY = os.getenv("AUTHORITY")

GRAPH_SCOPE =  os.getenv("GRAPH_SCOPE")



# Power BI REST API (Workspaces, Reports, Dashboards)
POWERBI_SCOPE = os.getenv("POWERBI_SCOPE")
# Fabric API (Semantic Models, TMDL)
SCOPE =os.getenv("SCOPE")

BASE =os.getenv("BASE")
WORKSPACE_ID=os.getenv("WORKSPACE_ID")

#MODEL_ID = "2ba8e970-d12e-448a-8d35-a6fd966e2b56" #SCAR

