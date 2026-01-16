import requests
import msal
import time
from  dotenv import load_dotenv
load_dotenv()
import json
import base64
import re,os
from config import TENANT_ID,CLIENT_ID,CLIENT_SECRET#,AUTHORITY,SCOPE,POWERBI_SCOPE,BASE,WORKSPACE_ID


AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
#SCOPE = ["https://api.fabric.microsoft.com/.default"]

BASE = "https://api.fabric.microsoft.com/v1"
WORKSPACE_ID = "bdd4cd59-3a78-414a-8fac-813eb25ad20e"


# ---------------------------------------------------------
# AUTH
# ---------------------------------------------------------
def get_token(scope=None):
    if scope is None:
        scope = ["https://api.fabric.microsoft.com/.default"]
    
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
,
    )
    tok = app.acquire_token_for_client(scopes=scope)
    if "access_token" not in tok:
        raise Exception("TOKEN ERROR", tok)
    return tok["access_token"]




# =================================================================
# WORKSPACE INFO EXTRACTION
# =================================================================
def get_reports():
    """Get all reports from the fixed workspace"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    url = f"{BASE}/workspaces/{WORKSPACE_ID}/reports"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        return []

def get_workspace_info():
    """Get workspace metadata"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}"
    
    print("📁 Fetching workspace info...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"  ⚠️  Error: {response.status_code}")
        return None


def get_workspace_users():
    """Get workspace users and permissions"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/roleAssignments"
    
    print("👥 Fetching workspace users...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"  ⚠️  Error: {response.status_code}")
        return None


def get_workspace_items():
    """Get all items in workspace (reports, datasets, etc.)"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/items"
    
    print("📦 Fetching workspace items...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"  ⚠️  Error: {response.status_code}")
        return None


def build_workspace_json():
    """Build complete workspace JSON"""
    workspace_data = {
        "workspace_metadata": get_workspace_info(),
        "workspace_users": get_workspace_users(),
        "workspace_items": get_workspace_items()
    }
    
    return workspace_data


# =================================================================
# REPORT INFO EXTRACTION
# =================================================================
def get_reports_in_workspace():
    """Get all reports in workspace"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/reports"
    
    print("\n📊 Fetching reports in workspace...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"  ⚠️  Error: {response.status_code}")
        return None


def get_report_details(report_id):
    """Get detailed info for a specific report"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/reports/{report_id}"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        return None


def get_report_pages_powerbi(report_id):
    """Get report pages using Power BI REST API"""
    token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/reports/{report_id}/pages"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        return None


def build_reports_json():
    """Build complete reports JSON with all reports and their details"""
    reports_list = get_reports_in_workspace()
    
    if not reports_list:
        return {"reports": []}
    
    all_reports = []
    
    reports = reports_list.get("value", [])
    print(f"\n📄 Found {len(reports)} reports")
    
    for report in reports:
        report_id = report.get("id")
        report_name = report.get("displayName", "Unknown")
        
        print(f"  Processing: {report_name}")
        
        # Get detailed info
        details = get_report_details(report_id)
        
        # Get pages
        pages = get_report_pages_powerbi(report_id)
        
        report_data = {
            "report_id": report_id,
            "report_name": report_name,
            "report_metadata": report,
            "report_details": details,
            "report_pages": pages
        }
        
        all_reports.append(report_data)
    
    return {"reports": all_reports}
#############################################################################################
def fetch_semantic_models(workspace_id):
    """Fetch all semantic models in a workspace"""
    token = get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = f"{BASE}/workspaces/{workspace_id}/semanticModels"

    print("🧠 Fetching semantic models...")
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"  ⚠️  Error: {response.status_code}")
        return None

    return response.json()


def build_semantic_models_index(workspace_id):
    """Build an index of all semantic models with their IDs"""
    response_json = fetch_semantic_models(workspace_id)

    if not response_json:
        return {"workspace_id": workspace_id, "semantic_models": []}

    models = response_json.get("value", [])

    semantic_models = []

    for model in models:
        semantic_models.append({
            "model_id": model.get("id"),
            "model_name": model.get("displayName"),
            "model_metadata": {
                "id": model.get("id"),
                "displayName": model.get("displayName"),
                "type": model.get("type"),
                "workspaceId": workspace_id
            }
        })

    print(f"✅ Found {len(semantic_models)} semantic models")
    
    return {
        "workspace_id": workspace_id,
        "semantic_models": semantic_models
    }


def find_model_id_by_name(semantic_models_index, report_name):
    """Find model ID by matching report name with model name"""
    semantic_models = semantic_models_index.get("semantic_models", [])
    
    # Try exact match first
    for model in semantic_models:
        if model.get("model_name") == report_name:
            print(f"✅ Exact match found: {model.get('model_name')} -> {model.get('model_id')}")
            return model.get("model_id")
    
    # Try case-insensitive match
    report_name_lower = report_name.lower()
    for model in semantic_models:
        if model.get("model_name", "").lower() == report_name_lower:
            print(f"✅ Case-insensitive match found: {model.get('model_name')} -> {model.get('model_id')}")
            return model.get("model_id")
    
    # Try partial match
    for model in semantic_models:
        model_name = model.get("model_name", "")
        if report_name in model_name or model_name in report_name:
            print(f"⚠️  Partial match found: {model.get('model_name')} -> {model.get('model_id')}")
            return model.get("model_id")
    
    print(f"❌ No matching model found for report: {report_name}")
    return None
# =================================================================
# SEMANTIC MODEL INFO EXTRACTION (Your existing logic)
# =================================================================
def get_model_definition_lro(MODEL_ID):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    start_url = f"{BASE}/workspaces/{WORKSPACE_ID}/semanticModels/{MODEL_ID}/getDefinition"
    payload = {"includeDynamicFormatStrings": True, "includeTmdl": True}

    print("\n🔍 Starting semantic model extraction...")
    start_response = requests.post(start_url, headers=headers, json=payload)

    if start_response.status_code != 202:
        print("Error:", start_response.status_code)
        print(start_response.text)
        return None

    poll_url = start_response.headers["Location"]
    print("⏳ Polling operation...")

    while True:
        time.sleep(2)
        poll_res = requests.get(poll_url, headers=headers)
        body = poll_res.json()
        status = body.get("status")

        print("  Status:", status)

        if status == "Succeeded":
            break

    final_url = poll_res.headers["Location"]
    return requests.get(final_url, headers=headers).json()


def get_semantic_model_metadata(MODEL_ID):
    """Get semantic model metadata including owner"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/semanticModels/{MODEL_ID}"
    
    print("👤 Fetching semantic model metadata...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        return None


def get_dataset_users_powerbi(MODEL_ID):
    """Get dataset users/permissions using Power BI REST API"""
    token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{MODEL_ID}/users"
    
    print("👥 Fetching dataset users/permissions...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"  ⚠️  Could not fetch dataset users: {response.status_code}")
        return None


def get_owner_info(MODEL_ID):
    """
    Get owner/configured by information from multiple sources
    Returns consolidated owner information
    """
    print("🔍 Fetching owner information...")
    
    owner_info = {
        "configured_by": None,
        "created_by": None,
        "last_modified_by": None,
        "dataset_users": None,
        "owner_details": None
    }
    
    # Try Fabric API metadata
    metadata = get_semantic_model_metadata(MODEL_ID)
    if metadata:
        owner_info["configured_by"] = metadata.get("configuredBy")
        owner_info["created_by"] = metadata.get("createdBy")
        owner_info["last_modified_by"] = metadata.get("lastModifiedBy")
        
        # Some additional owner fields that might be present
        if "properties" in metadata:
            props = metadata.get("properties", {})
            owner_info["owner_details"] = {
                "owner": props.get("owner"),
                "configuredByLegacy": props.get("configuredByLegacy")
            }
    
    # Try Power BI API for dataset users
    dataset_users = get_dataset_users_powerbi(MODEL_ID)
    if dataset_users:
        owner_info["dataset_users"] = dataset_users
    
    # Try to get from Power BI Groups API (additional method)
    token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
    headers = {"Authorization": f"Bearer {token}"}
    
    powerbi_url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{MODEL_ID}"
    response = requests.get(powerbi_url, headers=headers)
    
    if response.status_code == 200:
        pbi_metadata = response.json()
        owner_info["powerbi_configured_by"] = pbi_metadata.get("configuredBy")
        owner_info["powerbi_created_by"] = pbi_metadata.get("createdBy")
        
    print("✅ Owner information collected")
    
    return owner_info


def get_datasources(MODEL_ID):
    """Get datasources for the semantic model"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{WORKSPACE_ID}/semanticModels/{MODEL_ID}/datasources"
    
    print("🔌 Fetching datasources...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        # Fallback to Power BI API
        token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
        headers = {"Authorization": f"Bearer {token}"}
        powerbi_url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{MODEL_ID}/datasources"
        response = requests.get(powerbi_url, headers=headers)
        return response.json() if response.status_code == 200 else None


def get_refresh_schedule(MODEL_ID):
    """Get refresh schedule using Power BI REST API"""
    token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{MODEL_ID}/refreshSchedule"
    
    print("📅 Fetching refresh schedule...")
    response = requests.get(url, headers=headers)
    
    return response.json() if response.status_code == 200 else None


def get_refresh_history(MODEL_ID):
    """Get refresh history using Power BI REST API"""
    token = get_token(["https://analysis.windows.net/powerbi/api/.default"])
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{MODEL_ID}/refreshes?$top=10"
    
    print("🔄 Fetching refresh history...")
    response = requests.get(url, headers=headers)
    
    return response.json() if response.status_code == 200 else None


def extract_measures_universal(content, table_name):
    measures = []

    pattern = r"""
        measure                              
        \s+
        ['"]?([^'"\n=]+)['"]?                
        \s*=\s*                               
        (                                     
            (?:
                (?!lineageTag:|annotation\s|measure\s)   
                .|\n
            )*?
        )                                     
        (?=\n\s*(?:lineageTag:|annotation\s|measure\s|$))  
    """

    matches = re.findall(pattern, content, re.VERBOSE | re.IGNORECASE)

    for name, expr in matches:
        measures.append({
            "name": name.strip(),
            "expression": expr.strip(),
            "table": table_name
        })

    return measures

def parse_tmdl_table(content, table_name):
    table_info = {
        "name": table_name,
        "columns": [],
        "measures": [],
    }

    # Enhanced column pattern to capture dataType
    col_pattern = r"""
    column\s+
    (
        '(.*?)'        |   # single-quoted
        "(.*?)"        |   # double-quoted
        \[(.*?)\]      |   # bracketed
        ([^\s=]+)          # unquoted (no spaces)
    )
    (?:\s*\n\s*dataType:\s*(\w+))?  # optional dataType on next line
    """

    matches = re.findall(col_pattern, content, re.VERBOSE | re.DOTALL)

    for m in matches:
        col_name = next(filter(None, m[:-1]))  # pick the non-empty group (excluding dataType)
        data_type = m[-1] if m[-1] else "string"  # default to string if not found
        
        table_info["columns"].append({
            "name": col_name.strip(),
            "sourceColumn": col_name.strip(),
            "dataType": data_type.strip() if data_type else "string"
        })

    table_info["measures"] = extract_measures_universal(content, table_name)

    return table_info

# def parse_tmdl_table(content, table_name):
#     table_info = {
#         "name": table_name,
#         "columns": [],
#         "measures": [],
#     }

#     col_pattern = r"""
#     column\s+
#     (
#         '(.*?)'        |   # single-quoted
#         "(.*?)"        |   # double-quoted
#         \[(.*?)\]      |   # bracketed
#         ([^\s=]+)          # unquoted (no spaces)
#     )
#     """

#     matches = re.findall(col_pattern, content, re.VERBOSE)

#     for m in matches:
#         col_name = next(filter(None, m))  # pick the non-empty group
#         table_info["columns"].append({
#             "name": col_name.strip(),
#             "sourceColumn": col_name.strip()
#         })

#     table_info["measures"] = extract_measures_universal(content, table_name)

#     return table_info


# def parse_tmdl_table(content, table_name):
#     table_info = {
#         "name": table_name,
#         "columns": [],
#         "measures": [],
#     }

#     col_pattern = r"column\s+([^\s=]+)"
#     cols = re.findall(col_pattern, content)
#     for c in cols:
#         table_info["columns"].append({"name": c.strip(), "sourceColumn": c.strip()})

#     table_info["measures"] = extract_measures_universal(content, table_name)

#     return table_info


def parse_tmdl_relationship(content, filename):
    relationship_info = {
        "name": filename.replace(".tmdl", ""),
        "fromTable": None,
        "fromColumn": None,
        "toTable": None,
        "toColumn": None,
        "cardinality": None,
        "crossFilteringBehavior": None,
        "isActive": True
    }

    m1 = re.search(r"fromColumn:\s*'?(.*?)'?\.(\w+)", content)
    if m1:
        relationship_info["fromTable"] = m1.group(1).strip()
        relationship_info["fromColumn"] = m1.group(2).strip()

    m2 = re.search(r"toColumn:\s*'?(.*?)'?\.(\w+)", content)
    if m2:
        relationship_info["toTable"] = m2.group(1).strip()
        relationship_info["toColumn"] = m2.group(2).strip()

    c = re.search(r"cardinality:\s*(\w+)", content)
    if c:
        relationship_info["cardinality"] = c.group(1)

    cf = re.search(r"crossFilteringBehavior:\s*(\w+)", content)
    if cf:
        relationship_info["crossFilteringBehavior"] = cf.group(1)

    a = re.search(r"isActive:\s*(true|false)", content, re.IGNORECASE)
    if a:
        relationship_info["isActive"] = a.group(1).lower() == "true"

    return relationship_info


def extract_tables_and_measures(model_definition):
    result = {
        "tables": [],
        "measures": [],
        "relationships": [],
        "all_tmdl_files": {},
        "skipped_files": [],  # Track skipped files
    }

    parts = model_definition.get("definition", {}).get("parts", [])
    print(f"\n📦 Found {len(parts)} TMDL parts\n")

    for part in parts:
        path = part.get("path")
        payload = part.get("payload")

        # ✅ store path + payload together
        result["all_tmdl_files"][path] = {
            "payload": payload,
            "payloadType": part.get("payloadType", "InlineBase64")
        }

        # Decode payload
        try:
            content = base64.b64decode(payload).decode("utf-8")
        except UnicodeDecodeError:
            print(f"  ⚠️  Skipping binary file: {path}")
            result["skipped_files"].append({"path": path, "reason": "binary"})
            continue
        except Exception as e:
            print(f"  ⚠️  Error decoding {path}: {str(e)}")
            result["skipped_files"].append({"path": path, "reason": str(e)})
            continue


        if path.startswith("definition/tables/"):
            tname = path.split("/")[-1].replace(".tmdl", "")
            print("  📊 Table:", tname)
            table_info = parse_tmdl_table(content, tname)
            result["tables"].append(table_info)

            for m in table_info["measures"]:
                result["measures"].append(m)

        elif path.startswith("definition/relationships/") or path == "definition/relationships.tmdl":
            filename = path.split("/")[-1]
            print("  🔗 Relationship:", filename)

            blocks = re.split(r"(?=relationship\s)", content)

            for block in blocks:
                if block.strip().startswith("relationship"):
                    name_match = re.search(r"relationship\s+([^\s]+)", block)
                    rname = name_match.group(1) if name_match else filename

                    rel_info = parse_tmdl_relationship(block, rname)
                    result["relationships"].append(rel_info)

    print(f"\n✅ Extraction complete:")
    print(f"   • Tables: {len(result['tables'])}")
    print(f"   • Measures: {len(result['measures'])}")
    print(f"   • Relationships: {len(result['relationships'])}")
    print(f"   • Skipped files: {len(result['skipped_files'])}")

    return result


def build_semantic_model_json(MODEL_ID):
    """Build complete semantic model JSON"""
    
    # Get model definition and parse TMDL
    model_definition = get_model_definition_lro()
    parsed_data = extract_tables_and_measures(model_definition)
    
    # Get additional metadata
    metadata = get_semantic_model_metadata()
    owner_info = get_owner_info()
    datasources = get_datasources()
    refresh_schedule = get_refresh_schedule()
    refresh_history = get_refresh_history()
    
    semantic_model_data = {
        "semantic_model_id": MODEL_ID,
        "semantic_model_metadata": metadata,
        "owner_information": owner_info,
        "tables": parsed_data["tables"],
        "measures": parsed_data["measures"],
        "relationships": parsed_data["relationships"],
        "datasources": datasources,
        "refresh_schedule": refresh_schedule,
        "refresh_history": refresh_history,
        "tmdl_files": parsed_data["all_tmdl_files"]
    }
    
    return semantic_model_data


# =================================================================
# MAIN EXECUTION
# =================================================================
# if __name__ == "__main__":
#     print("\n" + "="*70)
#     print("🚀 POWER BI COMPLETE EXTRACTOR - 3 JSON FILES")
#     print("="*70 + "\n")

#     try:
#         # 1. Extract Workspace Info
#         print("\n" + "="*70)
#         print("📁 EXTRACTING WORKSPACE INFO")
#         print("="*70)
#         workspace_json = build_workspace_json()
        
#         with open("1_workspace_info.json", "w") as f:
#             json.dump(workspace_json, f, indent=2)
#         print("\n✅ Workspace info saved to: 1_workspace_info.json")
        
#         # 2. Extract Report Info
#         print("\n" + "="*70)
#         print("📊 EXTRACTING REPORT INFO")
#         print("="*70)
#         reports_json = build_reports_json()
        
#         with open("2_reports_info.json", "w") as f:
#             json.dump(reports_json, f, indent=2)
#         print("\n✅ Reports info saved to: 2_reports_info.json")
        
#         # 3. Extract Semantic Model Info
#         print("\n" + "="*70)
#         print("🗄️  EXTRACTING SEMANTIC MODEL INFO")
#         print("="*70)
#         semantic_model_json = build_semantic_model_json()
        
#         with open("3_semantic_model_info.json", "w") as f:
#             json.dump(semantic_model_json, f, indent=2)
#         print("\n✅ Semantic model info saved to: 3_semantic_model_info.json")
        
#         # Summary
#         print("\n" + "="*70)
#         print("🎉 EXTRACTION COMPLETE!")
#         print("="*70)
#         print("\n📦 Generated Files:")
#         print("  1️⃣  1_workspace_info.json")
#         print("     - Workspace metadata")
#         print("     - Workspace users & permissions")
#         print("     - All workspace items")
#         print("\n  2️⃣  2_reports_info.json")
#         print("     - All reports in workspace")
#         print("     - Report details & metadata")
#         print("     - Report pages")
#         print("\n  3️⃣  3_semantic_model_info.json")
#         print("     - Semantic model metadata")
#         print("     - Owner/Configured by information")
#         print("     - Tables, columns, measures")
#         print("     - Relationships")
#         print("     - Datasources")
#         print("     - Refresh schedule & history")
        
#         # Print counts
#         if workspace_json.get("workspace_items"):
#             items = workspace_json["workspace_items"].get("value", [])
#             print(f"\n📊 Statistics:")
#             print(f"  • Workspace items: {len(items)}")
        
#         if reports_json.get("reports"):
#             print(f"  • Reports: {len(reports_json['reports'])}")
        
#         if semantic_model_json.get("tables"):
#             print(f"  • Tables: {len(semantic_model_json['tables'])}")
#             print(f"  • Measures: {len(semantic_model_json['measures'])}")
#             print(f"  • Relationships: {len(semantic_model_json['relationships'])}")
        
#         # Print owner info if available
#         if semantic_model_json.get("owner_information"):
#             owner = semantic_model_json["owner_information"]
#             print(f"\n👤 Owner Information:")
#             if owner.get("configured_by"):
#                 print(f"  • Configured by: {owner['configured_by']}")
#             if owner.get("created_by"):
#                 print(f"  • Created by: {owner['created_by']}")
#             if owner.get("powerbi_configured_by"):
#                 print(f"  • Power BI configured by: {owner['powerbi_configured_by']}")
        
#         print("\n✨ All done!\n")
        
#     except Exception as e:
#         print(f"\n❌ ERROR: {e}")
#         import traceback
#         traceback.print_exc()

# if __name__=="__main__":
#         print("\n" + "="*70)
#         print("🗄️  EXTRACTING SEMANTIC MODEL INFO")
#         print("="*70)
#         semantic_model_json = build_semantic_model_json(MODEL_ID)
        
#         with open("3_semantic_model_info.json", "w") as f:
#             json.dump(semantic_model_json, f, indent=2)
#         print("\n✅ Semantic model info saved to: 3_semantic_model_info.json")
# if __name__ == "__main__":
#     print("WORKSPACE_ID:", WORKSPACE_ID)
#     print("Type:", type(WORKSPACE_ID))
#     print(TENANT_ID)