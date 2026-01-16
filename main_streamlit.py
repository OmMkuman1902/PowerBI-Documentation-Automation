import streamlit as st
import json
import os
import tempfile
import shutil
from datetime import datetime
from pipeline import build_semantic_model_with_sources
from Test_Main_Doc import MainDocumentation

from extraction import (
    get_token,
    get_report_details,
    get_report_pages_powerbi,
    get_workspace_info,
    get_workspace_users,
    get_workspace_items,
    get_semantic_model_metadata,
    get_owner_info,
    get_datasources,
    get_refresh_schedule,
    get_refresh_history,
    get_model_definition_lro,
    extract_tables_and_measures,
    build_semantic_models_index,  # Add this
    find_model_id_by_name,
)
import requests
import traceback

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Power BI Documentation Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)
# ---------- CUSTOM CSS ----------
st.markdown(
    """
    <style>
    /* ==============================
       PRIMARY BUTTON
    ============================== */
    div.stButton > button[kind="primary"] {
        background-color: #ff7a00 !important;
        color: black !important;
        border-radius: 8px;
        border: none;
        font-weight: 600;
    }

    div.stButton > button[kind="primary"]:hover {
        background-color: #e66f00 !important;
        color: black !important;
    }

    /* ==============================
       MAIN APP BACKGROUND (DARK)
    ============================== */
    .stApp {
        background-color: #0E1117 !important;
        color: #FAFAFA !important;
    }

    /* ==============================
       TEXT COLORS
    ============================== */
    html, body, [class*="css"] {
        color: #FAFAFA !important;
    }

    h1, h2, h3, h4, h5, h6 {
        color: #FAFAFA !important;
    }

    /* ==============================
       SIDEBAR
    ============================== */
    section[data-testid="stSidebar"] {
        background-color: #161B22 !important;
        color: #FAFAFA !important;
        border-right: 1px solid #2A2E35;
    }

    section[data-testid="stSidebar"] * {
        color: #FAFAFA !important;
    }

    /* ==============================
       INPUTS & SELECTBOX
    ============================== */
    input, textarea, select {
        background-color: #161B22 !important;
        color: #FAFAFA !important;
        border: 1px solid #30363D !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #161B22 !important;
        color: #FAFAFA !important;
    }

    /* ==============================
       CODE BLOCKS
    ============================== */
    pre, code {
        background-color: #161B22 !important;
        color: #FAFAFA !important;
        border-radius: 6px;
    }

    /* ==============================
       TABLES
    ============================== */
    thead tr th {
        background-color: #1F242C !important;
        color: #FAFAFA !important;
    }

    tbody tr td {
        color: #FAFAFA !important;
    }

    /* ==============================
       METRIC CARDS
    ============================== */
    div[data-testid="metric-container"] {
        background-color: #161B22 !important;
        color: #FAFAFA !important;
        border: 1px solid #30363D;
        border-radius: 8px;
        padding: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)




col_left, col_center, col_right = st.columns([2, 2, 1])
with col_center:
    st.image("image.png", width=275)






BASE = "https://api.fabric.microsoft.com/v1"
OUTPUT_FOLDER = r"C:\Temp\PowerBI_Docs"  # Changed to C:\Temp

# =========================================================
# SESSION STATE INITIALIZATION
# =========================================================
if 'workspace_loaded' not in st.session_state:
    st.session_state.workspace_loaded = False
if 'workspace_json' not in st.session_state:
    st.session_state.workspace_json = None
if 'reports_json' not in st.session_state:
    st.session_state.reports_json = None
if 'temp_folder' not in st.session_state:
    st.session_state.temp_folder = None
if 'semantic_json' not in st.session_state:
    st.session_state.semantic_json = None
if 'semantic_models_index' not in st.session_state:
    st.session_state.semantic_models_index = None

# =========================================================
# HELPER FUNCTIONS
# =========================================================
def create_output_folder():
   
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(OUTPUT_FOLDER, f"extraction_{timestamp}")
    os.makedirs(output_path, exist_ok=True)
    return output_path

def save_json_to_folder(folder_path, workspace_json=None, reports_json=None, semantic_json=None):
    """Save JSON files to the specified folder"""
    saved_files = []
    
    if workspace_json:
        file_path = os.path.join(folder_path, "1_workspace_info.json")
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(workspace_json, f, indent=2)
        saved_files.append(file_path)
    
    if reports_json:
        file_path = os.path.join(folder_path, "2_reports_info.json")
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(reports_json, f, indent=2)
        saved_files.append(file_path)
    
    if semantic_json:
        file_path = os.path.join(folder_path, "3_semantic_model_info.json")
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(semantic_json, f, indent=2)
        saved_files.append(file_path)
    
    return saved_files

# =========================================================
# HELPER FUNCTIONS
# =========================================================
def get_all_workspaces():
    """Fetch all workspaces the user has access to"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        return []

def get_reports_in_workspace(workspace_id):
    """Get all reports in a specific workspace"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"{BASE}/workspaces/{workspace_id}/reports"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        return []

def build_workspace_json(workspace_id):
    """Build complete workspace JSON"""
    workspace_data = {
        "workspace_metadata": get_workspace_info(),
        "workspace_users": get_workspace_users(),
        "workspace_items": get_workspace_items()
    }
    return workspace_data

def build_reports_json(workspace_id, progress_container):
    """Build complete reports JSON with all reports and their details"""
    reports_list = get_reports_in_workspace(workspace_id)
    
    if not reports_list:
        return {"reports": []}
    
    all_reports = []
    total_reports = len(reports_list)
    
    progress_bar = progress_container.progress(0)
    status_text = progress_container.empty()
    
    for i, report in enumerate(reports_list):
        report_id = report.get("id")
        report_name = report.get("displayName", "Unknown")
        
        status_text.text(f"Processing {i+1}/{total_reports}: {report_name}")
        progress_bar.progress((i + 1) / total_reports)
        
        try:
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
        except Exception as e:
            status_text.warning(f"⚠️ Could not extract details for {report_name}: {str(e)}")
            all_reports.append({
                "report_id": report_id,
                "report_name": report_name,
                "report_metadata": report,
                "report_details": None,
                "report_pages": None,
                "error": str(e)
            })
    
    status_text.empty()
    progress_bar.empty()
    
    return {"reports": all_reports}

def build_semantic_model_json(dataset_id):
    """Build complete semantic model JSON"""
    model_definition = get_model_definition_lro(dataset_id)
    parsed_data = extract_tables_and_measures(model_definition)
    
    metadata = get_semantic_model_metadata(dataset_id)
    owner_info = get_owner_info(dataset_id)
    datasources = get_datasources(dataset_id)
    refresh_schedule = get_refresh_schedule(dataset_id)
    refresh_history = get_refresh_history(dataset_id)
    
    semantic_model_data = {
        "semantic_model_id": dataset_id,
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

def create_temp_folder():
    """Create a temporary folder for this session"""
    temp_base = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_folder = os.path.join(temp_base, f"powerbi_extraction_{timestamp}")
    os.makedirs(temp_folder, exist_ok=True)
    return temp_folder

def save_json_files(temp_folder, workspace_json=None, reports_json=None, semantic_json=None):
    """Save JSON files to temp folder"""
    file_paths = []
    
    if workspace_json:
        workspace_file = os.path.join(temp_folder, "1_workspace_info.json")
        with open(workspace_file, 'w') as f:
            json.dump(workspace_json, f, indent=2)
        file_paths.append(workspace_file)
    
    if reports_json:
        report_file = os.path.join(temp_folder, "2_reports_info.json")
        with open(report_file, 'w') as f:
            json.dump(reports_json, f, indent=2)
        file_paths.append(report_file)
    
    if semantic_json:
        semantic_file = os.path.join(temp_folder, "3_semantic_model_info.json")
        with open(semantic_file, 'w') as f:
            json.dump(semantic_json, f, indent=2)
        file_paths.append(semantic_file)
    
    return file_paths

def create_zip_from_folder(temp_folder):
    """Create a zip file from the temp folder"""
    zip_path = f"{temp_folder}.zip"
    shutil.make_archive(temp_folder, 'zip', temp_folder)
    return zip_path

# =========================================================
# HEADER
# =========================================================
st.markdown(
    """
    <h1 style="text-align:center; margin-bottom:5px;">
        🚀 Power BI Documentation Automation
    </h1>
    <p style="text-align:center; color:gray;">
        Extract workspace, all reports, and semantic model information
    </p>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# =========================================================
# STEP 1: SELECT WORKSPACE
# =========================================================
st.markdown("### 🏢 Step 1: Select Workspace")

@st.cache_data(show_spinner="Loading workspaces…", ttl=300)
def load_workspaces():
    return get_all_workspaces()

with st.spinner("🔄 Loading workspaces..."):
    workspaces = load_workspaces()

if not workspaces:
    st.error("❌ No workspaces found or unable to connect to Power BI.")
    st.stop()

# Build Workspace ID → Name mapping
workspace_map = {
    ws["id"]: ws["displayName"]
    for ws in workspaces
}

# Sort alphabetically
workspace_map = dict(
    sorted(workspace_map.items(), key=lambda x: x[1].lower())
)

selected_workspace_name = st.selectbox(
    "Select Workspace",
    options=list(workspace_map.values()),
    key="workspace_selector",
    label_visibility="collapsed"
)

# Get workspace ID from selected name
selected_workspace_id = [k for k, v in workspace_map.items() if v == selected_workspace_name][0]

st.markdown("---")

# =========================================================
# STEP 2: LOAD WORKSPACE AND REPORTS DATA
# =========================================================
st.markdown("### 📊 Step 2: Load Workspace and Reports Data")

if st.button("🔄 Load Workspace and Reports Data", type="primary", use_container_width=True):
    
    # Create temp folder
    if not st.session_state.temp_folder:
        st.session_state.temp_folder = create_temp_folder()
    
    temp_folder = st.session_state.temp_folder
    
    with st.spinner(""):
        try:
            # Extract Workspace Info
            st.info("📁 Extracting workspace information...")
            workspace_json = build_workspace_json(selected_workspace_id)
            st.session_state.workspace_json = workspace_json
            st.success("✅ Workspace info extracted successfully")
            
            st.markdown("---")
            
            # Extract Reports Info
            st.info("📊 Extracting reports information...")
            progress_container = st.container()
            reports_json = build_reports_json(selected_workspace_id, progress_container)
            st.session_state.reports_json = reports_json
            st.success("✅ Reports info extracted successfully")
            
            st.markdown("---")
            
            # Extract Semantic Models Index - THIS WAS MISSING!
            st.info("🧠 Extracting semantic models index...")
            semantic_models_index = build_semantic_models_index(selected_workspace_id)
            st.session_state.semantic_models_index = semantic_models_index
            st.success(f"✅ Found {len(semantic_models_index.get('semantic_models', []))} semantic models")
            
            # Save files
            save_json_files(temp_folder, workspace_json, reports_json, None)
            
            # Also save semantic models index
            semantic_index_file = os.path.join(temp_folder, "semantic_models_index.json")
            with open(semantic_index_file, 'w') as f:
                json.dump(semantic_models_index, f, indent=2)
            
            st.session_state.workspace_loaded = True
            
        except Exception as e:
            st.error(f"❌ Error during extraction: {str(e)}")
            st.code(traceback.format_exc())

st.markdown("---")

# =========================================================
# STEP 3: SHOW STATISTICS (AFTER LOADING)
# =========================================================
if st.session_state.workspace_loaded:
    
    

    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.session_state.workspace_json and st.session_state.workspace_json.get("workspace_items"):
            items = st.session_state.workspace_json["workspace_items"].get("value", [])
            st.metric("📁 Workspace Items", len(items))
    
    with col2:
        if st.session_state.reports_json and st.session_state.reports_json.get("reports"):
            st.metric("📊 Total Reports", len(st.session_state.reports_json["reports"]))
    
    with col3:
        st.metric("📂 Files Generated", "2 / 3")
        st.caption("1_workspace_info.json  \n2_reports_info.json")
    
    st.markdown("---")
    
    # =========================================================
    # STEP 4: SELECT REPORT AND GENERATE SEMANTIC MODEL
    # =========================================================
    st.markdown("### 🗄️ Step 3: Generate Semantic Model")

    # Get reports list from session
    reports_list = st.session_state.reports_json.get("reports", [])

    if reports_list:
        # Build mapping: report_name -> report object
        report_map = {r["report_name"]: r for r in reports_list}

        # Streamlit selectbox for report names
        selected_report_name = st.selectbox(
            "Select Report",
            options=list(report_map.keys()),
            index=0,
            key="report_selector"
        )

        # Get the full report object
        selected_report = report_map.get(selected_report_name)
        selected_report_id = selected_report.get("report_id") if selected_report else None
        
        print("Selected report name =", selected_report_name)
        print("Selected report id =", selected_report_id)

        # Find matching semantic model by name
        matching_model_id = None
        if st.session_state.semantic_models_index:
            matching_model_id = find_model_id_by_name(
                st.session_state.semantic_models_index, 
                selected_report_name
            )
            
        if matching_model_id:
            st.info(f"📋 Semantic Model ID: `{matching_model_id}`")
            # 🔥 IMPORTANT: write into session_state BEFORE text_input
            st.session_state.model_input = matching_model_id
        else:
            st.warning(
                f"⚠️ No semantic model found matching report name: '{selected_report_name}'"
                )
            st.session_state.model_input = ""

            st.markdown("---")
                
        # Show all semantic models as reference
        if st.session_state.semantic_models_index:
            with st.expander("📊 View All Semantic Models in Workspace", expanded=False):
                semantic_models = st.session_state.semantic_models_index.get("semantic_models", [])
                
                if semantic_models:
                    st.markdown("**Available Semantic Models:**")
                    for model in semantic_models:
                        st.code(f"Name: {model.get('model_name', 'Unknown')}\nID: {model.get('model_id', 'Unknown')}")
                else:
                    st.info("No semantic models found.")
        
        st.markdown("---")
        
        # Manual override option
        st.markdown("**Or Enter Model ID Manually (Optional):**")
        if matching_model_id:
            st.session_state.model_input = matching_model_id

        model_id_input = st.text_input(
            "Semantic Model ID",
            value="model_input",
            placeholder="Model ID will be auto-filled if found...",
            key="model_input",
            help="Auto-filled based on report name match. You can override if needed."
        )

        final_model_id=model_id_input.strip()
        
        st.markdown("---")

        # Generate Semantic Model Button
    if st.button("🚀 Generate Semantic Model Data", type="primary", use_container_width=True):

        if not final_model_id:
            st.error("❌ No semantic model ID found or entered.")
            st.stop()

        try:
            with st.spinner("🗄️ Extracting semantic model information..."):
                st.info(f"📋 Using Semantic Model ID: `{final_model_id}`")

                # ✅ BASIC semantic model (FILE 3)
                semantic_json = build_semantic_model_json(final_model_id)

                st.session_state.semantic_json = semantic_json

                save_json_files(
                    st.session_state.temp_folder,
                    None,
                    None,
                    semantic_json
                )

                st.success("✅ 3_semantic_model_info.json generated")

                col1, col2, col3 = st.columns(3)
                col1.metric("Tables", len(semantic_json.get("tables", [])))
                col2.metric("Measures", len(semantic_json.get("measures", [])))
                col3.metric("Relationships", len(semantic_json.get("relationships", [])))

            # -----------------------------------------
            # STEP 2: BUILD FILE 4 FROM FILE 3
            # -----------------------------------------
            with st.spinner("🧠 Adding table sources & metadata..."):

                file_path, enriched_json = build_semantic_model_with_sources(
                    MODEL_ID=final_model_id,
                    base_semantic_json=semantic_json,  # 👈 IMPORTANT
                    output_folder=st.session_state.temp_folder
                )

                st.session_state.semantic_json_enriched = enriched_json
                st.session_state.semantic_json_path = file_path

                st.success("✅ 4_semantic_model_with_table_sources.json generated")
                st.info(f"📄 File created: `{file_path}`")

        except Exception:
            st.error("❌ Semantic model generation failed")
            st.code(traceback.format_exc())


        

    #     # Generate Semantic Model Button
    #     if st.button("🚀 Generate Semantic Model Data", type="primary", use_container_width=True):
    #         if not final_dataset_id:
    #             st.error("❌ Please select a report with a dataset ID or enter one manually.")
    #         else:
    #             with st.spinner("🗄️ Extracting semantic model information..."):
    #                 try:
    #                     semantic_json = build_semantic_model_json(selected_workspace_id, final_dataset_id)
    #                     st.session_state.semantic_json = semantic_json

    #                     # Save semantic model file
    #                     save_json_files(st.session_state.temp_folder, None, None, semantic_json)

    #                     st.success("✅ Semantic model info extracted successfully")
    #                 except Exception as e:
    #                     st.error(f"❌ Error extracting semantic model: {str(e)}")
    #                     st.code(traceback.format_exc())

    #     st.markdown("---")
    # else:
    #     st.warning("⚠️ No reports found. Please generate reports first.")
    
    # =========================================================
    # DOWNLOAD SECTION
    # =========================================================
    HALF_RESET_KEYS = [
        "selected_report_name",
        "model_input",
        "semantic_json",
        "selected_dataset_id",
    ]

    FULL_RESET_KEYS = [
        "workspace_loaded",
        "workspace_json",
        "reports_json",
        "semantic_models_index",
        "temp_folder",
        "selected_workspace_id",
        "selected_report_name",
        "model_input",
        "semantic_json",
        "selected_dataset_id",
    ]


    st.markdown("### 💾 Save Documentation to Local Folder")

    if st.button("💾 Save All Files to Local Folder", type="primary", use_container_width=True):

        try:
            output_folder = create_output_folder()

            saved_files = save_json_to_folder(
                folder_path=output_folder,
                workspace_json=st.session_state.workspace_json,
                reports_json=st.session_state.reports_json,
                semantic_json=st.session_state.semantic_json
            )

            # 👉 COPY 4th FILE
            if "semantic_json_path" in st.session_state:
                src = st.session_state.semantic_json_path
                dst = os.path.join(output_folder, "4_semantic_model_with_table_sources.json")
                shutil.copy(src, dst)
                saved_files.append(dst)

            st.success("✅ All files saved successfully!")
            st.info(f"📁 Saved to: `{output_folder}`")

            with st.expander("📄 Saved Files"):
                for f in saved_files:
                    st.code(f)

        except Exception:
            st.error("❌ Failed to save files")
            st.code(traceback.format_exc())
    st.markdown("### 📄 Generate Documentation (Word)")

    if st.button("📄 Generate Dashboard Documentation", type="primary", use_container_width=True):

        if "semantic_json_path" not in st.session_state:
            st.error("❌ Semantic model file not found. Generate semantic model first.")
            st.stop()

        try:
            with st.spinner("📝 Generating documentation..."):

                extraction_folder = os.path.dirname(st.session_state.semantic_json_path)


                # ----------------------------
                # TEMPLATE PATH (FIXED)
                # ----------------------------
                template_path = r"C:\Temp\Dashboarding - FDD Template.docx"

                if not os.path.exists(template_path):
                    st.error(f"❌ Template not found at: {template_path}")
                    st.stop()

                # ----------------------------
                # OUTPUT DOC PATH (SAME FOLDER)
                # ----------------------------
                dashboard_name = (
                    st.session_state.semantic_json
                    .get("semantic_model_metadata", {})
                    .get("displayName", "Dashboard")
                    .replace(" ", "_")
                )

                output_doc_path = os.path.join(
                    extraction_folder,
                    f"{dashboard_name}_Documentation.docx"
                )

                # ----------------------------
                # BUILD DOCUMENT
                # ----------------------------
                docs = MainDocumentation(
                    semantic_json_path=st.session_state.semantic_json_path,
                    template_path=template_path,
                    output_doc_path=output_doc_path
                )

                docs.build()

                st.success("✅ Documentation generated successfully!")
                st.info(f"📄 Saved at: `{output_doc_path}`")

        except Exception:
            st.error("❌ Failed to generate documentation")
            st.code(traceback.format_exc())


# =========================================================
# SIDEBAR INFO
# =========================================================
with st.sidebar:
    st.markdown("### ℹ️ Workflow")
    st.markdown(
        """
        **Step 1:** Select Workspace
        - Choose from available workspaces
        
        **Step 2:** Load Data
        - Click "Load Workspace and Reports Data"
        - Extracts workspace info
        - Extracts ALL reports (with progress bar)
        - Generates 2 JSON files
        
        **Step 3:** Generate Semantic Model
        - Select a report (auto-fills dataset ID)
        - Or enter dataset ID manually
        - Click "Generate Semantic Model Data"
        - Generates 3rd JSON file
        
        **Step 4:** Download
        - Download individual files
        - Or download complete ZIP package
        """
    )
    
    st.markdown("---")
    st.markdown("### 📁 Output Files")
    st.markdown(
        """
        1. `1_workspace_info.json`
        2. `2_reports_info.json`
        3. `3_semantic_model_info.json`
        """
    )
    
    st.markdown("---")
    
    # Show current status
    st.markdown("### 📊 Current Status")
    if st.session_state.workspace_loaded:
        st.success("✅ Workspace & Reports Loaded")
        if st.session_state.semantic_json:
            st.success("✅ Semantic Model Generated")
        else:
            st.info("⏳ Semantic Model Pending")
    else:
        st.info("⏳ Ready to Load Data")
    
    st.markdown("---")
    
    if st.button("🗑️ Clear Cache", use_container_width=True):
        st.cache_data.clear()
        st.success("Cache cleared!")
        st.rerun()