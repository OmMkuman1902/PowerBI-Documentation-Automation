from extraction import *
import os
import json
import base64
import re

DB_REGEX = re.compile(r'\[Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"Database"\]', re.IGNORECASE)
SCHEMA_REGEX = re.compile(r'\[Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"Schema"\]', re.IGNORECASE)
TABLE_REGEX = re.compile(r'\[Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"Table"\]', re.IGNORECASE)

def normalize_semantic_json(parsed_data: dict) -> dict:
    """
    Normalize extracted semantic model JSON into a stable schema
    consumed by Word documentation.
    """
    normalized = {}

    # --------------------------------------------------
    # Dashboard / Semantic model metadata
    # --------------------------------------------------
    # FIX: Use the correct key from parsed_data
    normalized["semantic_model_metadata"] = {
        "displayName": parsed_data.get("displayName", "Unknown Dashboard")
    }

    # --------------------------------------------------
    # Core model elements (pass-through)
    # --------------------------------------------------
    normalized["tables"] = parsed_data.get("tables", [])
    normalized["relationships"] = parsed_data.get("relationships", [])
    normalized["measures"] = parsed_data.get("measures", [])

    return normalized
class TMDLSourceExtractor:
    """
    Extract table-level datasource information from TMDL table files.
    SIMPLE, SAFE, NO POLLING, NO OVER-REGEX.
    """

    IGNORE_TABLE_PREFIXES = (
        "DateTableTemplate_",
        "LocalDateTable_"
    )

    def __init__(self, model_definition):
        self.model_definition = model_definition

    # --------------------------------------------------
    # Databricks Data Product extractor (tbl_*)
    # --------------------------------------------------
    def _extract_databricks_dp(self, m_query: str):
        db_match = DB_REGEX.search(m_query)
        schema_match = SCHEMA_REGEX.search(m_query)
        table_match = TABLE_REGEX.search(m_query)

        if not (db_match and schema_match and table_match):
            return None

        return {
            "source_type": "Azure Databricks",
            "database": db_match.group(1),
            "data_product": schema_match.group(1),
            "table": table_match.group(1)
        }
    
    # --------------------------------------------------
    # Excel extractor
    # --------------------------------------------------
    def _extract_excel_source(self, source_expression: str):
        """Extract Excel file information"""
        # Check if it's an Excel source
        if "Excel.Workbook" not in source_expression:
            return None
        
        # Try to extract URL
        url_match = re.search(r'Web\.Contents\("([^"]+)"\)', source_expression)
        file_match = re.search(r'File\.Contents\("([^"]+)"\)', source_expression)
        
        if url_match:
            url = url_match.group(1)
            # Extract filename from URL
            filename = url.split('/')[-1] if '/' in url else url
            return {
                "source_type": "Excel (Web)",
                "file_name": filename,
                "url": url
            }
        elif file_match:
            filepath = file_match.group(1)
            filename = filepath.split('\\')[-1] if '\\' in filepath else filepath
            return {
                "source_type": "Excel (File)",
                "file_name": filename,
                "file_path": filepath
            }
        else:
            return {
                "source_type": "Excel",
                "source_details": "Excel Workbook"
            }
    
    # --------------------------------------------------
    # SQL Server extractor
    # --------------------------------------------------
    def _extract_sql_source(self, source_expression: str):
        """Extract SQL Server information"""
        if "Sql.Database" not in source_expression:
            return None
        
        server_match = re.search(r'"([^"]+)"', source_expression)
        db_match = re.search(r'"([^"]+)",\s*"([^"]+)"', source_expression)
        
        if db_match:
            return {
                "source_type": "SQL Server",
                "server": db_match.group(1),
                "database": db_match.group(2)
            }
        elif server_match:
            return {
                "source_type": "SQL Server",
                "server": server_match.group(1)
            }
        
        return None
    
    # --------------------------------------------------
    # SharePoint extractor
    # --------------------------------------------------
    def _extract_sharepoint_source(self, source_expression: str):
        """Extract SharePoint information"""
        if "SharePoint" not in source_expression:
            return None
        
        url_match = re.search(r'"(https?://[^"]+)"', source_expression)
        if url_match:
            return {
                "source_type": "SharePoint",
                "url": url_match.group(1)
            }
        
        return {
            "source_type": "SharePoint",
            "source_details": "SharePoint List/Library"
        }

    # --------------------------------------------------
    # Main extraction
    # --------------------------------------------------
    def extract_table_sources(self):
        table_sources = {}

        parts = self.model_definition.get("definition", {}).get("parts", [])

        for part in parts:
            path = part.get("path", "")
            payload = part.get("payload")

            # Only table TMDL files
            if not path.startswith("definition/tables/"):
                continue

            table_name = path.split("/")[-1].replace(".tmdl", "")

            if table_name.startswith(self.IGNORE_TABLE_PREFIXES):
                continue

            # Decode payload
            try:
                content = base64.b64decode(payload).decode("utf-8")
            except Exception:
                continue

            # Extract full Source block (not just first line)
            source_block_match = re.search(
                r"Source\s*=\s*([\s\S]+?)\n\s*in\s",
                content,
                re.IGNORECASE
            )

            source_block = source_block_match.group(1).strip() if source_block_match else ""

            # --------------------------------------------------
            # Try different extractors in priority order
            # --------------------------------------------------
            
            # 1. Databricks Data Product (tbl_*)
            if table_name.lower().startswith("tbl_") and source_block:
                dp_info = self._extract_databricks_dp(source_block)
                if dp_info:
                    table_sources[table_name] = dp_info
                    continue
            
            # 2. Excel
            excel_info = self._extract_excel_source(source_block)
            if excel_info:
                table_sources[table_name] = excel_info
                continue
            
            # 3. SQL Server
            sql_info = self._extract_sql_source(source_block)
            if sql_info:
                table_sources[table_name] = sql_info
                continue
            
            # 4. SharePoint
            sp_info = self._extract_sharepoint_source(source_block)
            if sp_info:
                table_sources[table_name] = sp_info
                continue

            # --------------------------------------------------
            # DEFAULT (existing behavior)
            # --------------------------------------------------
            source_line_match = re.search(
                r"Source\s*=\s*(.+)",
                content
            )

            source_expression = (
                source_line_match.group(1).strip()
                if source_line_match
                else "Not Found"
            )

            table_sources[table_name] = {
                "source_type": "Unknown",
                "datasource_expression": source_expression
            }

        return table_sources


def build_semantic_model_with_sources(MODEL_ID, base_semantic_json, output_folder):

    model_definition = get_model_definition_lro(MODEL_ID)

    # ✅ COPY METADATA
    display_name = (
        model_definition.get("displayName")
        or model_definition.get("model", {}).get("name")
        or base_semantic_json
            .get("semantic_model_metadata", {})
            .get("displayName")
    )

    base_semantic_json["semantic_model_metadata"] = (
        base_semantic_json.get("semantic_model_metadata", {})
    )
    base_semantic_json["semantic_model_metadata"]["displayName"] = display_name

    extractor = TMDLSourceExtractor(model_definition)
    table_sources = extractor.extract_table_sources()

    for table in base_semantic_json["tables"]:
        table["source"] = table_sources.get(table["name"])

    output_path = os.path.join(
        output_folder,
        "4_semantic_model_with_table_sources.json"
    )

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(base_semantic_json, f, indent=2)

    return output_path, base_semantic_json

