import os
import sys

sys.path.append("common_modules.zip")

from base.base_logger import get_base_logger
from api_calls.admin_api import AdminApi
from api_calls.automation_api import AutomationApi
from api_calls.doc_api import DocApi
from api_calls.file_api import FileApi
from api_calls.ss_api import SSApi
from api_calls.rich_text_api import RichTextApi
from api_calls.table_api import TableApi
from api_calls.user_api import UserApi

logger = get_base_logger()

logger.info(f"Environ: {os.environ}")

admin_api = AdminApi()
automation_api = AutomationApi()
doc_api = DocApi()
file_api = FileApi()
ss_api = SSApi()
rich_text_api = RichTextApi()
table_api = TableApi()
user_api = UserApi()

CREATE_VERSION_AUTOMATION = "Create Version"
ROLLBACK_VERSION_AUTOMATION = "Rollback Version"

VERSION_FORMAT = "{major}.{minor}.{patch}"
SOURCE_DOCUMENT_ID = ""
source_document_container = ""
version_history_section_id = ""
version_section_rich_text_id = ""
version_section_rich_text_revision = ""
version_history_table_id = ""
version_history_table_sheet_id = ""

AUTOMATION_TRIGGER_USER = None
WORKIVA_CLUSTER_DOMAIN = os.getenv("WORKIVA_CLUSTER_DOMAIN")
WORKIVA_ACCOUNT_ID = os.getenv("WORKIVA_ACCOUNT_ID")

class VersionInfo:
    def __init__(self):
        self.major = None
        self.minor = None
        self.patch = None

    def __str__(self):
        return VERSION_FORMAT.format(major=self.major, minor=self.minor, patch=self.patch)
    
    def increment(self, part: str):
        if part == "major":
            self.major += 1
            self.minor = 0
            self.patch = 0
        elif part == "minor":
            self.minor += 1
            self.patch = 0
        elif part == "patch":
            self.patch += 1
        else:
            raise ValueError("Invalid version part. Use 'major', 'minor', or 'patch'.")
        
    def to_dict(self):
        return {
            "major": self.major,
            "minor": self.minor,
            "patch": self.patch
        }
    
    def from_dict(self, version_dict):
        self.major = version_dict.get("major", 0)
        self.minor = version_dict.get("minor", 0)
        self.patch = version_dict.get("patch", 0)
        return self
    
    @classmethod
    def from_string(cls, version_str):
        version_parts = version_str.split(".")
        if len(version_parts) != 3:
            raise ValueError("Version string must be in the format 'major.minor.patch'")
        instance = cls()
        instance.major = int(version_parts[0])
        instance.minor = int(version_parts[1])
        instance.patch = int(version_parts[2])
        return instance
    

def get_src_doc_id():
    document_id = os.getenv("DOCUMENT_ID", default="")
    if not document_id:
        raise ValueError("DOCUMENT_ID environment variable is not set.")
    return document_id


def check_version_history_section() -> str:
    global version_history_section_id
    all_sections = doc_api.get_list_of_sections(doc_id=SOURCE_DOCUMENT_ID)
    version_history_section = None
    version_history_section_id = None
    for section in all_sections:
        if section["name"] == "Version History":
            version_history_section = section
            version_history_section_id = version_history_section["id"]
            break
    logger.info(f"Version History Section: {version_history_section}")
    logger.info(f"Version History Section ID: {version_history_section_id}")
    return version_history_section_id


def create_version_history_section():
    global version_history_section_id
    section_config = {"name": "Version History", "index": 0, "nonPrinting": True}
    version_history_section = doc_api.create_new_section(
        doc_id=SOURCE_DOCUMENT_ID,
        body=section_config,
    )
    version_history_section_id = version_history_section["id"]
    logger.info(f"Version History Section: {version_history_section}")
    logger.info(f"Version History Section ID: {version_history_section_id}")


def get_section_rich_text():
    global version_section_rich_text_id
    global version_section_rich_text_revision
    version_section_rich_text_id, version_section_rich_text_revision = (
        rich_text_api.get_richtext_id_revision(
            doc_id=SOURCE_DOCUMENT_ID, section_id=version_history_section_id
        )
    )
    logger.info(f"Version History Section Rich Text ID: {version_section_rich_text_id}")
    logger.info(
        f"Version History Section Rich Text Revision: {version_section_rich_text_revision}"
    )


def check_version_history_table():
    global version_history_table_id
    get_section_elements = rich_text_api.get_richtext_content(
        richtext_id=version_section_rich_text_id
    )
    all_elements = get_section_elements["data"][0]["elements"]
    logger.info(f"All Elements: {all_elements}")
    table_element = [element for element in all_elements if element["type"] == "table"]
    if not len(table_element):
        logger.info("No table found in the Version History section.")
        create_version_history_table()
        return
    table_element = table_element[0]
    logger.info(f"Table Element: {table_element}")
    version_history_table_id = table_element["table"]["table"]["table"]
    version_history_table_revision = table_element["table"]["table"]["revision"]
    logger.info(f"Version Table ID: {version_history_table_id}")
    logger.info(f"Version Table Revision: {version_history_table_revision}")


def create_version_history_table():
    create_table_payload = {
        "revision": version_section_rich_text_revision,
        "isolateEdits": False,
        "data": [
            {
                "type": "insertTable",
                "insertTable": {
                    "columnCount": 6,
                    "insertAt": {"offset": 0, "paragraphIndex": 0},
                    "rowCount": 1,
                },
            }
        ],
    }
    version_table_info = rich_text_api.update_richtext_content(
        richtext_id=version_section_rich_text_id, payload=create_table_payload
    )
    logger.info(f"Version History Table Info: {version_table_info}")
    check_version_history_table()


def get_table_sheet_id():
    global version_history_table_sheet_id
    table_sheet_info = table_api.get_table_properties(table_id=version_history_table_id)
    logger.info(f"Version Table Sheet Info: {table_sheet_info}")
    version_history_table_sheet_id = table_sheet_info["sheet"]


def set_table_headers():
    ss_api.update_range(
        spreadsheet_id=SOURCE_DOCUMENT_ID,
        sheet_id=version_history_table_sheet_id,
        data_range="A1:F1",
        values=[
            [
                "Version",
                "Doc Name",
                "Link",
                "Created By",
                "Created At",
                "Select Version",
            ]
        ],
    )
    logger.info("Table headers updated successfully.")


def get_lastest_version() -> VersionInfo:
    version_info = ss_api.get_raw_data(
        spreadsheet_id=SOURCE_DOCUMENT_ID, sheet_id=version_history_table_sheet_id
    )
    logger.info(f"Version Table Info: {version_info}")
    filled_values = version_info["data"][0].get("values", [])
    
    if len(filled_values) <= 1:
        logger.info("No previous versions found.")
        return VersionInfo.from_string("0.1.0")

    latest_version_str = filled_values[1][0]    # Row right after the header
    logger.info(f"Latest Version String: {latest_version_str}")

    latest_version = VersionInfo.from_string(latest_version_str)
    latest_version.increment("minor")

    logger.info(f"Next Version: {latest_version}")
    return latest_version


def copy_file_to_version_container() -> dict:
    document_version_folder_info = file_api.get_list_of_files(
        container_id=source_document_container, name="Document Versions", kind="Folder"
    )
    logger.info(f"Document Version Folder Info: {document_version_folder_info}")
    if len(document_version_folder_info) == 0:
        document_version_folder_info = file_api.create_file(
            container_id=source_document_container,
            file_name="Document Versions",
            kind="Folder",
        )
        logger.info(f"Document Version Folder Created: {document_version_folder_info}")
    else:
        document_version_folder_info = document_version_folder_info[0]
        logger.info(f"Document Version Folder Exists: {document_version_folder_info}")

    document_versions_container = document_version_folder_info["id"]
    logger.info(f"Document Versions Container: {document_versions_container}")

    # version_info_sheet = file_api.get_list_of_files(
    #     container_id=document_versions_container,
    #     name="Version Sheet",
    #     kind="Spreadsheet",
    # )
    # logger.info(f"Version Info Sheet: {version_info_sheet}")

    # version_info_sheet_id = version_info_sheet[0]["id"]
    # logger.info(f"Version Info Sheet ID: {version_info_sheet_id}")

    # version_sheets = ss_api.get_sheets(spreadsheet_id=version_info_sheet_id)
    # logger.info(f"Sheets Info: {version_sheets}")
    # version_sheet_id = [
    #     sheet["id"] for sheet in version_sheets["data"] if sheet["name"] == "Sheet1"
    # ][0]
    # logger.info(f"Sheet ID: {version_sheet_id}")

    # version_data = ss_api.get_raw_data(
    #     spreadsheet_id=version_info_sheet_id, sheet_id=version_sheet_id
    # )
    # logger.info(f"Version Data: {version_data}")

    # version_list = version_data["data"]["values"]
    # logger.info(f"Version List: {version_list}")

    file_info = file_api.csr_copy_file(
        file_id=SOURCE_DOCUMENT_ID,
        destination_container_id=document_versions_container,
    )
    new_file_id = file_info["data"][0]["id"]
    file_details = file_api.get_file_detail(file_id=new_file_id)
    logger.info(f"File Details: {file_details}")
    return file_details


def rename_file(file_id: str, new_name: str):
    rename_payload = {
        "name": new_name
    }
    file_api.rename_file(file_id=file_id, payload=rename_payload)
    logger.info(f"File renamed to: {new_name}")


def version_file_cleanup(new_version_info):
    # remove create version automation
    # remove rollback version automation
    NEW_VERSION_FILE_ID = new_version_info["id"]
    all_automations = automation_api.get_automations_list(resource_id=f"wurl://docs.v1/doc:{NEW_VERSION_FILE_ID}/")
    for automation in all_automations.get("_embedded", {}).get("automations", []):
        if automation["name"] in (CREATE_VERSION_AUTOMATION, ROLLBACK_VERSION_AUTOMATION):
            automation_id = automation["_id"]
            automation_api.delete_automation(automation_id=automation_id)
            logger.info(f"Removed Automation: {automation['name']}")
    


def create_version(version_info: VersionInfo):
    new_file_info = copy_file_to_version_container()
    NEW_VERSION_FILE_ID = new_file_info["id"]
    NEW_VERSION_FILE_NAME = new_file_info["name"]
    NEW_VERSION_FILE_CREATED_AT = new_file_info["created"]["dateTime"]
    insert_row_payload = {
        "type": "insertRows",
        "insertRows": {
            "count": 1,
            "nextTo": 0,
            "side": "putAfter"
        },
    }
    table_api.update_table_content(table_id=version_history_table_id, payload=insert_row_payload)
    ss_api.update_range(
        spreadsheet_id=SOURCE_DOCUMENT_ID,
        sheet_id=version_history_table_sheet_id,
        data_range="A2:F2",
        values=[
            [
                str(version_info),
                f"{NEW_VERSION_FILE_NAME}",
                f"https://{WORKIVA_CLUSTER_DOMAIN}/a/{WORKIVA_ACCOUNT_ID}/doc/{NEW_VERSION_FILE_ID}",
                f"{AUTOMATION_TRIGGER_USER}",
                f"{NEW_VERSION_FILE_CREATED_AT}",
                "No",
            ]
        ],
    )
    version_file_cleanup(new_file_info)
    logger.info("Version file cleanup completed.")


def get_latest_automation_run_user():
    all_automations = automation_api.get_automations_list(resource_id=f"wurl://docs.v1/doc:{SOURCE_DOCUMENT_ID}/")
    latest_run = None
    for automation in all_automations.get("_embedded", {}).get("automations", []):
        if automation["name"] == CREATE_VERSION_AUTOMATION and automation["_summary"]["latestExecutionStatus"] != "SUCCEEDED":
            latest_run = automation
            break
    latest_executed_user_id = latest_run["_summary"]["latestExecutionContext"]["execution"]["userId"]
    logger.info(f"Latest Executed User ID: {latest_executed_user_id}")
    return latest_executed_user_id


def executed_by():
    global AUTOMATION_TRIGGER_USER
    all_users = admin_api.fetch_list_of_users()['data']
    user_dict = {user["id"]: user for user in all_users}
    latest_executed_user_id = get_latest_automation_run_user()
    if not latest_executed_user_id:
        logger.info("No latest executed user found.")
        return "Unknown"
    user_info = user_dict.get(latest_executed_user_id, None)
    if not user_info:
        logger.info("User not found in the user list.")
        return "Unknown"
    AUTOMATION_TRIGGER_USER = user_info.get("displayName", "Unknown")
    

def main():
    global SOURCE_DOCUMENT_ID
    global source_document_container
    SOURCE_DOCUMENT_ID = get_src_doc_id()
    source_document_container = file_api.get_parent_id(file_id=SOURCE_DOCUMENT_ID)
    logger.info(f"Source Document ID: {SOURCE_DOCUMENT_ID}")
    logger.info(f"Source Container: {source_document_container}")
    executed_by()
    version_history_section_id = check_version_history_section()
    create_version_history_section() if not version_history_section_id else None
    get_section_rich_text()
    check_version_history_table()
    get_table_sheet_id()
    set_table_headers()
    new_version_info = get_lastest_version()
    logger.info(f"New Version Info: {new_version_info}")
    create_version(new_version_info)


if __name__ == "__main__":
    main()
