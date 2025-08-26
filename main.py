import sys
import logging
from pathlib import Path

import pandas as pd
from helpers.gdrive_module import upload_file
from helpers.support_functions import (
    process_batch_reports,
    get_venture_info_from_db,
    init_google_drive_service,
    upload_file_to_drive,
)
from helpers.ai_content_function_testing import summarize_pptx_with_gemini

# Configurations
sys.path.append("/Users/Dang/Downloads/Auto Slides/Auto Slides/helpers/")

COOKIES_FILE = "cookies.json"
BASE_OUTPUT_DIR = "Auto Slides/Temp"
GOOGLE_AUTH_FILE = "Auto Slides/key.json"
START_DATE = "20250501"
END_DATE = "20250531"

# Setup Logging
logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)


# Step 1: Download reports
def download_reports(main_data: pd.DataFrame) -> pd.DataFrame:
    logger.info("Downloading and converting reports...")

    for idx, row in main_data.itertuples(index=True):
        reports_config = [{"brand_name": row.brand_name, "venture": row.venture}]

        try:
            results = process_batch_reports(
                reports_config, BASE_OUTPUT_DIR, COOKIES_FILE, START_DATE, END_DATE
            )
            main_data.at[idx, "result"] = results
        except Exception as e:
            logger.error(f"Error processing {row.brand_name} ({row.venture}): {e}")
            main_data.at[idx, "result"] = {"status": "failed"}

    return main_data


# Step 2: Summarize reports
def summarize_reports(main_data: pd.DataFrame) -> None:
    logger.info("Summarizing reports with Gemini...")

    for row in main_data.itertuples(index=False):
        result = row.result
        if isinstance(result, dict) and result.get("status") == "success":
            try:
                summarize_pptx_with_gemini(result["pptx_path"], skip_slides=[1, 5, 8])
            except Exception as e:
                logger.error(
                    f"Failed summarizing {row.brand_name} ({row.venture}): {e}"
                )


# Step 3: Upload reports
def upload_reports(main_data: pd.DataFrame, drive_service) -> list:
    logger.info("Uploading reports to Google Drive...")
    upload_results = []

    for row in main_data.itertuples(index=False):
        result = row.result
        if not (isinstance(result, dict) and result.get("status") == "success"):
            logger.warning(
                f"Skipping {row.brand_name} ({row.venture}) - No files to upload"
            )
            continue

        try:
            upload_result = upload_file_to_drive(
                drive_service=drive_service,
                venture=row.venture,
                parent_folder_id=row.parent_drive_folder_id,
                files_to_upload=result["pptx_path"],
            )
            upload_results.append(upload_result)
        except Exception as e:
            logger.error(f"Upload failed for {row.brand_name} ({row.venture}): {e}")

    return upload_results


# Main pipeline
def main():
    drive_service = init_google_drive_service(GOOGLE_AUTH_FILE)

    logger.info("Fetching venture info from database...")
    venture_info = get_venture_info_from_db()
    logger.info(f"Found {len(venture_info)} venture-brand combinations")

    main_data = pd.DataFrame(venture_info)

    main_data = download_reports(main_data)
    summarize_reports(main_data)
    upload_results = upload_reports(main_data, drive_service)

    logger.info("=== Processing Pipeline Completed ===")
    logger.info(f"Uploaded {len(upload_results)} reports successfully.")


if __name__ == "__main__":
    main()
