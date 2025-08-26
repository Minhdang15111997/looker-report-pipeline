import sys
sys.path.append('/Users/Dang/Downloads/Auto Slides/Auto Slides/helpers/')

import pandas as pd
import time
import fitz  # PyMuPDF
from PIL import Image
import io
from pptx import Presentation
from pptx.util import Inches
import glob
from pathlib import Path
from helpers.gdrive_module import create_folder, search_files, delete_file, upload_file
import mysql.connector
from google.oauth2 import service_account
from googleapiclient.discovery import build
from helpers.support_functions import  process_batch_reports, get_venture_info_from_db, init_google_drive_service, upload_file_to_drive
from helpers.gemini_ai import summarize_pptx_with_gemini

cookies_file = r"cookies.json"
base_output_dir = r"Auto Slides/Temp"
google_auth_file = r'Auto Slides/key.json'
start_date = '20250501'
end_date = '20250531'

# Initialize Drive service and display service account email
drive_service = init_google_drive_service(google_auth_file)

print("\n1. Fetching information from database...")
venture_info = get_venture_info_from_db()
print(f"Found {len(venture_info)} venture-brand combinations")
main_data = pd.DataFrame(venture_info)

print("\n2. Starting report download and conversion...")
#Start downloading
for key, brand_venture in main_data.iterrows():

    reports_config = [{"brand_name": brand_venture['brand_name'], "venture": brand_venture['venture']} ]
    
    try:
        # Download and convert reports for this venture
        results = process_batch_reports(reports_config, base_output_dir, cookies_file, start_date, end_date)
        main_data.at[key, 'result'] = results

    except Exception as e:
        print(f"Error processing data: {str(e)}")
        continue
        
for key, brand_venture in main_data.iterrows():
    summarize_pptx_with_gemini(brand_venture['result']['pptx_path'], skip_slides=[1,5,8])
    
upload_results = []
for key, brand_venture in main_data.iterrows() :
    # Filter successfully downloaded files
    
    if brand_venture['result']['status'] != 'success':
        print(f"\nSkipping brand {brand_venture['brand_name']} from {brand_venture['venture']} - No files to upload")
        continue
    
    # Upload to drive with detailed debugging
    result = upload_file_to_drive(
        drive_service=drive_service,
        venture=brand_venture['venture'],
        parent_folder_id=brand_venture['parent_drive_folder_id'],
        files_to_upload=brand_venture['result']['pptx_path']
    )
    upload_results.append(result)

print("\n=== Processing Pipeline Completed ===")