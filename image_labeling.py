import os
import pandas as pd
import requests
from pathlib import Path
import mimetypes
from urllib.parse import urlparse
import time

def download_image(url, barcode, folder='images'):
    """
    Downloads an image from the given URL and saves it with the barcode as filename
    """
    try:
        # create the images folder
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Created directory: {folder}")
        
        # cownload the image
        print(f"Downloading: {url}")
        response = requests.get(url, stream=True, timeout=10)
        
        if response.status_code != 200:
            print(f"Failed to download {url}, status code: {response.status_code}")
            return False
        
        # file extension
        content_type = response.headers.get('content-type')
        if content_type and 'image' in content_type:
            ext = mimetypes.guess_extension(content_type)
        else:
            # Try to get extension from the URL
            path = urlparse(url).path
            ext = os.path.splitext(path)[1]
            
        # default to .jpg
        if not ext or ext == '.jpe':
            ext = '.jpg'
        
        # clean extension
        ext = '.' + ext.lstrip('.')
        
        # filename with barcode
        filename = f"{barcode}{ext}"
        filepath = os.path.join(folder, filename)
        
        # save the image
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        print(f" Saved: {filepath}")
        return True
    
    except Exception as e:
        print(f" error {url}: {str(e)}")
        return False

def process_imictoko1_excel(excel_file="imictoko1.xlsx"):
    """
    processes the xlsx file
    """
    try:
        # Check if file exists
        if not os.path.exists(excel_file):
            print(f"  not found: {excel_file}")
            files = [f for f in os.listdir('../../../../Downloads') if f.endswith('.xlsx')]
            if files:
                print(f"excel files: {', '.join(files)}")
            return
            
        # read excel
        print(f"processing : {excel_file}")
        df = pd.read_excel(excel_file)
        
        # verify columns
        if 'image' not in df.columns or 'barkod' not in df.columns:
            print(f" must have 'image' and 'barkod'")
            print(f"found columns: {', '.join(df.columns)}")
            return
        
        total_rows = len(df)
        print(f"found {total_rows} images to download")
        
        # create counters for progress tracking
        success_count = 0
        failed_count = 0
        
        # process each row
        start_time = time.time()
        for index, row in df.iterrows():
            url = row['image']
            barcode = row['barkod']
            
            # skip if URL or barcode is empty
            if pd.isna(url) or pd.isna(barcode) or not url or not barcode:
                print(f" Skipping row {index+1}: Missing URL or barcode")
                failed_count += 1
                continue
            
            # Show progress
            print(f"\n[{index+1}/{total_rows}] Processing {barcode}")
            
            # Download the image
            if download_image(url, barcode):
                success_count += 1
            else:
                failed_count += 1
            
            # small delay to be nice to the server
            time.sleep(0.5)
        
        elapsed_time = time.time() - start_time
        
        # print summary
        print(f"\n Download summary for {excel_file}:")
        print(f" Time elapsed: {elapsed_time:.1f} seconds")
        print(f" Total images: {total_rows}")
        print(f" Successfully downloaded: {success_count}")
        print(f" Failed: {failed_count}")
        print(f" Images saved to: {os.path.abspath('images')}")
    
    except Exception as e:
        print(f" Error processing Excel file: {str(e)}")

if __name__ == "__main__":
    # try to find .xlsx in current directory
    if os.path.exists("imıcty2.xlsx"):
        process_imictoko1_excel("imıcty2.xlsx")
    else:
        # If not found, ask for the correct filename
        print("⚠ imictoko1.xlsx not found in current directory")
        excel_file = input("Enter the exact path to your Excel file (or press Enter to search): ")
        
        if excel_file:
            process_imictoko1_excel(excel_file)
        else:
            # try to find any excel files
            files = [f for f in os.listdir('../../../../Downloads') if f.endswith('.xlsx')]
            if not files:
                print(" No Excel files found in current directory")
            else:
                print(" Found these Excel files:")
                for i, file in enumerate(files):
                    print(f"  {i+1}. {file}")
                
                choice = input("\nEnter the number of the file to use (or press Enter to quit): ")
                if choice.strip() and choice.isdigit():
                    index = int(choice) - 1
                    if 0 <= index < len(files):
                        process_imictoko1_excel(files[index])
                    else:
                        print(" Invalid selection")