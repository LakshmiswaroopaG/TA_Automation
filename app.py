from fastapi import FastAPI, HTTPException
import pandas as pd
import time
import io
from pydantic import BaseModel
from fastapi.responses import JSONResponse
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.client_credential import ClientCredential
import uvicorn
import requests
from urllib.parse import quote

app = FastAPI()

# SharePoint configuration
tenant_id = "48ad28ed-b094-4a7f-b297-482a8c33ccb4"
client_id= "c1746570-9dc1-49b3-80f0-623afb3b2a38@48ad28ed-b094-4a7f-b297-482a8c33ccb4"
client_secret = "OM2+9oAV8pDrJpxIyh4lX98ytSdsKkVQcOZp1Uo3V/M="
resource = '00000003-0000-0ff1-ce00-000000000000/aeries.sharepoint.com@48ad28ed-b094-4a7f-b297-482a8c33ccb4'
site_url="https://aeries.sharepoint.com/sites/PoA"

def get_access_token():
    url = f"https://accounts.accesscontrol.windows.net/{tenant_id}/tokens/OAuth/2"
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    body = {
        'grant_type': 'client_credentials',
        'client_id':client_id,
        'client_secret': client_secret,
        'resource': resource
    }
    response = requests.post(url, headers=headers, data=body)
    return response.json().get('access_token')


def download_file_from_sharepoint(folder: str, file_name: str) -> bytes:
    try:
        access_token = get_access_token()
        encoded_folder = quote(folder)
        encoded_file_name = quote(file_name)
        # Correct URL format for SharePoint REST API
        file_url = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('SFTP TA/Server Files/manual csv to excel')/Files('Process Report.csv')/$value"
        
        print(f"Attempting to download from URL: {file_url}")  # Debugging line
        
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(file_url, headers=headers)
        # print(response)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Error downloading file: {str(e)}")

def upload_file_to_sharepoint(folder: str, file_name: str, file_content: bytes):
    try:
        access_token = get_access_token()
        upload_url = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('SFTP TA/Server Files/manual csv to excel')/Files/add(url='Example.xlsx', overwrite=true)"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }
        response = requests.post(upload_url, headers=headers, data=file_content)
        print("upload completed")
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Error uploading file: {str(e)}")

class FileConversionRequest(BaseModel):
    source_folder: str
    source_file_name: str
    destination_folder: str
    destination_file_name: str

def format_date(date_str):
    try:
        date = pd.to_datetime(date_str)
        return date.strftime('%m/%d/%Y')
    except Exception:
        return date_str

@app.post("/convert")
async def convert_csv_to_xlsx(request: FileConversionRequest):
    try:
        # Step 1: Download CSV file from SharePoint
        csv_content = download_file_from_sharepoint(request.source_folder, request.source_file_name)
        print("download completed")
        csv_file = io.StringIO(csv_content)
        
        # Step 2: Convert CSV to XLSX
        start_time = time.time()
        # chunksize = 1000
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df = pd.read_csv(csv_file)
            for column in df.columns:
                if df[column].dtype == 'object':
                    df[column] = df[column].apply(format_date)
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        
        output.seek(0)
        # with pd.ExcelWriter(output, engine='openpyxl') as writer:
        #     chunk_size = 500
        #     start_row = 0 

            # for i, chunk in enumerate(pd.read_csv(csv_file, chunksize=chunk_size)):
            #     print(f"Processing chunk {i}, number of rows: {len(chunk)}")
            #     for column in chunk.columns:
            #         if chunk[column].dtype == 'object':
            #             chunk[column] = chunk[column].apply(format_date)
            #     if 1000 in chunk.index:
            #         print("Row 1000 found in chunk", i)        
            #     # chunk.to_excel(writer, sheet_name='Sheet1', index=False, header=not bool(i), startrow=i*chunk_size)
            #     header = (i == 0)  # Write header only for the first chunk
            #     chunk.to_excel(writer, sheet_name='Sheet1', index=False, header=header, startrow=start_row)
            #     start_row += len(chunk)
        # output.seek(0)
        print("csvto excel completed")
        # # Step 3: Upload XLSX file to SharePoint
        upload_file_to_sharepoint(request.destination_folder, request.destination_file_name, output.getvalue())
        
        end_time = time.time()
        return JSONResponse(content={"message": f"Conversion completed in {end_time - start_time:.2f} seconds"}, status_code=200)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
