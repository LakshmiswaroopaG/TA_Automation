@echo off
cd C:\Users\Administrator\Desktop\TA_apis\TA_Automation
call myenv\Scripts\activate
uvicorn app:app --host 0.0.0.0 --port 8001