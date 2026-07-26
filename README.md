# Document Parser Service

Extracts text from files uploaded through the YZNU agent platform so a workflow can classify the course before writing the original file to a knowledge base.

Supported formats: `PDF`, `DOCX`, `PPTX`, and `TXT`.

Legacy `DOC` and `PPT` files must be converted to `DOCX` and `PPTX` first. Scanned PDFs need an OCR service and are intentionally rejected.

## Run Locally

```powershell
cd document-parser-service
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:PARSER_API_KEY='replace-with-a-long-random-secret'
uvicorn app:app --host 0.0.0.0 --port 8000
```

Verify the service:

```powershell
Invoke-WebRequest http://127.0.0.1:8000/health
```

## HTTP API

`POST /v1/extract`

Required header:

```text
X-API-Key: <PARSER_API_KEY>
```

JSON request body:

```json
{
  "file_url": "https://ai.yznu.edu.cn/api/proxy/down?...",
  "file_name": "courseware.pptx",
  "max_chars": 12000
}
```

Successful response:

```json
{
  "filename": "courseware.pptx",
  "document_type": "pptx",
  "text": "Extracted document text...",
  "truncated": false
}
```

## Deploy With Docker

Deploy this folder to a public container host such as Render. Set these environment variables in the host dashboard:

```text
PARSER_API_KEY=<long random secret>
ALLOWED_FILE_HOSTS=ai.yznu.edu.cn
```

The YZNU workflow must call the public HTTPS endpoint. A local `localhost` service is not reachable from the cloud workflow runtime.
