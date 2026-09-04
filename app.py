import os
import io
import uuid

from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

from mssql_python import connect
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-later")


# ============================================================
# Configuration
# ============================================================

SQL_CONNECTION_STRING = os.getenv("AZURE_SQL_CONNECTIONSTRING")
STORAGE_ACCOUNT_NAME = os.getenv("AZURE_STORAGE_ACCOUNT")

STORAGE_CONTAINER_NAME = "documents"

ALLOWED_EXTENSIONS = {
    "pdf",
    "docx",
    "xlsx"
}

MAX_FILE_SIZE_MB = 25

app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE_MB * 1024 * 1024


# ============================================================
# Database
# ============================================================

def get_db_connection():
    if not SQL_CONNECTION_STRING:
        raise RuntimeError(
            "AZURE_SQL_CONNECTIONSTRING environment variable is not configured."
        )

    return connect(SQL_CONNECTION_STRING)


# ============================================================
# Azure Blob Storage
# ============================================================

def get_blob_service_client():
    if not STORAGE_ACCOUNT_NAME:
        raise RuntimeError(
            "AZURE_STORAGE_ACCOUNT environment variable is not configured."
        )

    account_url = (
        f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net"
    )

    credential = DefaultAzureCredential()

    return BlobServiceClient(
        account_url=account_url,
        credential=credential
    )


# ============================================================
# File Validation
# ============================================================

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


# ============================================================
# Content Extraction
# ============================================================

def extract_pdf(file_data):
    reader = PdfReader(io.BytesIO(file_data))

    parts = []

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""

        if text.strip():
            parts.append(
                f"\n--- Page {page_number} ---\n{text}"
            )

    return "\n".join(parts)


def extract_docx(file_data):
    document = Document(io.BytesIO(file_data))

    parts = []

    # Paragraphs
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()

        if text:
            parts.append(text)

    # Tables
    for table_index, table in enumerate(document.tables, start=1):
        parts.append(
            f"\n--- Table {table_index} ---"
        )

        for row in table.rows:
            values = [
                cell.text.strip()
                for cell in row.cells
            ]

            parts.append(" | ".join(values))

    return "\n".join(parts)


def extract_xlsx(file_data):
    workbook = load_workbook(
        io.BytesIO(file_data),
        read_only=True,
        data_only=True
    )

    parts = []

    for worksheet in workbook.worksheets:
        parts.append(
            f"\n--- Worksheet: {worksheet.title} ---"
        )

        for row in worksheet.iter_rows(values_only=True):
            values = []

            for value in row:
                if value is None:
                    values.append("")
                else:
                    values.append(str(value))

            if any(value.strip() for value in values):
                parts.append(" | ".join(values))

    workbook.close()

    return "\n".join(parts)


def extract_content(filename, file_data):
    extension = filename.lower().rsplit(".", 1)[-1]

    if extension == "pdf":
        return extract_pdf(file_data)

    if extension == "docx":
        return extract_docx(file_data)

    if extension == "xlsx":
        return extract_xlsx(file_data)

    raise ValueError(
        f"Unsupported file type: {extension}"
    )


# ============================================================
# Home Page
# ============================================================

@app.route("/")
def index():
    documents = []
    error_message = None

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                DocumentID,
                FileName,
                BlobName,
                FileType,
                FileSizeBytes,
                UploadDate,
                Status
            FROM dbo.Documents
            ORDER BY UploadDate DESC;
        """)

        rows = cursor.fetchall()

        for row in rows:
            documents.append({
                "DocumentID": row[0],
                "FileName": row[1],
                "BlobName": row[2],
                "FileType": row[3],
                "FileSizeBytes": row[4],
                "UploadDate": row[5],
                "Status": row[6]
            })

        cursor.close()
        conn.close()

    except Exception as exc:
        error_message = str(exc)

    return render_template(
        "index.html",
        documents=documents,
        error_message=error_message
    )


# ============================================================
# Upload Document
# ============================================================

@app.route("/upload", methods=["POST"])
def upload_document():
    if "file" not in request.files:
        flash("No file was selected.", "error")
        return redirect(url_for("index"))

    file = request.files["file"]

    if file.filename == "":
        flash("No file was selected.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash(
            "Unsupported file type. Only PDF, DOCX, and XLSX files are allowed.",
            "error"
        )
        return redirect(url_for("index"))

    original_filename = secure_filename(file.filename)
    extension = original_filename.rsplit(".", 1)[1].lower()

    unique_blob_name = (
        f"{uuid.uuid4()}-{original_filename}"
    )

    document_id = None
    conn = None
    cursor = None

    try:
        # ====================================================
        # Read file
        # ====================================================

        file_data = file.read()
        file_size = len(file_data)

        # ====================================================
        # Upload to Blob Storage
        # ====================================================

        blob_service_client = get_blob_service_client()

        blob_client = blob_service_client.get_blob_client(
            container=STORAGE_CONTAINER_NAME,
            blob=unique_blob_name
        )

        blob_client.upload_blob(
            file_data,
            overwrite=False
        )

        # ====================================================
        # Create SQL metadata row
        # ====================================================

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO dbo.Documents
            (
                FileName,
                BlobName,
                FileType,
                FileSizeBytes,
                Status
            )
            OUTPUT INSERTED.DocumentID
            VALUES
            (?, ?, ?, ?, ?);
        """,
        (
            original_filename,
            unique_blob_name,
            extension,
            file_size,
            "Processing"
        ))

        row = cursor.fetchone()
        document_id = row[0]

        conn.commit()

        # ====================================================
        # Extract document content
        # ====================================================

        extracted_text = extract_content(
            original_filename,
            file_data
        )

        # ====================================================
        # Update SQL with extracted content
        # ====================================================

        cursor.execute("""
            UPDATE dbo.Documents
            SET
                ExtractedText = ?,
                Status = ?
            WHERE DocumentID = ?;
        """,
        (
            extracted_text,
            "Processed",
            document_id
        ))

        conn.commit()

        flash(
            f"{original_filename} uploaded and processed successfully.",
            "success"
        )

    except Exception as exc:
        try:
            if document_id:
                if conn is None:
                    conn = get_db_connection()

                if cursor is None:
                    cursor = conn.cursor()

                cursor.execute("""
                    UPDATE dbo.Documents
                    SET Status = ?
                    WHERE DocumentID = ?;
                """,
                (
                    "Processing Failed",
                    document_id
                ))

                conn.commit()

        except Exception:
            pass

        flash(
            f"Upload failed: {str(exc)}",
            "error"
        )

    finally:
        try:
            if cursor:
                cursor.close()
        except Exception:
            pass

        try:
            if conn:
                conn.close()
        except Exception:
            pass

    return redirect(url_for("index"))


# ============================================================
# Health Check
# ============================================================

@app.route("/health")
def health():
    results = {
        "database": "unknown",
        "storage": "unknown"
    }

    status_code = 200

    # ========================================================
    # SQL Test
    # ========================================================

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT 1;")
        cursor.fetchone()

        cursor.close()
        conn.close()

        results["database"] = "connected"

    except Exception as exc:
        results["database"] = f"error: {str(exc)}"
        status_code = 500

    # ========================================================
    # Blob Storage Test
    # ========================================================

    try:
        blob_service_client = get_blob_service_client()

        container_client = (
            blob_service_client.get_container_client(
                STORAGE_CONTAINER_NAME
            )
        )

        container_client.get_container_properties()

        results["storage"] = "connected"

    except Exception as exc:
        results["storage"] = f"error: {str(exc)}"
        status_code = 500

    results["status"] = (
        "healthy"
        if status_code == 200
        else "unhealthy"
    )

    return results, status_code


# ============================================================
# Application Entry Point
# ============================================================

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=8000
    )