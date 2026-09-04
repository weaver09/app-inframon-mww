import os
import uuid

from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient

from mssql_python import connect


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


    try:

        original_filename = secure_filename(file.filename)

        extension = original_filename.rsplit(".", 1)[1].lower()

        unique_blob_name = (
            f"{uuid.uuid4()}-{original_filename}"
        )


        # ====================================================
        # Read file into memory
        # ====================================================

        file_data = file.read()

        file_size = len(file_data)


        # ====================================================
        # Upload to Azure Blob Storage
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
        # Save metadata to SQL
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
            VALUES
            (?, ?, ?, ?, ?);
        """,
        (
            original_filename,
            unique_blob_name,
            extension,
            file_size,
            "Uploaded"
        ))

        conn.commit()

        cursor.close()
        conn.close()


        flash(
            f"{original_filename} uploaded successfully.",
            "success"
        )


    except Exception as exc:

        flash(
            f"Upload failed: {str(exc)}",
            "error"
        )


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
    # Test SQL
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
    # Test Blob Storage
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
