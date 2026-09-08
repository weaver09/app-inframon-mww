import os
import io
import uuid
import json

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file
)

from werkzeug.utils import secure_filename

from azure.identity import (
    DefaultAzureCredential,
    get_bearer_token_provider
)

from azure.storage.blob import BlobServiceClient
from openai import OpenAI
from mssql_python import connect
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


# ============================================================
# Flask Application
# ============================================================

app = Flask(__name__)

app.secret_key = os.getenv(
    "FLASK_SECRET_KEY",
    "change-this-later"
)


# ============================================================
# Configuration
# ============================================================

SQL_CONNECTION_STRING = os.getenv(
    "AZURE_SQL_CONNECTIONSTRING"
)

STORAGE_ACCOUNT_NAME = os.getenv(
    "AZURE_STORAGE_ACCOUNT"
)

AZURE_OPENAI_ENDPOINT = os.getenv(
    "AZURE_OPENAI_ENDPOINT"
)

AZURE_OPENAI_DEPLOYMENT = os.getenv(
    "AZURE_OPENAI_DEPLOYMENT"
)

STORAGE_CONTAINER_NAME = "documents"

ALLOWED_EXTENSIONS = {
    "pdf",
    "docx",
    "xlsx"
}

MAX_FILE_SIZE_MB = 25

app.config["MAX_CONTENT_LENGTH"] = (
    MAX_FILE_SIZE_MB * 1024 * 1024
)


# ============================================================
# Database Connection
# ============================================================

def get_db_connection():

    if not SQL_CONNECTION_STRING:
        raise RuntimeError(
            "AZURE_SQL_CONNECTIONSTRING environment variable "
            "is not configured."
        )

    return connect(
        SQL_CONNECTION_STRING
    )


# ============================================================
# Azure Credential
# ============================================================

def get_azure_credential():

    return DefaultAzureCredential()


# ============================================================
# Azure Blob Storage
# ============================================================

def get_blob_service_client():

    if not STORAGE_ACCOUNT_NAME:
        raise RuntimeError(
            "AZURE_STORAGE_ACCOUNT environment variable "
            "is not configured."
        )

    account_url = (
        f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net"
    )

    credential = get_azure_credential()

    return BlobServiceClient(
        account_url=account_url,
        credential=credential
    )


# ============================================================
# Azure OpenAI / Foundry
# ============================================================

def get_openai_client():

    if not AZURE_OPENAI_ENDPOINT:
        raise RuntimeError(
            "AZURE_OPENAI_ENDPOINT environment variable "
            "is not configured."
        )

    if not AZURE_OPENAI_DEPLOYMENT:
        raise RuntimeError(
            "AZURE_OPENAI_DEPLOYMENT environment variable "
            "is not configured."
        )

    credential = DefaultAzureCredential()

    token_provider = get_bearer_token_provider(
        credential,
        "https://cognitiveservices.azure.com/.default"
    )

    return OpenAI(
        base_url=AZURE_OPENAI_ENDPOINT,
        api_key=token_provider
    )


# ============================================================
# Helpers
# ============================================================

def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


def safe_json_loads(value):

    if not value:
        return []

    try:
        return json.loads(value)
    except Exception:
        return []


# ============================================================
# PDF Extraction
# ============================================================

def extract_pdf(file_data):

    reader = PdfReader(
        io.BytesIO(file_data)
    )

    parts = []

    for page_number, page in enumerate(
        reader.pages,
        start=1
    ):

        text = page.extract_text() or ""

        if text.strip():
            parts.append(
                f"\n--- Page {page_number} ---\n{text}"
            )

    return "\n".join(parts)


# ============================================================
# DOCX Extraction
# ============================================================

def extract_docx(file_data):

    document = Document(
        io.BytesIO(file_data)
    )

    parts = []

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:
            parts.append(text)

    for table_index, table in enumerate(
        document.tables,
        start=1
    ):

        parts.append(
            f"\n--- Table {table_index} ---"
        )

        for row in table.rows:

            values = [
                cell.text.strip()
                for cell in row.cells
            ]

            parts.append(
                " | ".join(values)
            )

    return "\n".join(parts)


# ============================================================
# XLSX Extraction
# ============================================================

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

        for row in worksheet.iter_rows(
            values_only=True
        ):

            values = []

            for value in row:

                if value is None:
                    values.append("")
                else:
                    values.append(
                        str(value)
                    )

            if any(
                value.strip()
                for value in values
            ):
                parts.append(
                    " | ".join(values)
                )

    workbook.close()

    return "\n".join(parts)


# ============================================================
# Extract Content
# ============================================================

def extract_content(
    filename,
    file_data
):

    extension = (
        filename
        .lower()
        .rsplit(".", 1)[-1]
    )

    if extension == "pdf":
        return extract_pdf(
            file_data
        )

    if extension == "docx":
        return extract_docx(
            file_data
        )

    if extension == "xlsx":
        return extract_xlsx(
            file_data
        )

    raise ValueError(
        f"Unsupported file type: {extension}"
    )


# ============================================================
# AI Document Analysis
# ============================================================

def analyze_document_with_ai(
    extracted_text
):

    if not extracted_text:
        raise RuntimeError(
            "No extracted text was available "
            "for AI analysis."
        )

    client = get_openai_client()

    max_characters = 60000

    text_to_analyze = (
        extracted_text[:max_characters]
    )

    system_message = """
You are an AI document analysis system.

Analyze documents accurately and conservatively.

Do not invent information.

Return only valid JSON.

If information is not present in the document,
use an empty array or null value where appropriate.
"""

    user_message = f"""
Analyze the following document.

Return ONLY valid JSON using exactly this structure:

{{
    "summary": "Concise but useful document summary",

    "key_topics": [
        "topic"
    ],

    "entities": [
        {{
            "name": "entity name",
            "type": "person|organization|product|location|system|other"
        }}
    ],

    "tags": [
        "tag"
    ],

    "risks": [
        {{
            "risk": "risk description",
            "severity": "Low|Medium|High"
        }}
    ],

    "action_items": [
        {{
            "action": "action description",
            "owner": null,
            "due_date": null
        }}
    ],

    "important_dates": [
        {{
            "date": "YYYY-MM-DD or original date text",
            "description": "why this date matters"
        }}
    ]
}}

Rules:

1. Do not invent facts.
2. Summary should be factual and useful.
3. Key topics should identify major themes.
4. Entities should identify important people,
   organizations, systems, products, and locations.
5. Tags should be short and useful for searching.
6. Risks must be supported by the document.
7. If no risks exist, return an empty array.
8. Action items must come from explicit or strongly
   implied tasks in the document.
9. Preserve owners when explicitly stated.
10. Preserve due dates when explicitly stated.
11. If no action items exist, return an empty array.
12. Important dates should only include meaningful dates.
13. If no important dates exist, return an empty array.

DOCUMENT:

{text_to_analyze}
"""

    response = client.responses.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        input=[
            {
                "role": "system",
                "content": system_message
            },
            {
                "role": "user",
                "content": user_message
            }
        ]
    )

    content = response.output_text

    if not content:
        raise RuntimeError(
            "Azure OpenAI returned an empty response."
        )

    try:
        analysis = json.loads(
            content
        )

    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "Azure OpenAI returned invalid JSON."
        ) from exc

    analysis.setdefault("summary", "")
    analysis.setdefault("key_topics", [])
    analysis.setdefault("entities", [])
    analysis.setdefault("tags", [])
    analysis.setdefault("risks", [])
    analysis.setdefault("action_items", [])
    analysis.setdefault("important_dates", [])

    return analysis


# ============================================================
# Ask Document
# ============================================================

def ask_document_with_ai(
    extracted_text,
    question
):

    if not extracted_text:
        raise RuntimeError(
            "No extracted text is available."
        )

    if not question:
        raise RuntimeError(
            "A question is required."
        )

    client = get_openai_client()

    max_characters = 60000

    text_to_analyze = (
        extracted_text[:max_characters]
    )

    system_message = """
You answer questions about a supplied document.

Only use information contained in the document.

Do not invent facts.

If the answer cannot be determined from the document,
say that clearly.

Provide a concise but useful answer.
"""

    user_message = f"""
DOCUMENT:

{text_to_analyze}


QUESTION:

{question}
"""

    response = client.responses.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        input=[
            {
                "role": "system",
                "content": system_message
            },
            {
                "role": "user",
                "content": user_message
            }
        ]
    )

    answer = response.output_text

    if not answer:
        raise RuntimeError(
            "Azure OpenAI returned an empty response."
        )

    return answer


# ============================================================
# Save Analysis
# ============================================================

def save_analysis(
    cursor,
    document_id,
    analysis
):

    cursor.execute("""
        SELECT
            AnalysisID
        FROM dbo.DocumentAnalysis
        WHERE
            DocumentID = ?;
    """,
    (
        document_id,
    ))

    existing = cursor.fetchone()

    if existing:

        cursor.execute("""
            UPDATE dbo.DocumentAnalysis
            SET
                Summary = ?,
                KeyTopics = ?,
                EntitiesJson = ?,
                TagsJson = ?,
                RisksJson = ?,
                ActionItemsJson = ?,
                ImportantDatesJson = ?,
                AnalysisDate = SYSUTCDATETIME()
            WHERE
                DocumentID = ?;
        """,
        (
            analysis.get(
                "summary"
            ),

            json.dumps(
                analysis.get(
                    "key_topics",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "entities",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "tags",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "risks",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "action_items",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "important_dates",
                    []
                )
            ),

            document_id
        ))

    else:

        cursor.execute("""
            INSERT INTO dbo.DocumentAnalysis
            (
                DocumentID,
                Summary,
                KeyTopics,
                EntitiesJson,
                TagsJson,
                RisksJson,
                ActionItemsJson,
                ImportantDatesJson
            )
            VALUES
            (
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?
            );
        """,
        (
            document_id,

            analysis.get(
                "summary"
            ),

            json.dumps(
                analysis.get(
                    "key_topics",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "entities",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "tags",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "risks",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "action_items",
                    []
                )
            ),

            json.dumps(
                analysis.get(
                    "important_dates",
                    []
                )
            )
        ))


# ============================================================
# Home Page
# ============================================================

@app.route("/")
def index():

    documents = []
    error_message = None

    search_text = (
        request.args.get(
            "q",
            ""
        ).strip()
    )

    file_type = (
        request.args.get(
            "type",
            ""
        ).strip().lower()
    )

    status_filter = (
        request.args.get(
            "status",
            ""
        ).strip()
    )

    stats = {
        "TotalDocuments": 0,
        "AnalyzedDocuments": 0,
        "FailedDocuments": 0,
        "TotalCharacters": 0,
        "PDFDocuments": 0,
        "DOCXDocuments": 0,
        "XLSXDocuments": 0
    }

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        # ====================================================
        # Dashboard Statistics
        # ====================================================

        cursor.execute("""
            SELECT
                COUNT(*) AS TotalDocuments,

                SUM(
                    CASE
                        WHEN Status = 'Analyzed'
                        THEN 1
                        ELSE 0
                    END
                ) AS AnalyzedDocuments,

                SUM(
                    CASE
                        WHEN Status = 'Processing Failed'
                        THEN 1
                        ELSE 0
                    END
                ) AS FailedDocuments,

                SUM(
                    ISNULL(
                        LEN(ExtractedText),
                        0
                    )
                ) AS TotalCharacters,

                SUM(
                    CASE
                        WHEN FileType = 'pdf'
                        THEN 1
                        ELSE 0
                    END
                ) AS PDFDocuments,

                SUM(
                    CASE
                        WHEN FileType = 'docx'
                        THEN 1
                        ELSE 0
                    END
                ) AS DOCXDocuments,

                SUM(
                    CASE
                        WHEN FileType = 'xlsx'
                        THEN 1
                        ELSE 0
                    END
                ) AS XLSXDocuments

            FROM dbo.Documents;
        """)

        row = cursor.fetchone()

        if row:

            stats = {
                "TotalDocuments": row[0] or 0,
                "AnalyzedDocuments": row[1] or 0,
                "FailedDocuments": row[2] or 0,
                "TotalCharacters": row[3] or 0,
                "PDFDocuments": row[4] or 0,
                "DOCXDocuments": row[5] or 0,
                "XLSXDocuments": row[6] or 0
            }

        # ====================================================
        # Document Search
        # ====================================================

        sql = """
            SELECT
                d.DocumentID,
                d.FileName,
                d.BlobName,
                d.FileType,
                d.FileSizeBytes,
                d.UploadDate,
                d.Status,
                LEN(d.ExtractedText)
                    AS ExtractedCharacters,
                a.AnalysisID,
                a.Summary
            FROM dbo.Documents d
            LEFT JOIN dbo.DocumentAnalysis a
                ON d.DocumentID = a.DocumentID
            WHERE
                1 = 1
        """

        parameters = []

        if search_text:

            sql += """
                AND
                (
                    d.FileName LIKE ?
                    OR a.Summary LIKE ?
                    OR a.KeyTopics LIKE ?
                    OR a.TagsJson LIKE ?
                )
            """

            wildcard = (
                f"%{search_text}%"
            )

            parameters.extend([
                wildcard,
                wildcard,
                wildcard,
                wildcard
            ])

        if file_type:

            sql += """
                AND d.FileType = ?
            """

            parameters.append(
                file_type
            )

        if status_filter:

            sql += """
                AND d.Status = ?
            """

            parameters.append(
                status_filter
            )

        sql += """
            ORDER BY
                d.UploadDate DESC;
        """

        cursor.execute(
            sql,
            tuple(parameters)
        )

        rows = cursor.fetchall()

        for row in rows:

            documents.append({
                "DocumentID": row[0],
                "FileName": row[1],
                "BlobName": row[2],
                "FileType": row[3],
                "FileSizeBytes": row[4],
                "UploadDate": row[5],
                "Status": row[6],
                "ExtractedCharacters": row[7],
                "AnalysisID": row[8],
                "Summary": row[9]
            })

    except Exception as exc:

        error_message = str(exc)

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

    return render_template(
        "index.html",
        documents=documents,
        error_message=error_message,
        stats=stats,
        search_text=search_text,
        file_type=file_type,
        status_filter=status_filter
    )


# ============================================================
# Analysis Page
# ============================================================

@app.route("/analysis/<int:document_id>")
def view_analysis(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                d.DocumentID,
                d.FileName,
                d.FileType,
                d.UploadDate,
                d.Status,
                LEN(d.ExtractedText)
                    AS ExtractedCharacters,
                a.Summary,
                a.KeyTopics,
                a.EntitiesJson,
                a.TagsJson,
                a.RisksJson,
                a.ActionItemsJson,
                a.ImportantDatesJson,
                a.AnalysisDate
            FROM dbo.Documents d
            LEFT JOIN dbo.DocumentAnalysis a
                ON d.DocumentID = a.DocumentID
            WHERE
                d.DocumentID = ?;
        """,
        (
            document_id,
        ))

        row = cursor.fetchone()

        if not row:
            return "Document not found.", 404

        analysis = {
            "DocumentID": row[0],
            "FileName": row[1],
            "FileType": row[2],
            "UploadDate": row[3],
            "Status": row[4],
            "ExtractedCharacters": row[5],
            "Summary": row[6],
            "KeyTopics": safe_json_loads(
                row[7]
            ),
            "Entities": safe_json_loads(
                row[8]
            ),
            "Tags": safe_json_loads(
                row[9]
            ),
            "Risks": safe_json_loads(
                row[10]
            ),
            "ActionItems": safe_json_loads(
                row[11]
            ),
            "ImportantDates": safe_json_loads(
                row[12]
            ),
            "AnalysisDate": row[13]
        }

        return render_template(
            "analysis.html",
            analysis=analysis,
            question=None,
            answer=None
        )

    except Exception as exc:

        return (
            f"Unable to load analysis: {str(exc)}",
            500
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


# ============================================================
# Ask This Document
# ============================================================

@app.route(
    "/analysis/<int:document_id>/ask",
    methods=["POST"]
)
def ask_document(document_id):

    question = (
        request.form.get(
            "question",
            ""
        ).strip()
    )

    if not question:

        flash(
            "Enter a question first.",
            "error"
        )

        return redirect(
            url_for(
                "view_analysis",
                document_id=document_id
            )
        )

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                d.DocumentID,
                d.FileName,
                d.FileType,
                d.UploadDate,
                d.Status,
                LEN(d.ExtractedText)
                    AS ExtractedCharacters,
                d.ExtractedText,
                a.Summary,
                a.KeyTopics,
                a.EntitiesJson,
                a.TagsJson,
                a.RisksJson,
                a.ActionItemsJson,
                a.ImportantDatesJson,
                a.AnalysisDate
            FROM dbo.Documents d
            LEFT JOIN dbo.DocumentAnalysis a
                ON d.DocumentID = a.DocumentID
            WHERE
                d.DocumentID = ?;
        """,
        (
            document_id,
        ))

        row = cursor.fetchone()

        if not row:
            return "Document not found.", 404

        answer = ask_document_with_ai(
            row[6],
            question
        )

        analysis = {
            "DocumentID": row[0],
            "FileName": row[1],
            "FileType": row[2],
            "UploadDate": row[3],
            "Status": row[4],
            "ExtractedCharacters": row[5],
            "Summary": row[7],
            "KeyTopics": safe_json_loads(
                row[8]
            ),
            "Entities": safe_json_loads(
                row[9]
            ),
            "Tags": safe_json_loads(
                row[10]
            ),
            "Risks": safe_json_loads(
                row[11]
            ),
            "ActionItems": safe_json_loads(
                row[12]
            ),
            "ImportantDates": safe_json_loads(
                row[13]
            ),
            "AnalysisDate": row[14]
        }

        return render_template(
            "analysis.html",
            analysis=analysis,
            question=question,
            answer=answer
        )

    except Exception as exc:

        flash(
            f"Question failed: {str(exc)}",
            "error"
        )

        return redirect(
            url_for(
                "view_analysis",
                document_id=document_id
            )
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


# ============================================================
# Re-analyze Document
# ============================================================

@app.route(
    "/analysis/<int:document_id>/reanalyze",
    methods=["POST"]
)
def reanalyze_document(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                ExtractedText
            FROM dbo.Documents
            WHERE
                DocumentID = ?;
        """,
        (
            document_id,
        ))

        row = cursor.fetchone()

        if not row:
            return "Document not found.", 404

        extracted_text = row[0]

        if not extracted_text:

            raise RuntimeError(
                "This document has no extracted text."
            )

        cursor.execute("""
            UPDATE dbo.Documents
            SET
                Status = ?
            WHERE
                DocumentID = ?;
        """,
        (
            "Processing",
            document_id
        ))

        conn.commit()

        analysis = analyze_document_with_ai(
            extracted_text
        )

        save_analysis(
            cursor,
            document_id,
            analysis
        )

        cursor.execute("""
            UPDATE dbo.Documents
            SET
                Status = ?
            WHERE
                DocumentID = ?;
        """,
        (
            "Analyzed",
            document_id
        ))

        conn.commit()

        flash(
            "Document re-analyzed successfully.",
            "success"
        )

    except Exception as exc:

        try:

            if conn and cursor:

                cursor.execute("""
                    UPDATE dbo.Documents
                    SET
                        Status = ?
                    WHERE
                        DocumentID = ?;
                """,
                (
                    "Processing Failed",
                    document_id
                ))

                conn.commit()

        except Exception:
            pass

        flash(
            f"Re-analysis failed: {str(exc)}",
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

    return redirect(
        url_for(
            "view_analysis",
            document_id=document_id
        )
    )


# ============================================================
# Download Original Document
# ============================================================

@app.route(
    "/document/<int:document_id>/download"
)
def download_document(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                FileName,
                BlobName
            FROM dbo.Documents
            WHERE
                DocumentID = ?;
        """,
        (
            document_id,
        ))

        row = cursor.fetchone()

        if not row:
            return "Document not found.", 404

        filename = row[0]
        blob_name = row[1]

        blob_service_client = (
            get_blob_service_client()
        )

        blob_client = (
            blob_service_client
            .get_blob_client(
                container=STORAGE_CONTAINER_NAME,
                blob=blob_name
            )
        )

        file_data = (
            blob_client
            .download_blob()
            .readall()
        )

        return send_file(
            io.BytesIO(file_data),
            as_attachment=True,
            download_name=filename
        )

    except Exception as exc:

        return (
            f"Download failed: {str(exc)}",
            500
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


# ============================================================
# Delete Document
# ============================================================

@app.route(
    "/document/<int:document_id>/delete",
    methods=["POST"]
)
def delete_document(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                FileName,
                BlobName
            FROM dbo.Documents
            WHERE
                DocumentID = ?;
        """,
        (
            document_id,
        ))

        row = cursor.fetchone()

        if not row:

            flash(
                "Document not found.",
                "error"
            )

            return redirect(
                url_for("index")
            )

        filename = row[0]
        blob_name = row[1]

        # Delete the Blob first
        blob_service_client = (
            get_blob_service_client()
        )

        blob_client = (
            blob_service_client
            .get_blob_client(
                container=STORAGE_CONTAINER_NAME,
                blob=blob_name
            )
        )

        blob_client.delete_blob(
            delete_snapshots="include"
        )

        # Delete child analysis record first
        cursor.execute("""
            DELETE FROM dbo.DocumentAnalysis
            WHERE
                DocumentID = ?;
        """,
        (
            document_id,
        ))

        # Delete document record
        cursor.execute("""
            DELETE FROM dbo.Documents
            WHERE
                DocumentID = ?;
        """,
        (
            document_id,
        ))

        conn.commit()

        flash(
            f"{filename} was deleted.",
            "success"
        )

    except Exception as exc:

        try:
            if conn:
                conn.rollback()
        except Exception:
            pass

        flash(
            f"Delete failed: {str(exc)}",
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

    return redirect(
        url_for("index")
    )


# ============================================================
# Upload Document
# ============================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload_document():

    if "file" not in request.files:

        flash(
            "No file was selected.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    file = request.files["file"]

    if file.filename == "":

        flash(
            "No file was selected.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    if not allowed_file(
        file.filename
    ):

        flash(
            "Unsupported file type. "
            "Only PDF, DOCX, and XLSX "
            "files are allowed.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    original_filename = (
        secure_filename(
            file.filename
        )
    )

    extension = (
        original_filename
        .rsplit(".", 1)[1]
        .lower()
    )

    unique_blob_name = (
        f"{uuid.uuid4()}-"
        f"{original_filename}"
    )

    document_id = None
    conn = None
    cursor = None

    try:

        file_data = file.read()

        file_size = len(
            file_data
        )

        blob_service_client = (
            get_blob_service_client()
        )

        blob_client = (
            blob_service_client
            .get_blob_client(
                container=STORAGE_CONTAINER_NAME,
                blob=unique_blob_name
            )
        )

        blob_client.upload_blob(
            file_data,
            overwrite=False
        )

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
            OUTPUT
                INSERTED.DocumentID
            VALUES
            (
                ?,
                ?,
                ?,
                ?,
                ?
            );
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

        extracted_text = (
            extract_content(
                original_filename,
                file_data
            )
        )

        cursor.execute("""
            UPDATE dbo.Documents
            SET
                ExtractedText = ?,
                Status = ?
            WHERE
                DocumentID = ?;
        """,
        (
            extracted_text,
            "Extracted",
            document_id
        ))

        conn.commit()

        analysis = (
            analyze_document_with_ai(
                extracted_text
            )
        )

        save_analysis(
            cursor,
            document_id,
            analysis
        )

        cursor.execute("""
            UPDATE dbo.Documents
            SET
                Status = ?
            WHERE
                DocumentID = ?;
        """,
        (
            "Analyzed",
            document_id
        ))

        conn.commit()

        return redirect(
            url_for(
                "view_analysis",
                document_id=document_id
            )
        )

    except Exception as exc:

        try:

            if document_id:

                if conn is None:
                    conn = (
                        get_db_connection()
                    )

                if cursor is None:
                    cursor = (
                        conn.cursor()
                    )

                cursor.execute("""
                    UPDATE dbo.Documents
                    SET
                        Status = ?
                    WHERE
                        DocumentID = ?;
                """,
                (
                    "Processing Failed",
                    document_id
                ))

                conn.commit()

        except Exception:
            pass

        flash(
            f"Upload failed: "
            f"{str(exc)}",
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

    return redirect(
        url_for("index")
    )


# ============================================================
# Health Check
# ============================================================

@app.route("/health")
def health():

    results = {
        "database": "unknown",
        "storage": "unknown",
        "openai": "unknown"
    }

    status_code = 200

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute(
            "SELECT 1;"
        )

        cursor.fetchone()

        cursor.close()
        conn.close()

        results["database"] = "connected"

    except Exception as exc:

        results["database"] = (
            f"error: {str(exc)}"
        )

        status_code = 500

    try:

        blob_service_client = (
            get_blob_service_client()
        )

        container_client = (
            blob_service_client
            .get_container_client(
                STORAGE_CONTAINER_NAME
            )
        )

        container_client.get_container_properties()

        results["storage"] = "connected"

    except Exception as exc:

        results["storage"] = (
            f"error: {str(exc)}"
        )

        status_code = 500

    try:

        if not AZURE_OPENAI_ENDPOINT:

            raise RuntimeError(
                "AZURE_OPENAI_ENDPOINT "
                "is not configured."
            )

        if not AZURE_OPENAI_DEPLOYMENT:

            raise RuntimeError(
                "AZURE_OPENAI_DEPLOYMENT "
                "is not configured."
            )

        get_openai_client()

        results["openai"] = "configured"

    except Exception as exc:

        results["openai"] = (
            f"error: {str(exc)}"
        )

        status_code = 500

    results["status"] = (
        "healthy"
        if status_code == 200
        else "unhealthy"
    )

    return (
        results,
        status_code
    )


# ============================================================
# Application Entry Point
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=8000
    )