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

from reportlab.lib.pagesizes import letter

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

from reportlab.lib import colors

from reportlab.lib.styles import (
    getSampleStyleSheet
)

from document_processing import (
    ALLOWED_EXTENSIONS,
    STORAGE_CONTAINER_NAME,
    allowed_file,
    safe_json_loads,
    get_db_connection,
    get_blob_service_client,
    get_openai_client,
    ask_document_with_ai,
    analyze_document_with_ai,
    save_analysis
)


# ============================================================
# Flask Application
# ============================================================

app = Flask(__name__)

app.secret_key = os.getenv(
    "FLASK_SECRET_KEY",
    "change-this-later"
)

MAX_FILE_SIZE_MB = 25

app.config["MAX_CONTENT_LENGTH"] = (
    MAX_FILE_SIZE_MB * 1024 * 1024
)


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

        cursor.execute(
            """
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
            """
        )

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

        sql = """
            SELECT
                d.DocumentID,
                d.FileName,
                d.BlobName,
                d.FileType,
                d.FileSizeBytes,
                d.UploadDate,
                d.Status,

                LEN(
                    d.ExtractedText
                ) AS ExtractedCharacters,

                a.AnalysisID,
                a.Summary

            FROM dbo.Documents d

            LEFT JOIN dbo.DocumentAnalysis a
                ON d.DocumentID = a.DocumentID

            WHERE 1 = 1
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

            parameters.extend(
                [
                    wildcard,
                    wildcard,
                    wildcard,
                    wildcard
                ]
            )

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

            documents.append(
                {
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
                }
            )

    except Exception as exc:

        error_message = str(
            exc
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

    return render_template(
        "index.html",
        documents=documents,
        error_message=error_message,
        stats=stats,
        search_text=search_text,
        file_type=file_type,
        status_filter=status_filter,
        allowed_extensions=sorted(
            ALLOWED_EXTENSIONS
        )
    )


# ============================================================
# Analysis Page
# ============================================================

@app.route(
    "/analysis/<int:document_id>"
)
def view_analysis(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT
                d.DocumentID,
                d.FileName,
                d.FileType,
                d.UploadDate,
                d.Status,

                LEN(
                    d.ExtractedText
                ) AS ExtractedCharacters,

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
            )
        )

        row = cursor.fetchone()

        if not row:

            return (
                "Document not found.",
                404
            )

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

        cursor.execute(
            """
            SELECT
                d.DocumentID,
                d.FileName,
                d.FileType,
                d.UploadDate,
                d.Status,

                LEN(
                    d.ExtractedText
                ) AS ExtractedCharacters,

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
            )
        )

        row = cursor.fetchone()

        if not row:

            return (
                "Document not found.",
                404
            )

        if row[4] != "Analyzed":

            flash(
                "This document has not finished processing yet.",
                "error"
            )

            return redirect(
                url_for(
                    "view_analysis",
                    document_id=document_id
                )
            )

        answer = (
            ask_document_with_ai(
                row[6],
                question
            )
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
#
# For now, re-analysis still runs synchronously.
# We can queue this later if desired.
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

        cursor.execute(
            """
            SELECT
                ExtractedText

            FROM dbo.Documents

            WHERE
                DocumentID = ?;
            """,
            (
                document_id,
            )
        )

        row = cursor.fetchone()

        if not row:

            return (
                "Document not found.",
                404
            )

        extracted_text = row[0]

        if not extracted_text:

            raise RuntimeError(
                "This document has no extracted text."
            )

        cursor.execute(
            """
            UPDATE dbo.Documents

            SET
                Status = ?

            WHERE
                DocumentID = ?;
            """,
            (
                "Processing",
                document_id
            )
        )

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

        cursor.execute(
            """
            UPDATE dbo.Documents

            SET
                Status = ?

            WHERE
                DocumentID = ?;
            """,
            (
                "Analyzed",
                document_id
            )
        )

        conn.commit()

        flash(
            "Document re-analyzed successfully.",
            "success"
        )

    except Exception as exc:

        try:

            if conn and cursor:

                cursor.execute(
                    """
                    UPDATE dbo.Documents

                    SET
                        Status = ?

                    WHERE
                        DocumentID = ?;
                    """,
                    (
                        "Processing Failed",
                        document_id
                    )
                )

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
# Export Analysis PDF
# ============================================================

@app.route(
    "/analysis/<int:document_id>/export/pdf"
)
def export_analysis_pdf(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()

        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT
                d.FileName,
                d.FileType,
                d.UploadDate,
                d.Status,

                LEN(
                    d.ExtractedText
                ) AS ExtractedCharacters,

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
            )
        )

        row = cursor.fetchone()

        if not row:

            return (
                "Document not found.",
                404
            )

        filename = row[0]

        file_type = row[1]

        upload_date = row[2]

        status = row[3]

        extracted_characters = row[4]

        summary = row[5]

        key_topics = (
            safe_json_loads(
                row[6]
            )
        )

        entities = (
            safe_json_loads(
                row[7]
            )
        )

        tags = (
            safe_json_loads(
                row[8]
            )
        )

        risks = (
            safe_json_loads(
                row[9]
            )
        )

        action_items = (
            safe_json_loads(
                row[10]
            )
        )

        important_dates = (
            safe_json_loads(
                row[11]
            )
        )

        analysis_date = row[12]

        if status != "Analyzed":

            return (
                "Document analysis is not complete.",
                409
            )

        buffer = io.BytesIO()

        pdf = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=50,
            bottomMargin=50,
            title=(
                f"{filename} - AI Analysis"
            )
        )

        styles = (
            getSampleStyleSheet()
        )

        story = []

        story.append(
            Paragraph(
                "AI Document Analysis",
                styles["Title"]
            )
        )

        story.append(
            Spacer(
                1,
                8
            )
        )

        story.append(
            Paragraph(
                filename,
                styles["Heading2"]
            )
        )

        story.append(
            Spacer(
                1,
                18
            )
        )

        metadata = [
            [
                Paragraph(
                    "<b>File</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    str(filename),
                    styles["BodyText"]
                )
            ],
            [
                Paragraph(
                    "<b>Type</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    str(
                        file_type or ""
                    ).upper(),
                    styles["BodyText"]
                )
            ],
            [
                Paragraph(
                    "<b>Status</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    str(
                        status or ""
                    ),
                    styles["BodyText"]
                )
            ],
            [
                Paragraph(
                    "<b>Extracted Characters</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    f"{extracted_characters or 0:,}",
                    styles["BodyText"]
                )
            ],
            [
                Paragraph(
                    "<b>Uploaded</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    str(
                        upload_date or ""
                    ),
                    styles["BodyText"]
                )
            ],
            [
                Paragraph(
                    "<b>Analyzed</b>",
                    styles["BodyText"]
                ),
                Paragraph(
                    str(
                        analysis_date or "N/A"
                    ),
                    styles["BodyText"]
                )
            ]
        ]

        metadata_table = Table(
            metadata,
            colWidths=[
                140,
                360
            ]
        )

        metadata_table.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (0, -1),
                        colors.whitesmoke
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.5,
                        colors.lightgrey
                    ),
                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "TOP"
                    ),
                    (
                        "LEFTPADDING",
                        (0, 0),
                        (-1, -1),
                        8
                    ),
                    (
                        "RIGHTPADDING",
                        (0, 0),
                        (-1, -1),
                        8
                    ),
                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, -1),
                        7
                    ),
                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, -1),
                        7
                    )
                ]
            )
        )

        story.append(
            metadata_table
        )

        story.append(
            Spacer(
                1,
                22
            )
        )

        def add_section(
            title,
            paragraphs
        ):

            story.append(
                Paragraph(
                    title,
                    styles["Heading2"]
                )
            )

            story.append(
                Spacer(
                    1,
                    6
                )
            )

            if not paragraphs:

                story.append(
                    Paragraph(
                        "None identified.",
                        styles["BodyText"]
                    )
                )

            else:

                for paragraph in paragraphs:

                    story.append(
                        Paragraph(
                            paragraph,
                            styles["BodyText"]
                        )
                    )

                    story.append(
                        Spacer(
                            1,
                            5
                        )
                    )

            story.append(
                Spacer(
                    1,
                    14
                )
            )

        add_section(
            "Summary",
            [
                summary or
                "No summary available."
            ]
        )

        add_section(
            "Key Topics",
            [
                f"• {str(topic)}"
                for topic in key_topics
            ]
        )

        add_section(
            "Tags",
            [
                ", ".join(
                    str(tag)
                    for tag in tags
                )
            ] if tags else []
        )

        entity_lines = []

        for entity in entities:

            name = entity.get(
                "name",
                ""
            )

            entity_type = entity.get(
                "type",
                "other"
            )

            entity_lines.append(
                f"• <b>{name}</b> "
                f"({entity_type})"
            )

        add_section(
            "Entities",
            entity_lines
        )

        risk_lines = []

        for risk in risks:

            severity = risk.get(
                "severity",
                "Unknown"
            )

            risk_text = risk.get(
                "risk",
                ""
            )

            risk_lines.append(
                f"• <b>{severity}</b>: "
                f"{risk_text}"
            )

        add_section(
            "Risks",
            risk_lines
        )

        action_lines = []

        for item in action_items:

            action = item.get(
                "action",
                ""
            )

            owner = item.get(
                "owner"
            )

            due_date = item.get(
                "due_date"
            )

            line = (
                f"• {action}"
            )

            if owner:

                line += (
                    "<br/>&nbsp;&nbsp;&nbsp;"
                    f"<b>Owner:</b> {owner}"
                )

            if due_date:

                line += (
                    "<br/>&nbsp;&nbsp;&nbsp;"
                    f"<b>Due:</b> {due_date}"
                )

            action_lines.append(
                line
            )

        add_section(
            "Action Items",
            action_lines
        )

        date_lines = []

        for item in important_dates:

            date_value = item.get(
                "date",
                ""
            )

            description = item.get(
                "description",
                ""
            )

            date_lines.append(
                f"• <b>{date_value}</b>: "
                f"{description}"
            )

        add_section(
            "Important Dates",
            date_lines
        )

        story.append(
            Spacer(
                1,
                12
            )
        )

        story.append(
            Paragraph(
                "Generated by AI Document Analyzer.",
                styles["Italic"]
            )
        )

        pdf.build(
            story
        )

        buffer.seek(0)

        base_name = os.path.splitext(
            filename
        )[0]

        return send_file(
            buffer,
            as_attachment=True,
            download_name=(
                f"{base_name}_analysis.pdf"
            ),
            mimetype="application/pdf"
        )

    except Exception as exc:

        return (
            f"PDF export failed: {str(exc)}",
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
# Download Original
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

        cursor.execute(
            """
            SELECT
                FileName,
                BlobName

            FROM dbo.Documents

            WHERE
                DocumentID = ?;
            """,
            (
                document_id,
            )
        )

        row = cursor.fetchone()

        if not row:

            return (
                "Document not found.",
                404
            )

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
            io.BytesIO(
                file_data
            ),
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

        cursor.execute(
            """
            SELECT
                FileName,
                BlobName

            FROM dbo.Documents

            WHERE
                DocumentID = ?;
            """,
            (
                document_id,
            )
        )

        row = cursor.fetchone()

        if not row:

            flash(
                "Document not found.",
                "error"
            )

            return redirect(
                url_for(
                    "index"
                )
            )

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

        blob_client.delete_blob(
            delete_snapshots="include"
        )

        cursor.execute(
            """
            DELETE FROM dbo.DocumentAnalysis

            WHERE
                DocumentID = ?;
            """,
            (
                document_id,
            )
        )

        cursor.execute(
            """
            DELETE FROM dbo.Documents

            WHERE
                DocumentID = ?;
            """,
            (
                document_id,
            )
        )

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
        url_for(
            "index"
        )
    )


# ============================================================
# Upload Document
#
# IMPORTANT:
# This now ONLY stores the document and queues it.
# The WebJob handles extraction and AI analysis.
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
            url_for(
                "index"
            )
        )

    file = request.files[
        "file"
    ]

    if file.filename == "":

        flash(
            "No file was selected.",
            "error"
        )

        return redirect(
            url_for(
                "index"
            )
        )

    if not allowed_file(
        file.filename
    ):

        flash(
            "Unsupported file type.",
            "error"
        )

        return redirect(
            url_for(
                "index"
            )
        )

    original_filename = (
        secure_filename(
            file.filename
        )
    )

    extension = (
        original_filename
        .rsplit(
            ".",
            1
        )[1]
        .lower()
    )

    unique_blob_name = (
        f"{uuid.uuid4()}-"
        f"{original_filename}"
    )

    conn = None
    cursor = None

    blob_uploaded = False

    try:

        file_data = (
            file.read()
        )

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

        blob_uploaded = True

        conn = (
            get_db_connection()
        )

        cursor = (
            conn.cursor()
        )

        cursor.execute(
            """
            INSERT INTO dbo.Documents
            (
                FileName,
                BlobName,
                FileType,
                FileSizeBytes,
                Status
            )
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
                "Queued"
            )
        )

        conn.commit()

        flash(
            f"{original_filename} was uploaded "
            "and queued for analysis.",
            "success"
        )

    except Exception as exc:

        try:

            if conn:
                conn.rollback()

        except Exception:
            pass

        if blob_uploaded:

            try:

                blob_client.delete_blob(
                    delete_snapshots="include"
                )

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

    return redirect(
        url_for(
            "index"
        )
    )


# ============================================================
# Health Check
# ============================================================

@app.route(
    "/health"
)
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

        results[
            "database"
        ] = "connected"

    except Exception as exc:

        results[
            "database"
        ] = (
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

        results[
            "storage"
        ] = "connected"

    except Exception as exc:

        results[
            "storage"
        ] = (
            f"error: {str(exc)}"
        )

        status_code = 500

    try:

        get_openai_client()

        results[
            "openai"
        ] = "configured"

    except Exception as exc:

        results[
            "openai"
        ] = (
            f"error: {str(exc)}"
        )

        status_code = 500

    results[
        "status"
    ] = (
        "healthy"
        if status_code == 200
        else "unhealthy"
    )

    return (
        results,
        status_code
    )


# ============================================================
# Entry Point
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=8000
    )