import os
import sys
import time
import json
import traceback

# ============================================================
# WebJob Local Python Packages
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

LOCAL_PACKAGES = os.path.join(
    SCRIPT_DIR,
    ".python_packages",
    "lib",
    "site-packages"
)

if os.path.isdir(LOCAL_PACKAGES):
    sys.path.insert(0, LOCAL_PACKAGES)

# Make sure files packaged beside run.py are importable.
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)


# ============================================================
# Imports
# ============================================================

from document_processing import (
    get_db_connection,
    get_blob_service_client,
    extract_content,
    analyze_document_with_ai,
    save_analysis,
    STORAGE_CONTAINER_NAME
)


# ============================================================
# Configuration
# ============================================================

POLL_INTERVAL_SECONDS = 10


# ============================================================
# Claim Next Queued Document
# ============================================================

def claim_next_document():

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
        cursor = conn.cursor()

        # Find the oldest queued document.
        cursor.execute("""
            SELECT TOP 1
                DocumentID,
                FileName,
                BlobName
            FROM dbo.Documents
            WHERE Status = ?
            ORDER BY UploadDate ASC,
                     DocumentID ASC;
        """,
        (
            "Queued",
        ))

        row = cursor.fetchone()

        if not row:
            return None

        document_id = row[0]
        file_name = row[1]
        blob_name = row[2]

        # Claim the document.
        cursor.execute("""
            UPDATE dbo.Documents
            SET Status = ?
            WHERE DocumentID = ?
              AND Status = ?;
        """,
        (
            "Processing",
            document_id,
            "Queued"
        ))

        conn.commit()

        # If nothing was updated, another worker got it.
        if cursor.rowcount == 0:
            return None

        return {
            "DocumentID": document_id,
            "FileName": file_name,
            "BlobName": blob_name
        }

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
# Mark Document Failed
# ============================================================

def mark_document_failed(document_id):

    conn = None
    cursor = None

    try:

        conn = get_db_connection()
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
# Process Document
# ============================================================

def process_document(document):

    document_id = document["DocumentID"]
    file_name = document["FileName"]
    blob_name = document["BlobName"]

    conn = None
    cursor = None

    print("")
    print("============================================================")
    print(f"Processing DocumentID: {document_id}")
    print(f"File: {file_name}")
    print(f"Blob: {blob_name}")
    print("============================================================")

    try:

        # ====================================================
        # Download Blob
        # ====================================================

        blob_service_client = get_blob_service_client()

        blob_client = blob_service_client.get_blob_client(
            container=STORAGE_CONTAINER_NAME,
            blob=blob_name
        )

        file_data = (
            blob_client
            .download_blob()
            .readall()
        )

        print(
            f"Downloaded {len(file_data):,} bytes."
        )


        # ====================================================
        # Extract Document Content
        # ====================================================

        extracted_text = extract_content(
            file_name,
            file_data
        )

        if not extracted_text:
            raise RuntimeError(
                "No text could be extracted from the document."
            )

        if not extracted_text.strip():
            raise RuntimeError(
                "Extracted document text was empty."
            )

        print(
            f"Extracted {len(extracted_text):,} characters."
        )


        # ====================================================
        # Save Extracted Text
        # ====================================================

        conn = get_db_connection()
        cursor = conn.cursor()

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

        print("Extracted text saved.")


        # ====================================================
        # AI Analysis
        # ====================================================

        print("Sending document to AI analysis...")

        analysis = analyze_document_with_ai(
            extracted_text
        )

        print("AI analysis completed.")


        # ====================================================
        # Save Analysis
        # ====================================================

        save_analysis(
            cursor,
            document_id,
            analysis
        )

        conn.commit()

        print("Analysis saved.")


        # ====================================================
        # Mark Document Analyzed
        # ====================================================

        cursor.execute("""
            UPDATE dbo.Documents
            SET Status = ?
            WHERE DocumentID = ?;
        """,
        (
            "Analyzed",
            document_id
        ))

        conn.commit()

        print(
            f"DocumentID {document_id} completed successfully."
        )

    except Exception as exc:

        print("")
        print("DOCUMENT PROCESSING FAILED")
        print(
            f"DocumentID: {document_id}"
        )
        print(
            f"Error: {str(exc)}"
        )

        traceback.print_exc()

        try:

            mark_document_failed(
                document_id
            )

        except Exception as status_exc:

            print(
                "Unable to mark document as failed:"
            )

            print(
                str(status_exc)
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
# Main Worker Loop
# ============================================================

def main():

    print("")
    print("============================================================")
    print("Document Analyzer WebJob")
    print("============================================================")
    print(
        f"Python: {sys.version}"
    )
    print(
        f"Script directory: {SCRIPT_DIR}"
    )
    print(
        f"Local packages: {LOCAL_PACKAGES}"
    )
    print(
        f"Local packages exist: "
        f"{os.path.isdir(LOCAL_PACKAGES)}"
    )
    print(
        f"Polling every "
        f"{POLL_INTERVAL_SECONDS} seconds."
    )
    print("============================================================")
    print("")

    while True:

        try:

            document = claim_next_document()

            if document:

                process_document(
                    document
                )

            else:

                time.sleep(
                    POLL_INTERVAL_SECONDS
                )

        except KeyboardInterrupt:

            print(
                "WebJob stopped."
            )

            break

        except Exception as exc:

            print("")
            print("WORKER LOOP ERROR")
            print(
                str(exc)
            )

            traceback.print_exc()

            time.sleep(
                POLL_INTERVAL_SECONDS
            )


# ============================================================
# Entry Point
# ============================================================

if __name__ == "__main__":
    main()