import os

from flask import Flask, render_template
from mssql_python import connect

app = Flask(__name__)

CONNECTION_STRING = os.getenv("AZURE_SQL_CONNECTIONSTRING")


def get_db_connection():
    if not CONNECTION_STRING:
        raise RuntimeError(
            "AZURE_SQL_CONNECTIONSTRING environment variable is not configured."
        )

    return connect(CONNECTION_STRING)


@app.route("/")
def index():
    endpoints = []
    error_message = None

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                EndpointID,
                Name,
                Url,
                ExpectedStatus,
                CheckInterval,
                IsEnabled,
                CreatedDate
            FROM dbo.MonitoredEndpoints
            ORDER BY Name;
        """)

        rows = cursor.fetchall()

        for row in rows:
            endpoints.append({
                "EndpointID": row[0],
                "Name": row[1],
                "Url": row[2],
                "ExpectedStatus": row[3],
                "CheckInterval": row[4],
                "IsEnabled": row[5],
                "CreatedDate": row[6],
            })

        cursor.close()
        conn.close()

    except Exception as exc:
        error_message = str(exc)

    return render_template(
        "index.html",
        endpoints=endpoints,
        error_message=error_message
    )


@app.route("/health")
def health():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT 1;")
        cursor.fetchone()

        cursor.close()
        conn.close()

        return {
            "status": "healthy",
            "database": "connected"
        }, 200

    except Exception as exc:
        return {
            "status": "unhealthy",
            "database": "disconnected",
            "error": str(exc)
        }, 500


if __name__ == "__main__":
    app.run()