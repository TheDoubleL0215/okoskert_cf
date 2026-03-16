import json
import os
from datetime import datetime, timedelta

from firebase_functions import https_fn
from firebase_admin import initialize_app, firestore, storage
import google.cloud.firestore
from google.cloud.firestore import FieldFilter

from export_excel import build_export_xlsx

os.environ.pop("FIRESTORE_EMULATOR_HOST", None)
initialize_app()


def _serialize_value(v):
    """Firestore to_dict() értékek JSON-barát formára (datetime, ref stb.)."""
    if v is None:
        return None
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if hasattr(v, "path"):  # DocumentReference
        return v.path
    if isinstance(v, dict):
        return _serialize_dict(v)
    if isinstance(v, list):
        return [_serialize_value(x) for x in v]
    return v


def _serialize_dict(d):
    if d is None:
        return None
    return {k: _serialize_value(v) for k, v in d.items()}


@https_fn.on_request()
def projectExport(req: https_fn.Request) -> https_fn.Response:
    projectId = req.args.get("projectId")
    if not projectId:
        return https_fn.Response(
            "No projectId parameter provided",
            status=400,
        )

    db: google.cloud.firestore.Client = firestore.client()
    project_ref = db.collection("projects").document(projectId)
    project_snapshot = project_ref.get()
    if not project_snapshot.exists:
        return https_fn.Response("Project not found", status=404)

    project_dict = project_snapshot.to_dict()
    team_id = project_dict.get("teamId")
    if not team_id:
        return https_fn.Response("Project has no teamId", status=400)

    worklog_query = db.collection_group("worklogs").where(
        filter=FieldFilter("assignedProjectId", "==", projectId)
    )
    material_query = db.collection_group("materials").where(
        filter=FieldFilter("projectId", "==", projectId)
    )
    users_query = db.collection("users").where(
        filter=FieldFilter("teamId", "==", team_id)
    )
    machines_query = db.collection("machines").where(
        filter=FieldFilter("teamId", "==", team_id)
    )
    machine_worklog_ref = db.collection("projects").document(projectId).collection("machineWorklog")

    worklog_items = []
    for doc in worklog_query.stream():
        workspace_id = doc.reference.parent.parent.id if doc.reference.parent else None
        worklog_items.append({"id": doc.id, "workspaceId": workspace_id, **doc.to_dict()})

    material_items = [{"id": doc.id, **doc.to_dict()} for doc in material_query.stream()]
    users_items = [{"id": doc.id, **doc.to_dict()} for doc in users_query.stream()]
    machines_items = [{"id": doc.id, **doc.to_dict()} for doc in machines_query.stream()]
    machine_worklog_items = [
        {"id": doc.id, **doc.to_dict()}
        for doc in machine_worklog_ref.stream()
    ]

    export_data = {
        "project": _serialize_dict(project_dict),
        "worklog": [_serialize_dict(x) for x in worklog_items],
        "material": [_serialize_dict(x) for x in material_items],
        "users": [_serialize_dict(x) for x in users_items],
        "machines": [_serialize_dict(x) for x in machines_items],
        "machineWorklog": [_serialize_dict(x) for x in machine_worklog_items],
    }

    try:
        xlsx_bytes = build_export_xlsx(export_data)
    except Exception as e:
        return https_fn.Response(
            json.dumps({"error": "Excel export failed", "detail": str(e)}),
            status=500,
            headers={"Content-Type": "application/json; charset=utf-8"},
        )

    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    file_name = f"projekt_jelentes_{timestamp}.xlsx"
    storage_path = f"exports/{projectId}/{file_name}"

    try:
        bucket = storage.bucket()
        blob = bucket.blob(storage_path)
        blob.upload_from_string(
            xlsx_bytes,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return https_fn.Response(
            json.dumps({"error": "Storage upload failed", "detail": str(e)}),
            status=500,
            headers={"Content-Type": "application/json; charset=utf-8"},
        )

    download_url = None
    try:
        download_url = blob.generate_signed_url(
            expiration=timedelta(hours=1),
            method="GET",
        )
    except Exception:
        # Lokális/emulator: nincs privát kulcs a credentialban, signed URL nem lehet.
        # A kliens a storagePath-tal a Firebase Storage SDK getDownloadURL() használatával lekérheti az URL-t.
        pass

    payload = {
        "fileName": file_name,
        "storagePath": storage_path,
    }
    if download_url:
        payload["downloadUrl"] = download_url

    return https_fn.Response(
        json.dumps(payload, ensure_ascii=False),
        status=200,
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
