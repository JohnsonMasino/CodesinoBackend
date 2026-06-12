"""
Codesino Agreement Generator — Django View
==========================================
File: agreement/views.py

Handles POST /api/generate-agreement/
Receives JSON form data and returns a .docx file as a download.
"""

import json
import math
import random
from datetime import datetime
from io import BytesIO

from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods

from .generator import generate_agreement_buffer


@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def generate_agreement(request):
    """
    POST /api/generate-agreement/
    Body: JSON with all form data fields
    Returns: .docx file download
    """

    # Handle CORS preflight
    if request.method == "OPTIONS":
        response = HttpResponse()
        response["Access-Control-Allow-Origin"] = "*"
        response["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    try:
        data = json.loads(request.body)
    except (json.JSONDecodeError, Exception) as e:
        return JsonResponse({"error": "Invalid JSON body", "detail": str(e)}, status=400)

    try:
        buffer = generate_agreement_buffer(data)
    except Exception as e:
        return JsonResponse({"error": "Failed to generate agreement", "detail": str(e)}, status=500)

    client_name = (data.get("clientName") or "Client").replace(" ", "_")
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"Codesino_Agreement_{client_name}_{date_str}.docx"

    response = HttpResponse(
        buffer,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    response["Access-Control-Allow-Origin"] = "*"
    return response