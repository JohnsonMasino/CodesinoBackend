"""
Codesino Agreement Generator — Django Views
============================================
File: agreement/views.py

Four agreement endpoints:

  POST /api/generate-agreement/              →  Client/Project Agreement (original)
  POST /api/generate-executive-agreement/    →  Executive / Senior Roles
  POST /api/generate-support-agreement/      →  Support & Operations Roles
  POST /api/generate-technical-agreement/    →  Technical / Dev Team Roles

Each endpoint accepts a JSON body, generates a .docx file, and streams
it back as a download attachment.
"""

import json
from datetime import datetime

from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods

from .generator import (
    generate_agreement_buffer,
    generate_executive_agreement_buffer,
    generate_support_agreement_buffer,
    generate_technical_agreement_buffer,
)


# ─────────────────────────────────────────────────────────────────────────────
# INTERNAL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _cors_preflight():
    """Return a 200 OK response for CORS OPTIONS pre-flight requests."""
    response = HttpResponse()
    response["Access-Control-Allow-Origin"]  = "*"
    response["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    response["Access-Control-Allow-Headers"] = "Content-Type"
    return response


def _add_cors(response):
    """Attach CORS headers to any outgoing response."""
    response["Access-Control-Allow-Origin"] = "*"
    return response


def _parse_body(request):
    """
    Parse the JSON request body.
    Returns (data_dict, error_response).
    If parsing succeeds, error_response is None.
    If parsing fails, data_dict is None and error_response is a JsonResponse.
    """
    try:
        data = json.loads(request.body)
        if not isinstance(data, dict):
            raise ValueError("Request body must be a JSON object.")
        return data, None
    except (json.JSONDecodeError, ValueError) as exc:
        return None, _add_cors(
            JsonResponse({"error": "Invalid JSON body", "detail": str(exc)}, status=400)
        )


def _build_docx_response(buffer: bytes, name: str, doc_type: str) -> HttpResponse:
    """
    Wrap raw .docx bytes in an HTTP download response with a descriptive filename.

    doc_type examples: "Agreement", "Executive_Agreement",
                       "Support_Agreement", "Technical_Agreement"
    """
    safe_name = (name or "Unknown").replace(" ", "_")
    date_str  = datetime.now().strftime("%Y-%m-%d")
    filename  = f"Codesino_{doc_type}_{safe_name}_{date_str}.docx"

    response = HttpResponse(
        buffer,
        content_type=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return _add_cors(response)


# ─────────────────────────────────────────────────────────────────────────────
# ENDPOINT 1 — CLIENT / PROJECT AGREEMENT  (original)
# POST /api/generate-agreement/
# ─────────────────────────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def generate_agreement(request):
    """
    POST /api/generate-agreement/
    Body: JSON with all form data fields (clientName, …)
    Returns: .docx file download
    """
    if request.method == "OPTIONS":
        return _cors_preflight()

    data, err = _parse_body(request)
    if err:
        return err

    try:
        buffer = generate_agreement_buffer(data)
    except Exception as exc:
        return _add_cors(
            JsonResponse(
                {"error": "Failed to generate agreement", "detail": str(exc)},
                status=500,
            )
        )

    return _build_docx_response(
        buffer,
        name=data.get("clientName", ""),
        doc_type="Agreement",
    )


# ─────────────────────────────────────────────────────────────────────────────
# ENDPOINT 2 — EXECUTIVE / SENIOR ROLES
# POST /api/generate-executive-agreement/
# ─────────────────────────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def generate_executive_agreement(request):
    """
    Generate an employment agreement for executive and senior staff.

    Suitable for: CTO, Co-Founder, Managing Director, Head of Engineering,
    Director-level roles, Senior Project Managers, and similar senior positions.

    The document includes:
      • Personalised congratulatory letter of appointment
      • Full employment & compensation details
      • Key responsibilities for senior leaders
      • Comprehensive company notice covering:
          – Company policy acceptance
          – Company-information confidentiality
          – Client-information confidentiality
          – Professional conduct & ethics
          – Non-solicitation & non-compete clause
          – Additional rules
      • Acceptance statement before the signature block
      • Intellectual property, termination, and dispute-resolution clauses
      • Dual signature block with company stamp
    """
    if request.method == "OPTIONS":
        return _cors_preflight()

    data, err = _parse_body(request)
    if err:
        return err

    try:
        buffer = generate_executive_agreement_buffer(data)
    except Exception as exc:
        return _add_cors(
            JsonResponse(
                {"error": "Failed to generate executive agreement", "detail": str(exc)},
                status=500,
            )
        )

    return _build_docx_response(
        buffer,
        name=data.get("employeeName", ""),
        doc_type="Executive_Agreement",
    )


# ─────────────────────────────────────────────────────────────────────────────
# ENDPOINT 3 — SUPPORT / OPERATIONS ROLES
# POST /api/generate-support-agreement/
# ─────────────────────────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def generate_support_agreement(request):
    """
    Generate an employment agreement for support and operations staff.

    Suitable for: Customer Support Representatives, Social Media Managers,
    Community Managers, Marketing Associates, Administrative Assistants,
    Sales Representatives, Virtual Assistants, and similar roles.

    Employees in this category may hold more than one role simultaneously.
    Pass all applicable roles as an array in the `positions` field.

    The document includes:
      • Personalised congratulatory letter of appointment
      • Full employment details with multi-position selection
      • Flexible compensation block (supports monthly salary, commission %,
        hourly rate, retainer, etc.)
      • General support responsibilities
      • Customer-support-specific conduct standards
      • Social media & content role-specific responsibilities
      • Company notice covering:
          – Policy acceptance
          – Company confidentiality
          – Client confidentiality
          – Customer interaction standards
          – Online / social media conduct
      • Acceptance statement before the signature block
      • Standard IP, termination, and dispute-resolution clauses
      • Dual signature block with company stamp
    """
    if request.method == "OPTIONS":
        return _cors_preflight()

    data, err = _parse_body(request)
    if err:
        return err

    try:
        buffer = generate_support_agreement_buffer(data)
    except Exception as exc:
        return _add_cors(
            JsonResponse(
                {"error": "Failed to generate support agreement", "detail": str(exc)},
                status=500,
            )
        )

    return _build_docx_response(
        buffer,
        name=data.get("employeeName", ""),
        doc_type="Support_Agreement",
    )


# ─────────────────────────────────────────────────────────────────────────────
# ENDPOINT 4 — TECHNICAL / DEV TEAM ROLES
# POST /api/generate-technical-agreement/
# ─────────────────────────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def generate_technical_agreement(request):
    """
    Generate an employment agreement for technical and development team members.

    Suitable for: Frontend Developers, Backend Developers, Full Stack Developers,
    Mobile Developers, UI/UX Designers, DevOps Engineers, Cloud Engineers,
    Database Administrators, Cybersecurity Engineers, ML Engineers, and similar.

    The document includes:
      • Personalised congratulatory letter of appointment
      • Full employment details with technical stack / tools section
      • Flexible compensation block
      • General engineering responsibilities
      • Frontend developer specific standards:
          – Responsive, modern UI/UX requirements
          – Animation, performance, and accessibility standards
      • Backend developer specific standards:
          – Security-first engineering requirements
          – API, database, and testing standards
      • DevOps / infrastructure specific standards
      • Technical company notice covering:
          – Policy acceptance
          – Source code & codebase confidentiality
          – Company & client data confidentiality (including breach reporting)
          – Security standards & engineering ethics
          – Professional conduct & collaboration standards
      • Acceptance statement before the signature block
      • Standard IP, termination, and dispute-resolution clauses
      • Dual signature block with company stamp
    """
    if request.method == "OPTIONS":
        return _cors_preflight()

    data, err = _parse_body(request)
    if err:
        return err

    try:
        buffer = generate_technical_agreement_buffer(data)
    except Exception as exc:
        return _add_cors(
            JsonResponse(
                {"error": "Failed to generate technical agreement", "detail": str(exc)},
                status=500,
            )
        )

    return _build_docx_response(
        buffer,
        name=data.get("employeeName", ""),
        doc_type="Technical_Agreement",
    )