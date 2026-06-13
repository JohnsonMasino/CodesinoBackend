"""
URL configuration for the agreement app.
File: agreement/urls.py

Include in your project's root urls.py:

    from django.urls import path, include

    urlpatterns = [
        path("api/", include("agreement.urls")),
        # ... your other routes
    ]

──────────────────────────────────────────────────────────────────────────────
AVAILABLE ENDPOINTS
──────────────────────────────────────────────────────────────────────────────

CLIENT PROJECT AGREEMENT
  POST  /api/generate-agreement/
        Generate a client project / service agreement between Codesino and
        an external client.  (Original endpoint — unchanged.)

EMPLOYMENT AGREEMENTS
  POST  /api/generate-executive-agreement/
        Employment agreement for executive and senior company roles.
        (CTO, Co-Founder, MD, Head of Engineering, Director-level, etc.)

  POST  /api/generate-support-agreement/
        Employment agreement for support and operations staff.
        (Customer Support, Social Media, Marketing, Admin, Sales, etc.)
        Supports multiple concurrent positions per employee.

  POST  /api/generate-technical-agreement/
        Employment agreement for technical / development team members.
        (Frontend Dev, Backend Dev, Full Stack, DevOps, UI/UX, Mobile, etc.)

All endpoints:
  • Accept:   Content-Type: application/json
  • Return:   application/vnd.openxmlformats-officedocument.wordprocessingml.document
  • Support:  CORS pre-flight (OPTIONS) with Access-Control-Allow-Origin: *
  • On error: JSON  { "error": "...", "detail": "..." }  with 400 or 500 status

See employment_views.py for the full list of accepted JSON body fields.
──────────────────────────────────────────────────────────────────────────────
"""

from django.urls import path

from . import views

urlpatterns = [
    # ── Original client project agreement ────────────────────────────────────
    path(
        "generate-agreement/",
        views.generate_agreement,
        name="generate_agreement",
    ),

    # ── Employment agreement — Executive / Senior roles ───────────────────────
    path(
        "generate-executive-agreement/",
        views.generate_executive_agreement,
        name="generate_executive_agreement",
    ),

    # ── Employment agreement — Support / Operations roles ─────────────────────
    path(
        "generate-support-agreement/",
        views.generate_support_agreement,
        name="generate_support_agreement",
    ),

    # ── Employment agreement — Technical / Dev Team roles ─────────────────────
    path(
        "generate-technical-agreement/",
        views.generate_technical_agreement,
        name="generate_technical_agreement",
    ),
]