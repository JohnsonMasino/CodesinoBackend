"""
URL configuration for the agreement app.
File: agreement/urls.py

Include in your project's urls.py:
    path('api/', include('agreement.urls')),
"""

from django.urls import path
from . import views

urlpatterns = [
    path("generate-agreement/", views.generate_agreement, name="generate_agreement"),
]