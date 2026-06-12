from django.contrib import admin

# Register your models here.
"""
agreement/admin.py

No database models are registered — the agreement generator is stateless
(JSON in → .docx out), so there is nothing to expose in Django admin.

If you uncomment the AgreementLog model in models.py, register it here:

from django.contrib import admin
from .models import AgreementLog

@admin.register(AgreementLog)
class AgreementLogAdmin(admin.ModelAdmin):
    list_display = ("reference", "client_name", "client_company", "project_name",
                     "currency", "total_amount", "payment_plan", "generated_at")
    list_filter = ("currency", "payment_plan", "generated_at")
    search_fields = ("reference", "client_name", "client_email", "client_company", "project_name")
    readonly_fields = ("generated_at",)
"""