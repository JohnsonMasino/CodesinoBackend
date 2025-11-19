from django.contrib import admin
from .models import ServiceRequest

@admin.register(ServiceRequest)
class ServiceRequestAdmin(admin.ModelAdmin):
    list_display = ("name", "email", "message", "submitted_at")
    ordering = ("-submitted_at",)
    search_fields = ("name", "email")
