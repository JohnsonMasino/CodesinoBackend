"""
agreement/models.py

No database models are required for the agreement generator —
the app is purely stateless (JSON in → .docx out).

If you later want to log/archive generated agreements, you can
uncomment and migrate the AgreementLog model below.
"""

# from django.db import models
#
# class AgreementLog(models.Model):
#     """Optional audit log of every generated agreement."""
#     reference       = models.CharField(max_length=30, unique=True)
#     client_name     = models.CharField(max_length=200, blank=True)
#     client_email    = models.EmailField(blank=True)
#     client_company  = models.CharField(max_length=200, blank=True)
#     project_name    = models.CharField(max_length=300, blank=True)
#     currency        = models.CharField(max_length=10, blank=True)
#     total_amount    = models.CharField(max_length=50, blank=True)
#     payment_plan    = models.CharField(max_length=5, blank=True)
#     generated_at    = models.DateTimeField(auto_now_add=True)
#     raw_payload     = models.JSONField(default=dict)
#
#     class Meta:
#         ordering = ["-generated_at"]
#         verbose_name = "Agreement Log"
#         verbose_name_plural = "Agreement Logs"
#
#     def __str__(self):
#         return f"{self.reference} — {self.client_name}"