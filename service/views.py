from rest_framework import viewsets
from .models import ServiceRequest
from .serializers import ServiceRequestSerializer

class ServiceRequestViewSet(viewsets.ModelViewSet):
    queryset = ServiceRequest.objects.all().order_by('-submitted_at')
    serializer_class = ServiceRequestSerializer

