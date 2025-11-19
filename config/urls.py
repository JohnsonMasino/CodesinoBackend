from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/', include('contact.urls')),
    path('api/', include('service.urls')),
    path('api/', include('newsletter.urls')),
    path('blog/', include('blog.urls')),
]

