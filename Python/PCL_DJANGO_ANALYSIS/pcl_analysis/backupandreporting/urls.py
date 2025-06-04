from django.urls import path
from .views import generate, analysis, download

urlpatterns = [
    path("generate/", generate, name="generate"),
    path("analysis/", analysis, name="analysis"),
    path("download/", download, name="download"),
]
