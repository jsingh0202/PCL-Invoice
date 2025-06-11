from django.urls import path
from .views import generate, analysis, download, analysis_pdf

urlpatterns = [
    path("generate/", generate, name="generate"),
    path("analysis/", analysis, name="analysis"),
    path("analysis/pdf/", analysis_pdf, name="analysis_pdf"),
    path("download/", download, name="download"),
]
