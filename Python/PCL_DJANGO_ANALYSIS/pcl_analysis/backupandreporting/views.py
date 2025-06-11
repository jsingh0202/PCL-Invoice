import tempfile
import uuid
from django.http import FileResponse
from django.shortcuts import render, HttpResponse

from .utils.generate.wr_export import generate_export
from .utils.analysis.export import analyze
import os


def generate(request):
    if request.method == "POST":
        file = request.FILES.get("file")
        if file:
            # Generate Excel
            workbook = generate_export(file)

            # Save to temp file with unique name
            temp_dir = tempfile.gettempdir()
            export_id = uuid.uuid4()
            file_path = os.path.join(temp_dir, f"export_{export_id}.xlsx")
            workbook.save(file_path)

            return render(
                request,
                "generate.html",
                {
                    "download_url": f"/download/?file={file_path}",
                    "analyze_url": f"/analysis/?file={file_path}",
                },
            )

    return render(request, "upload.html")


def download(request):
    file_path = request.GET.get("file")
    if file_path and os.path.exists(file_path):
        return FileResponse(
            open(file_path, "rb"),
            filename="export.xlsx",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    return HttpResponse("File not found", status=404)


def analysis(request):
    file_path = request.GET.get("file")
    if file_path and os.path.exists(file_path):
        results = analyze(file_path)  # This returns the reporting_values dict
        return render(request, "analysis.html", {"reporting": results})
    return HttpResponse("File not found", status=404)
