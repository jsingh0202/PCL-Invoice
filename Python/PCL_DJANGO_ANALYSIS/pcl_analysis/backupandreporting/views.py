from bs4 import BeautifulSoup
from bs4.element import Tag
from django.http import FileResponse
from django.shortcuts import render, HttpResponse
from io import BytesIO
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

from .utils.generate.wr_export import generate_export
from .utils.analysis.export import analyze

import os
import tempfile
import uuid


def generate(request):
    if request.method == "POST":
        file = request.FILES.get("file")
        if file:
            # Generate Excel
            workbook, date = generate_export(file)
            request.session["date"] = date
            
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
        date = request.session.get("date")
        return FileResponse(
            open(file_path, "rb"),
            filename=f"{date} Export.xlsx",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    return HttpResponse("File not found", status=404)


def analysis(request):
    file_path = request.GET.get("file")
    if file_path and os.path.exists(file_path):
        results = analyze(file_path)  # This returns the reporting_values dict
        request.session["analysis_results"] = results
        return render(request, "analysis.html", {"reporting": results})
    return HttpResponse("File not found", status=404)


def analysis_pdf(request):
    results = request.session.get("analysis_results")
    if not results:
        return HttpResponse("No analysis results found in session.", status=404)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    styles = getSampleStyleSheet()
    elements = []

    elements.append(
        Paragraph("SMA Consulting PCL Invoice Analysis Report", styles["Title"])
    )
    elements.append(Spacer(1, 18))

    for idx, (title, table_html) in enumerate(results.items()):
        elements.append(Paragraph(str(title), styles["Heading2"]))
        # Parse HTML table to data
        soup = BeautifulSoup(table_html, "html.parser")
        table = soup.find("table")
        data = []
        if isinstance(table, Tag):
            # Get headers
            header_style = ParagraphStyle(
                "header_style",
                parent=styles["Normal"],
                alignment=1,  # Center
                fontName="Helvetica-Bold",
                fontSize=8,
                textColor=colors.white,
                leading=10,
            )

            cell_style = ParagraphStyle(
                "cell_style",
                parent=styles["Normal"],
                alignment=1,  # Center
                fontName="Helvetica",
                fontSize=8,
                leading=10,
            )

            headers = [
                Paragraph(th.get_text(strip=True), header_style)
                for th in table.find_all("th")
            ]
            if headers:
                data.append(headers)
            # Get rows
            for row in table.find_all("tr"):
                if isinstance(row, Tag):
                    cells = row.find_all("td")
                    if cells:
                        data.append(
                            [
                                Paragraph(cell.get_text(strip=True), cell_style)
                                for cell in cells
                            ]
                        )
        else:
            # fallback: just show as text
            data = [[table_html]]

        # Style the table
        tbl_style = TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0d6efd")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ]
        )
        tbl = Table(data, repeatRows=1)
        tbl.setStyle(tbl_style)
        elements.append(tbl)
        elements.append(Spacer(1, 18))
        if idx < len(results) - 1:
            elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)
    
    date = request.session.get("date")
    return FileResponse(buffer, as_attachment=True, filename=f"{date} Report.pdf")
