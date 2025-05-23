from django.shortcuts import render, HttpResponse
from .forms import UploadFileForm
from .utils.generate import generate_backup
# Create your views here.
def generate(request):
    print(request)
    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            # Handle the uploaded file
            file = request.FILES['file']
            # Save the file or process it as needed
            # For example, you can save it to a specific location
           
            # Call the generate_report function
            backup = generate_backup(file)
            return render(request, "generate.html", {"backup": backup})
    else:
        form = UploadFileForm()
    return render(request, "upload.html", {"form": form})