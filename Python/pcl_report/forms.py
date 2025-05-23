from django import forms

class uploadFileForm(forms.Form):
    file = forms.FileField(label='Select a file')
    # Add any additional fields you need for your form
    # For example, if you want to include a text field:
    # text_field = forms.CharField(max_length=100, required=False)