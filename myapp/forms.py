from django import forms

class UploadFileForm(forms.Form):
    gst_file = forms.FileField(label='GST Excel File')
    tally_file = forms.FileField(label='Tally Excel File')
    