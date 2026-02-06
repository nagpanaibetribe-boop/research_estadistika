from django import forms

class UploadForm(forms.Form):
    title = forms.CharField()
    objective = forms.CharField(widget=forms.Textarea)
    problem = forms.CharField(widget=forms.Textarea)
    variable = forms.CharField()
    file = forms.FileField()