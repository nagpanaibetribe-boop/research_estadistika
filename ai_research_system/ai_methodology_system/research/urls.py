from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("analyze/", views.analyze, name="analyze"),
    path("export-word/", views.export_word, name="export_word"),  # ← THIS IS IMPORTANT
]
