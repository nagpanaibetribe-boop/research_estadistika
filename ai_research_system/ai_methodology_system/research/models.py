from django.db import models

class ResearchProject(models.Model):
    title = models.CharField(max_length=200)
    objective = models.TextField()
    problem = models.TextField()
    variable = models.CharField(max_length=100)
    values = models.TextField()

    created_at = models.DateTimeField(auto_now_add=True)
