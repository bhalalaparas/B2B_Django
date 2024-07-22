from django.db import models

# Create your models here.

class image(models.Model):
    index = models.IntegerField()

class video(models.Model):
    index = models.IntegerField()
