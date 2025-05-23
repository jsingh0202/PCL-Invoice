from django.db import models

# Create your models here.
class Backup(models.Model):
    """
    Model to represent a backup.
    """
    name = models.CharField(max_length=255)
    date = models.DateTimeField(auto_now_add=True)
    size = models.IntegerField()
    status = models.CharField(max_length=50, choices=[('success', 'Success'), ('failure', 'Failure')])
    
    def __str__(self):
        return f"{self.name} - {self.status} - {self.date}"