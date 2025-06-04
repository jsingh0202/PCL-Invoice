from django.db import models
import uuid
import os


# Create your models here.
class Backup(models.Model):
    """
    Model to represent a backup.
    """

    name = models.CharField(max_length=255)
    date = models.DateTimeField(auto_now_add=True)
    size = models.IntegerField()
    status = models.CharField(
        max_length=50, choices=[("success", "Success"), ("failure", "Failure")]
    )

    def __str__(self):
        return f"{self.name} - {self.status} - {self.date}"


class Export(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    file_path = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    ip_address = models.GenericIPAddressField(null=True)

    @property
    def filename(self):
        return os.path.basename(self.file_path)

    class Meta:
        ordering = ["-created_at"]
