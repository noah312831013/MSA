from django.core.management.base import BaseCommand
from tutorial.models import SharePointClientConfig

class Command(BaseCommand):
    help = "Seed the default SharePointClientConfig data"

    def handle(self, *args, **kwargs):
        if not SharePointClientConfig.objects.exists():
            SharePointClientConfig.objects.create(
                drive_name="ScrumSprints",
                file_path="Feature to do list+Q&A/[19.10] Mx Feature_to do list+ Q&A.xlsx",
                routine_interval=60*60*24,  # Default to 60 seconds
                polling_interval=30,  # Default to 30 seconds
                is_active=False
            )
            self.stdout.write(self.style.SUCCESS("Default SharePointClientConfig created."))
        else:
            self.stdout.write(self.style.WARNING("SharePointClientConfig already exists."))