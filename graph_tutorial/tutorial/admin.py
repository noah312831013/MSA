from django.contrib import admin
from .models import AutoScheduleMeeting, TaskNotification

# 註冊模型
admin.site.register(AutoScheduleMeeting)
admin.site.register(TaskNotification)