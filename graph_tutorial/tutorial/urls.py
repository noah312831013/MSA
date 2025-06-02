# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from django.urls import path

from . import views

urlpatterns = [
  # /
  path('', views.home, name='home'),
  path('signin', views.sign_in, name='signin'),
  path('signout', views.sign_out, name='signout'),
  path('calendar', views.calendar, name='calendar'),
  path('callback', views.callback, name='callback'),
  path('calendar/new', views.new_event, name='newevent'),
  path('auto-schedule-meeting', views.schedule_meeting, name='auto_schedule_meeting'),
  path('webhook/response/', views.meeting_response, name='meeting_response'),
  path('meeting-status/<uuid:meeting_uuid>/', views.meeting_status, name='meeting_status'),
  path('api/contactors/', views.get_contacts, name='get_contacts'),
  path("sharepoint-reminder/", views.sharepoint_reminder_dashboard, name="sharepoint_reminder_dashboard"),
  path("start-tasks/", views.start_tasks, name="start_tasks"),
  path("stop-tasks/", views.stop_tasks, name="stop_tasks"),
  path("get-tracking-items/", views.get_tracking_items, name="get_tracking_items"),
]
