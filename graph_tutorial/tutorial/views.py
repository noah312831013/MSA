# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
from datetime import datetime, timedelta
from django.shortcuts import render
from django.http import HttpResponseRedirect, HttpResponseBadRequest, JsonResponse
from django.urls import reverse
from django.contrib import messages
from dateutil import tz, parser
from tutorial.auth_helper import (get_sign_in_flow, get_token_from_code, store_user,
    remove_user_and_token, get_token)
from tutorial.graph_helper import get_user, get_iana_from_windows, get_calendar_events, create_event, get_meeting_times_slots, get_user_info, get_chat_ids, get_users, inform_attendees
from .models import AutoScheduleMeeting
import uuid

def initialize_context(request):
    context = {}

    # Check for any errors in the session
    error = request.session.pop('flash_error', None)

    if error is not None:
        context['errors'] = []
        context['errors'].append(error)

    # Check for user in the session
    context['user'] = request.session.get('user', {'is_authenticated': False})
    return context

def home(request):
    context = initialize_context(request)
    return render(request, 'tutorial/home.html', context)

def sign_in(request):
    # Get the sign-in flow
    flow = get_sign_in_flow()
    # Save the expected flow so we can use it in the callback
    request.session['auth_flow'] = flow

    # Redirect to the Azure sign-in page
    return HttpResponseRedirect(flow['auth_uri'])

def callback(request):
    # Make the token request
    result = get_token_from_code(request)

    #Get the user's profile
    user = get_user(result['access_token'])

    # Store user
    store_user(request, user)
    return HttpResponseRedirect(reverse('home'))

def sign_out(request):
    # Clear out the user and token
    remove_user_and_token(request)

    return HttpResponseRedirect(reverse('home'))

def calendar(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))

    # Load the user's time zone
    # Microsoft Graph can return the user's time zone as either
    # a Windows time zone name or an IANA time zone identifier
    # Python datetime requires IANA, so convert Windows to IANA
    time_zone = get_iana_from_windows(user['timeZone'])
    tz_info = tz.gettz(time_zone)

    # Get midnight today in user's time zone
    today = datetime.now(tz_info).replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0)

    # Based on today, get the start of the week (Sunday)
    if today.weekday() != 6:
        start = today - timedelta(days=today.isoweekday())
    else:
        start = today

    end = start + timedelta(days=7)

    token = get_token(request)

    events = get_calendar_events(
        token,
        start.isoformat(timespec='seconds'),
        end.isoformat(timespec='seconds'),
        user['timeZone'])

    if events:
        # Convert the ISO 8601 date times to a datetime object
        # This allows the Django template to format the value nicely
        for event in events['value']:
            event['start']['dateTime'] = parser.parse(event['start']['dateTime'])
            event['end']['dateTime'] = parser.parse(event['end']['dateTime'])

        context['events'] = events['value']

    return render(request, 'tutorial/calendar.html', context)

def new_event(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))

    if request.method == 'POST':
        # Validate the form values
        # Required values
        if (not request.POST['ev-subject']) or \
            (not request.POST['ev-start']) or \
            (not request.POST['ev-end']):
            context['errors'] = [
                {
                    'message': 'Invalid values',
                    'debug': 'The subject, start, and end fields are required.'
                }
            ]
            return render(request, 'tutorial/newevent.html', context)

        attendees = None
        if request.POST['ev-attendees']:
            attendees = request.POST['ev-attendees'].split(';')

        # Create the event
        token = get_token(request)

        create_event(
          token,
          request.POST['ev-subject'],
          request.POST['ev-start'],
          request.POST['ev-end'],
          attendees,
          request.POST['ev-body'],
          user['timeZone'])

        # Redirect back to calendar view
        return HttpResponseRedirect(reverse('calendar'))
    else:
        # Render the form
        return render(request, 'tutorial/newevent.html', context)

def schedule_meeting(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))
    token = get_token(request)
    if request.method == 'POST':
        # 獲取時區信息
        time_zone = get_iana_from_windows(user['timeZone'])
        tz_info = tz.gettz(time_zone)
        
        # 解析並添加時區信息到日期時間
        start_time = parser.parse(request.POST.get('start_time'))
        end_time = parser.parse(request.POST.get('end_time'))
        
        # 如果日期時間沒有時區信息，添加用戶的時區
        if start_time.tzinfo is None:
            start_time = start_time.replace(tzinfo=tz_info)
        if end_time.tzinfo is None:
            end_time = end_time.replace(tzinfo=tz_info)

        # 獲取與會者信息
        attendees = request.POST.getlist('attendees')
        token = get_token(request)
        user_ids = [get_user_info(token, email)['id'] for email in attendees]
        chat_ids = get_chat_ids(token, user_ids)
        
        # 創建新的排程記錄
        meeting = AutoScheduleMeeting(
            title=request.POST.get('title'),
            description=request.POST.get('description'),
            duration=int(request.POST.get('duration')),
            start_time=start_time,
            end_time=end_time,
            status='pending'
        )
        
        # 設置與會者
        meeting.set_attendees(attendees, user_ids, chat_ids)
        meeting.host_email = get_user(token)['mail']
        meeting.time_zone = time_zone

        # 嘗試獲取候選時間
        try:
            time_slots = get_meeting_times_slots(token, meeting, time_zone)
        except Exception as e:
            messages.error(request, "cannot find available time for meeting for all attendees")
            return render(request, 'tutorial/auto_schedule_meeting.html', context)

        if not time_slots:
            messages.error(request, "cannot find available time for meeting for all attendees")
            return render(request, 'tutorial/auto_schedule_meeting.html', context)
        
        # 設定候選時間並儲存
        meeting.set_candidate_times(time_slots)
        meeting.status = 'waiting'
        meeting.save()

        inform_attendees(token, meeting)
        context['meeting'] = meeting
        return render(request, 'tutorial/auto_schedule_meeting_progress.html', context)

    return render(request, 'tutorial/auto_schedule_meeting.html', context)


# 用來處理會議回應的 webhook
def meeting_response(request):
    tenant_id = request.GET.get('tenantId')
    uuid_str = request.GET.get('uuid')
    response_status = request.GET.get('response')

    if not tenant_id or not uuid_str or not response_status:
        return HttpResponseBadRequest("Missing parameters")

    try:
        meeting_uuid = uuid.UUID(uuid_str)
        meeting = AutoScheduleMeeting.objects.get(uuid=meeting_uuid)
    except (ValueError, AutoScheduleMeeting.DoesNotExist):
        return HttpResponseBadRequest("Invalid meeting UUID")

    # 找出對應的 email
    attendees = meeting.get_attendee_responses()
    matched_email = None

    for email, data in attendees.items():
        if data.get('tenant_id') == tenant_id:
            matched_email = email
            break

    if not matched_email:
        return HttpResponseBadRequest("Attendee not found for tenant")

    # 更新回應
    meeting.update_attendee_response(matched_email, status=response_status)
    meeting.save()
    return render(request, 'tutorial/auto_close.html', {
        'email': matched_email,
        'response': response_status
    })


def meeting_status(request, meeting_uuid):
    try:
        meeting = AutoScheduleMeeting.objects.get(uuid=meeting_uuid)
        token = get_token(request)
        # 更新會議狀態邏輯
        if meeting.status == 'waiting':
            response_summary = meeting.get_response_summary()

            # 如果有與會者拒絕，嘗試下一個候選時間
            if response_summary['declined'] > 0:
                try:
                    # declined_attendees = [email for email, response in meeting.get_attendee_responses().items() if response['status'] == 'declined']
                    # declined_list = ', '.join(declined_attendees)
                    # declined_html = ''.join(f"<li>{email}</li>" for email in declined_list)
                    # msg = (
                    #     f"<p><strong>⚠️ A participant has <span style='color:red;'>declined</span> the meeting.</strong></p>"
                    #     f"<p><strong>Declined by:</strong></p>"
                    #     f"<ul>{declined_html}</ul>"
                    #     f"<p><strong>Meeting UUID:</strong> <code>{meeting.uuid}</code></p>"
                    #     f"<p><strong>Declined meeting time:</strong><br>"
                    #     f"<span style='color:#0078D4;'>{meeting.get_candidate_time()['start']} - {meeting.get_candidate_time()['end']}</span></p>"
                    # )
                    # 更新下一段時間並初始化與會者狀態
                    meeting.try_next()
                    # inform_attendees(token, meeting, msg)
                    # 通知與會者
                    inform_attendees(token, meeting)
                except ValueError:
                    # 沒有更多候選時間，標記為失敗
                    meeting.status = 'failed'
                meeting.save()

            # 如果所有人都接受，更新狀態為 'done' 並設置選定時間
            elif response_summary['pending'] == 0 and response_summary['declined'] == 0:
                meeting.status = 'done'
                meeting.selected_time = meeting.get_candidate_time()
                meeting.save()
                # 寄出會議邀請
                attendees_emails = [email for email in meeting.get_attendee_responses().keys()]
                attendees_emails.append(meeting.host_email)
                res = create_event(
                    token,
                    meeting.title,
                    meeting.selected_time["start"],
                    meeting.selected_time["end"],
                    attendees_emails,
                    meeting.description,
                    meeting.time_zone
                )

        # 準備與會者數據
        attendees = []
        for email, response in meeting.get_attendee_responses().items():
            status_class = {
                'pending': 'warning',
                'accepted': 'success',
                'declined': 'danger',
                'tentative': 'info'
            }.get(response['status'], 'secondary')
            
            status_text = {
                'pending': 'Waiting for response',
                'accepted': 'Accepted',
                'declined': 'Declined',
                'tentative': 'Tentative'
            }.get(response['status'], '未知')
            
            attendees.append({
                'email': email,
                'status': response['status'],
                'status_class': status_class,
                'status_text': status_text,
                'response_time': response.get('response_time')
            })
        
        # 準備狀態消息
        status_messages = {
            'pending': 'Initializng...',
            'waiting': 'Waiting...',
            'done': 'Meeting scheduled successfully!',
            'failed': 'Meeting scheduling failed'
        }
        
        status_classes = {
            'pending': 'info',
            'waiting': 'warning',
            'done': 'success',
            'failed': 'danger'
        }
        
        response_data = {
            'status': meeting.status,
            'status_message': status_messages.get(meeting.status, 'unkown status'),
            'status_class': status_classes.get(meeting.status, 'secondary'),
            'attendees': attendees,
            'selected_time': meeting.selected_time if meeting.selected_time else None
        }
        
        return JsonResponse(response_data)
    except AutoScheduleMeeting.DoesNotExist:
        return JsonResponse({'error': 'Meeting not found'}, status=404)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

def get_contacts(request):
    context = initialize_context(request)
    user = context['user']
    if not user['is_authenticated']:
        return HttpResponseRedirect(reverse('signin'))

    token = get_token(request)
    query = request.GET.get('query','')  # 確保獲取 query 參數

    # 檢查 query 是否為空，若是空則返回錯誤訊息
    if not query:
        return JsonResponse({"error": "Query parameter is required"}, status=400)

    contacts = get_users(token, query=query)
    return JsonResponse(contacts, safe=False)

#我想要以下設計 我是用django
#     """
#     handles SharePoint-based reminders using GraphSharePointClient(# 一份excel 實例一個).
#     Planned Features:
#     1. Automatically trigger self.scan_routine(), which serves as a cron-like background scanner.
#     2. Start a polling task using SP.polling_task_pool() to track message replies.
    
#     UI Elements to Implement:
#     - A time picker for users to set the scan routine schedule.
#     - A time picker for users to configure the polling interval.
#     - A "Start" button to begin both the routine scan and polling process.

#     Once activated:
#     - The app should render a web page that displays all tracking items from self.model (TaskNotification).
#     - The page should support dynamic refresh (e.g., via AJAX or WebSockets) to reflect updates in real-time.
#     """

from asgiref.sync import sync_to_async
from tutorial.models import TaskNotification, SharePointClientConfig
from tutorial.graph_helper import GraphSharePointClient
import threading
import time

# Global variables to manage threads
scan_thread = None
polling_thread = None
stop_threads = False

def sharepoint_reminder_dashboard(request):
    """
    Renders the SharePoint Reminder Dashboard.
    """
    context = {
        "config": SharePointClientConfig.objects.first()  # Load the first config if it exists
    }
    return render(request, "tutorial/sharepoint_reminder_dashboard.html", context)

def start_tasks(request):
    """
    Starts the scan routine and polling tasks.
    """
    global scan_thread, polling_thread, stop_threads

    # Get the SharePointClientConfig
    config = SharePointClientConfig.objects.first()
    if not config:
        return JsonResponse({"error": "No configuration found"}, status=400)

    # Initialize GraphSharePointClient
    token = request.session.get("access_token")  # Assuming token is stored in the session
    SP = GraphSharePointClient(token)

    # Stop any existing threads
    stop_threads = True
    if scan_thread and scan_thread.is_alive():
        scan_thread.join()
    if polling_thread and polling_thread.is_alive():
        polling_thread.join()
    stop_threads = False

    # Start the scan routine thread
    def scan_routine():
        while not stop_threads:
            sync_to_async(SP.scan_routine)()
            time.sleep(config.routine_interval)

    scan_thread = threading.Thread(target=scan_routine)
    scan_thread.start()

    # Start the polling thread
    def polling_task():
        while not stop_threads:
            sync_to_async(SP.polling_task_pool)()
            time.sleep(config.polling_interval)

    polling_thread = threading.Thread(target=polling_task)
    polling_thread.start()

    return JsonResponse({"message": "Tasks started successfully"})

def stop_tasks(request):
    """
    Stops the scan routine and polling tasks.
    """
    global stop_threads
    stop_threads = True
    return JsonResponse({"message": "Tasks stopped successfully"})

def get_tracking_items(request):
    """
    Returns the list of tracking items for real-time updates.
    """
    items = TaskNotification.objects.all().values(
        "sheet_name", "row", "task", "owner_name", "reason", "replied"
    )
    return JsonResponse(list(items), safe=False)