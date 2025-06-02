from django.db import models
import json
from datetime import datetime
import uuid

# Create your models here.

class AutoScheduleMeeting(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('waiting', 'Waiting'),
        ('done', 'Done'),
        ('failed', 'Failed'),
    ]

    RESPONSE_CHOICES = [
        ('pending', 'Pending'),
        ('accepted', 'Accepted'),
        ('declined', 'Declined'),
        ('tentative', 'Tentative'),
    ]

    uuid = models.UUIDField(default=uuid.uuid4, editable=False, unique=True)
    host_email = models.EmailField(help_text="Email address of the host")
    attendees = models.JSONField(help_text="List of attendee email addresses")
    attendee_responses = models.JSONField(
        default=dict,
        help_text="Dictionary of attendee responses: {email: {'status': status, 'response_time': timestamp, 'tenant_id': tenant_id, 'chat_id': chat_id}}"
    )
    candidate_times = models.JSONField(help_text="List of candidate time slots")
    status = models.CharField(
        max_length=10,
        choices=STATUS_CHOICES,
        default='pending',
        help_text="Current status of the scheduling process"
    )
    current_try = models.IntegerField(default=0, help_text="Current attempt number")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    title = models.CharField(max_length=200, blank=True)
    description = models.TextField(blank=True)
    duration = models.IntegerField(help_text="Duration in minutes")
    start_time = models.DateTimeField()
    end_time = models.DateTimeField()
    selected_time = models.JSONField(null=True, blank=True)
    time_zone = models.CharField(max_length=50, default='UTC', help_text="Time zone of the meeting")
    def __str__(self):
        return f"Auto Schedule Meeting {self.id} - {self.status}"

    def set_attendees(self, attendees_list, tenant_ids=None, chat_ids=None):
        """
        設置與會者列表和他們的 tenant ID 及 chat ID
        :param attendees_list: 與會者郵箱列表
        :param tenant_ids: 與會者對應的 tenant ID 列表，如果為 None 則所有與會者使用相同的 tenant ID
        :param chat_ids: 與會者對應的 chat ID 列表，如果為 None 則所有與會者使用相同的 chat ID
        """
        self.attendees = json.dumps(attendees_list)
        # 初始化每個與會者的回應狀態
        responses = {}
        for i, email in enumerate(attendees_list):
            tenant_id = tenant_ids[i] if tenant_ids and i < len(tenant_ids) else None
            chat_id = chat_ids[i] if chat_ids and i < len(chat_ids) else None
            responses[email] = {
                'status': 'pending',
                'response_time': None,
                'tenant_id': tenant_id,
                'chat_id': chat_id
            }
        self.attendee_responses = json.dumps(responses)

    def get_attendees(self):
        return json.loads(self.attendees)

    def get_attendee_responses(self):
        return json.loads(self.attendee_responses)

    def update_attendee_response(self, email, status, tenant_id=None, chat_id=None):
        """
        更新與會者的回應狀態
        :param email: 與會者郵箱
        :param status: 回應狀態
        :param tenant_id: 可選的 tenant ID
        :param chat_id: 可選的 chat ID
        """
        responses = self.get_attendee_responses()
        if email in responses:
            responses[email].update({
                'status': status,
                'response_time': datetime.now().isoformat()
            })
            if tenant_id is not None:
                responses[email]['tenant_id'] = tenant_id
            if chat_id is not None:
                responses[email]['chat_id'] = chat_id
            self.attendee_responses = json.dumps(responses)
            self.save()

    def get_attendee_status(self, email):
        responses = self.get_attendee_responses()
        return responses.get(email, {}).get('status', 'pending')

    def get_attendee_tenant_id(self, email):
        responses = self.get_attendee_responses()
        return responses.get(email, {}).get('tenant_id')

    def get_attendee_chat_id(self, email):
        responses = self.get_attendee_responses()
        return responses.get(email, {}).get('chat_id')

    def get_attendees_by_tenant(self, tenant_id):
        """
        獲取特定 tenant 的所有與會者
        :param tenant_id: tenant ID
        :return: 該 tenant 的與會者郵箱列表
        """
        responses = self.get_attendee_responses()
        return [email for email, data in responses.items() 
                if data.get('tenant_id') == tenant_id]

    def set_candidate_times(self, times_list):
        self.candidate_times = json.dumps(times_list)

    def get_candidate_times(self):
        return json.loads(self.candidate_times)

    def get_response_summary(self):
        responses = self.get_attendee_responses()
        summary = {
            'pending': 0,
            'accepted': 0,
            'declined': 0,
            'tentative': 0
        }
        for response in responses.values():
            summary[response['status']] += 1
        return summary

    def get_tenant_summary(self):
        """
        獲取每個 tenant 的回應統計
        :return: {tenant_id: {'pending': 0, 'accepted': 0, 'declined': 0, 'tentative': 0}}
        """
        responses = self.get_attendee_responses()
        tenant_summary = {}
        
        for response in responses.values():
            tenant_id = response.get('tenant_id')
            if tenant_id not in tenant_summary:
                tenant_summary[tenant_id] = {
                    'pending': 0,
                    'accepted': 0,
                    'declined': 0,
                    'tentative': 0
                }
            tenant_summary[tenant_id][response['status']] += 1
            
        return tenant_summary
    def try_next(self):
        """
        更新當前嘗試次數
        :return: None
        """
        if self.current_try >= len(self.candidate_times):
            raise ValueError("No more candidate times available.")
        self.current_try += 1
        responses = self.get_attendee_responses()
        for email in responses:
            responses[email]['status'] = 'pending'
            responses[email]['response_time'] = None
        self.attendee_responses = json.dumps(responses)
        self.save()
    def get_candidate_time(self):
        """
        獲取當前嘗試的候選時間
        :return: 包含 'start' 和 'end' 的字典，或 None 如果沒有更多候選時間
        """
        return self.get_candidate_times()[self.current_try] if self.current_try < len(self.candidate_times) else None
       

    class Meta:
        ordering = ['-created_at']


class TaskNotification(models.Model):
    uuid = models.UUIDField(default=uuid.uuid4, editable=False, unique=True)  # 新增 UUID 欄位
    sheet_name = models.CharField(max_length=255)
    row = models.IntegerField()
    task = models.CharField(max_length=255)
    owner_id = models.CharField(max_length=255)
    owner_email = models.EmailField()
    owner_name = models.CharField(max_length=255)
    teams_group_id = models.CharField(max_length=255)
    teams_group_name = models.CharField(max_length=255)
    field_address = models.CharField(max_length=255)
    reason = models.CharField(max_length=255)
    msg_id = models.JSONField(default=list)
    replied = models.BooleanField(default=False)


    def __str__(self):
        return f"{self.sheet_name} - Row {self.row}: {self.task}"
