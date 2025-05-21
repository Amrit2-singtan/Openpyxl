import pandas as pd
from django.conf import settings
from datetime import datetime
import os
from rest_framework.views import APIView
from irhrs.export.constants import NORMAL_USER, ADMIN,SUPERVISOR, QUEUED, FAILED, COMPLETED, PROCESSING
from irhrs.core.mixins.serializers import create_dummy_serializer, DummySerializer

ExportNameSerializer = create_dummy_serializer({
    "export_name": serializers.CharField(max_length=150, allow_blank=True, allow_null=True,
                                         required=False)
})

class ExportMixin(APIView):
    export_fields = None
    export_type = None

    export_description = None
    heading_map = None
    footer_data = None
    frontend_redirect_url = None
    notification_permissions = []

    def get_frontend_redirect_url(self):
        """Notification Redirect URL"""
        return self.frontend_redirect_url

    def get_notification_permissions(self):
        """Notification Permissions"""
        return self.notification_permissions

    def get_export_description(self):
        """Description lines to be added before data"""
        return self.export_description

    def get_export_type(self):
        assert self.export_type is not None, f"{self.__class__} must define  `export_type`"
        return self.export_type

    def get_export_fields(self):
        assert self.export_fields is not None, f"{self.__class__} must define  `export_fields`"
        return self.export_fields
    
    def get_exported_as(self):
        mode = self.request.query_params.get('as')
        return {
            "hr": ADMIN,
            "supervisor": SUPERVISOR
        }.get(mode, NORMAL_USER)
    
    def get_export_name(self):
        serializer = ExportNameSerializer(data=self.request.data)
        serializer.is_valid(raise_exception=True)
        return serializer.data.get('export_name') or self.get_export_type()
    
    def get_extra_export_data(self):
        return dict(
            organization=self.get_organization() if hasattr(self, 'get_organization') else None,
            redirect_url=self.get_frontend_redirect_url(),
            exported_as=self.get_exported_as(),
            notification_permissions=self.get_notification_permissions()
        )
    
    def _export_get(self):
        """
        Get all task in xlsx
        returns previously exported file if found
        """
        if hasattr(self, 'get_organization'):
            organization = self.get_organization()
        else:
            organization = None
        latest_export = get_latest_export(
            export_type=self.get_export_type(),
            user=self.request.user,
            exported_as=self.get_exported_as(),
            organization=organization
        )

        if latest_export:
            return Response({
                'url': get_complete_url(latest_export.export_file.url),
                'created_on': latest_export.modified_at
            })
        else:
            return Response({
                'message': 'Previous Export file couldn\'t be found.',
                'url': ''
            }, status=status.HTTP_200_OK
            )
        
    @classmethod
    def get_exported_file_content(cls, queryset, title, columns, extra_content,
                                  description=None, heading_map=None, footer_data=None
                                  ):
        """Return contents of exported file of type ContentFile"""
        raise NotImplementedError

    @classmethod
    def save_file_content(cls, export_instance, file_content):
        """Save file_content and set export_instance.export_file and export_instance.status to Successful"""
        raise NotImplementedError

    def get_footer_data(self):
        return self.footer_data
    
    @classmethod
    def send_success_notification(cls, obj, url, exported_as, permissions):
        if exported_as == ADMIN:
            notify_organization(
                permissions=permissions,
                text=f"{pretty_name(obj.name).replace('report', ' ')} report has been generated.",
                action=obj,
                organization=obj.organization,
                actor=get_system_admin(),
                url=url+f"/?export={obj.id}" if url else ''
            )
        else:
            name = obj.name or ""
            word_list=re.findall('[A-Z][^A-Z]*', name)
            seprated_word_name=" ".join(word_list)
            add_notification(
                text=f"{pretty_name(seprated_word_name).replace('report', ' ')} report has been generated.",
                action=obj,
                recipient=obj.user,
                actor=get_system_admin(),
                url=url+f"/?export={obj.id}" if url else ''
            )

    @classmethod
    def send_failed_notification(cls, obj, url, exported_as, permissions):
        if exported_as == ADMIN:
            notify_organization(
                permissions=permissions,
                text=f"Failed to generate {pretty_name(obj.name).replace('report', '')} report",
                action=obj,
                organization=obj.organization,
                actor=get_system_admin(),
                url=url+f"/?export={obj.id}" if url else ''
            )
        else:
            add_notification(
                text=f"Failed to generate {pretty_name(obj.name).replace('report', '')} report",
                action=obj,
                recipient=obj.user,
                actor=get_system_admin(),
                url=url+f"/?export={obj.id}" if url else ''
            )

    @action(methods=['GET', 'POST'], detail=False, serializer_class=ExportNameSerializer)
    def export(self, *args, **kwargs):
        if self.request.method.upper() == 'GET':
            return self._export_get()
        else:
            return self._export_post()
        
    
    def get_export_filename(self):
        base_name = self.request.GET.get("filename") or self.export_filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_{timestamp}.xlsx"

    def _export_post(self, data):
        export_fields = self.get_export_fields()
        export = Export.objects.create(
            user=self.request.user,
            name=self.get_export_name(),
            exported_as=self.get_exported_as(),
            export_type=self.get_export_type(),
            organization=organization,
            status=QUEUED,
            remarks=''
        )

        try:
            if export_fields:
                data = [{field: self._get_nested_value(item, field) for field in export_fields} for item in data]

            df = pd.DataFrame(data)

            filename = self.get_export_filename()
            file_path = os.path.join(settings.MEDIA_ROOT, filename)

            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name=self.get_export_title())

            file_url = os.path.join(settings.MEDIA_URL, filename)
            return file_url
        
        except Exception as e:
            import traceback

            logger.error(e, exc_info=True)

            export.status = FAILED
            export.message = "Could not start export."
            export.traceback = str(traceback.format_exc())
            export.save()
            if getattr(settings, 'DEBUG', False):
                raise e
            return Response({
                'message': 'The export could not be completed.'
            }, status=400)

        return Response({
            'message': 'Your request is being processed in the background . Please check back later'})

    def _get_nested_value(self, data, dotted_key):
        keys = dotted_key.split(".")
        for key in keys:
            data = data.get(key, {})
        return data if data != {} else ""


# your_app/views.py

from django.views import View
from django.http import JsonResponse
from .utils.export_mixin import ExportMixin

SAMPLE_DATA = [
    {
        "timesheet_for": "2025-04-17",
        "timesheet_user": {
            "full_name": "kushal bhurtel",
            "email": "kushal@gmail.com",
            "employee_code": "EMP1033",
        },
        "punch_in": "N/A",
        "punch_out": "N/A",
        "worked_hours": "00:00:00",
        "expected_work_hours": "09:00:00",
        "overtime": "00:00:00",
        "coefficient": "Workday",
        "leave_coefficient": "No Leave",
    },
]


class AttendanceExportView(ExportMixin, View):
    export_fields = [
        "timesheet_for",
        "timesheet_user.full_name",
        "timesheet_user.email",
        "timesheet_user.employee_code",
        "punch_in",
        "punch_out",
        "worked_hours",
        "expected_work_hours",
        "overtime",
        "coefficient",
        "leave_coefficient",
    ]
    export_title = "Daily Attendance"
    export_filename = "attendance_report"

    def get(self, request, *args, **kwargs):
        if request.GET.get("export") == "excel":
            return self.export_to_excel(SAMPLE_DATA)

        return JsonResponse({
            "message": "Add '?export=excel' to the URL to download the Excel file."
        })
