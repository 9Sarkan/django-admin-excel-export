import datetime

from django.core.exceptions import PermissionDenied
from django.urls import path

from apps.export.helpers import export_excel
from apps.export.response import response_xls


class ExportAdminMixin(object):
    def export_excel(self, request, *args, **kwargs):
        if not self.has_change_permission(request):
            raise PermissionDenied
        queryset = self.get_changelist_instance(request).get_queryset(request)

        related_fields = getattr(self, "excel_related_fields", [])
        excel_fields_exclude = getattr(self, "excel_fields_exclude", [])

        wb = export_excel(queryset, excel_fields_exclude, related_fields)
        now = datetime.datetime.now()
        file_name = f'{self.model._meta.model_name.upper()}_{now.strftime("%d%b%Y_%H-%M-%S")}'
        return response_xls(file_name, wb)

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('export/excel/',
                 self.export_excel,
                 name=f'{self.model._meta.model_name}-export-excel')
        ]
        return custom_urls + urls


def export_as_xlsx(ModelAdmin, request, queryset):
    """
    used for action
    """
    related_fields = getattr(ModelAdmin, "excel_related_fields", [])
    excel_fields_exclude = getattr(ModelAdmin, "excel_fields_exclude", [])
    wb = export_excel(queryset, excel_fields_exclude, related_fields)
    now = datetime.datetime.now()
    file_name = f'{queryset[0]._meta.model._meta.model_name.upper()}_{now.strftime("%d%b%Y_%H-%M-%S")}'
    return response_xls(file_name, wb)
