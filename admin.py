from django.contrib import admin
from . import mixins

mixins.export_as_xlsx.short_description = "Export as Excel file"


class UserAdmin(mixins.ExportAdminMixin, admin.ModelAdmin):
    excel_related_fields = (
        ("profile", "name"),
        ("subscription", "plan"),
        ("subscription", "status"),
    )
    excel_fields_exclude = ("id", "is_superuser", "is_staff", "is_active",
                            "pa_token")
    actions = (mixins.export_as_xlsx, )
