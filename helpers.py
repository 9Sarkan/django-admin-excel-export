import datetime
import logging

import xlwt
from django.core.exceptions import FieldDoesNotExist
from django.db.models import ForeignKey, OneToOneField
from django.db.models.fields.files import ImageField

logger = logging.getLogger(__name__)


def export_excel(queryset,
                 excel_fields_exclude: list = [],
                 related_fields: list = []):
    wb = xlwt.Workbook(encoding='utf-8')

    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    model = queryset[0]._meta.model if queryset.exists() else None
    if model:
        ws = wb.add_sheet(model._meta.model_name)
        write_to_sheet(ws, queryset, model, excel_fields_exclude,
                       related_fields)
    else:
        ws = wb.add_sheet("empty")
    return wb


def write_to_sheet(ws,
                   queryset,
                   model,
                   excel_fields_exclude: list = [],
                   related_fields: list = []):
    fa_columns = []
    columns = []
    row_num = 0
    font_style = xlwt.XFStyle()
    validated_related_fields = []
    excel_fields_exclude = (*excel_fields_exclude, "password")

    # add main fields
    for field in model._meta.fields:
        if field.name not in excel_fields_exclude and not isinstance(
                field, ImageField):
            fa_columns.append(str(field.verbose_name))
            columns.append(field.name)

    # add related fields
    for field in related_fields:
        # get model
        if not (isinstance(field, tuple) and len(field) == 2):
            # raise invalid data in excel_related_fields
            raise Exception(
                "invalid data in excel_related_fields, it's must be a list of tuble that has 2 elements."
            )
        for item in field:
            if not isinstance(item, str):
                raise Exception(
                    "excel_related_fields inner items must be str!")

    # add main field titles
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, fa_columns[col_num], font_style)
    # add related fields
    for index, field in enumerate(related_fields):
        try:
            col_num = len(columns) + index
            ws.write(row_num, col_num, field[1], font_style)
            validated_related_fields.append(field)
        except Exception as e:
            logger.error(e)

    for i in range(queryset.count()):
        for col_num in range(0, len(columns)):
            data = get_field_data(queryset[i], model, columns[col_num])
            ws.write(i + 1, col_num, data, font_style)
        # add related fields data if exists
        for index, field in enumerate(validated_related_fields):
            col_num = len(columns) + index
            item = queryset[i]
            related_model = getattr(item, field[0], None)
            if related_model:
                related_model = related_model._meta.model
            else:
                raise Exception(f"{field[0]} is not a related model.")
            data = get_field_data(item, related_model, field, True)
            ws.write(i + 1, col_num, data, font_style)


def get_field_data(item, model, field_name, related=False):
    field = model._meta.get_field(
        field_name) if not related else model._meta.get_field(field_name[1])
    if isinstance(field, (OneToOneField, ForeignKey)):
        if related:
            item = getattr(item, field_name[0], None)
        if field.name == 'user':
            data = getattr(item, f'{field.name}')
            data = data.email if data else data.__str__()
        else:
            data = getattr(item, f'{field.name}').__str__()
    else:
        if related:
            related_object = getattr(item, field_name[0], None)
            data = field.value_from_object(related_object)
        else:
            data = getattr(item, field_name)

    if isinstance(data, datetime.date):
        date_format = '%Y/%m/%d %H:%M' if isinstance(
            data, datetime.datetime) else '%Y/%m/%d'
        data = data.strftime(date_format)
    return data
