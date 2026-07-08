# -*- coding: utf-8 -*-
from __future__ import print_function
from pyrevit import forms, revit, DB
import csv
import io

doc = revit.doc


def to_text(value):
    if value is None:
        return u''
    try:
        if isinstance(value, bytes):
            return value.decode('utf-8')
    except:
        pass
    try:
        return unicode(value)
    except:
        return str(value)


def get_param_text(param):
    if not param or not param.HasValue:
        return u''

    try:
        storage_type = param.StorageType
        if storage_type == DB.StorageType.String:
            return to_text(param.AsString())
        if storage_type == DB.StorageType.Integer:
            value_string = param.AsValueString()
            return to_text(value_string) if value_string else to_text(param.AsInteger())
        if storage_type == DB.StorageType.Double:
            value_string = param.AsValueString()
            return to_text(value_string) if value_string else to_text(param.AsDouble())
        if storage_type == DB.StorageType.ElementId:
            element_id = param.AsElementId()
            if element_id and element_id.IntegerValue > 0:
                element = doc.GetElement(element_id)
                return to_text(element.Name) if element else to_text(element_id.IntegerValue)
            return u''
        return to_text(param.AsValueString())
    except:
        return u''


def get_view_type_name(view):
    try:
        return to_text(view.ViewType)
    except:
        return u''


def is_exportable_view(view):
    if not view:
        return False
    try:
        if view.IsTemplate:
            return False
    except:
        pass
    if isinstance(view, DB.ViewSheet):
        return False
    try:
        if view.ViewType in (
            DB.ViewType.ProjectBrowser,
            DB.ViewType.SystemBrowser,
            DB.ViewType.Internal
        ):
            return False
    except:
        pass
    return True


def collect_views():
    views = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
        if is_exportable_view(view):
            views.append(view)

    def sort_key(view):
        return (get_view_type_name(view), to_text(view.Name), view.Id.IntegerValue)

    return sorted(views, key=sort_key)


class ParamOption(forms.TemplateListItem):
    @property
    def name(self):
        return self.item


def collect_param_names(views):
    names = set()
    for view in views:
        for param in view.Parameters:
            try:
                if param.Definition and param.Definition.Name:
                    names.add(param.Definition.Name)
            except:
                pass
    return sorted(list(names))


def build_row(view, selected_params):
    row = [
        to_text(view.Id.IntegerValue),
        to_text(view.UniqueId),
        get_view_type_name(view),
        to_text(view.Name)
    ]

    for param_name in selected_params:
        param = view.LookupParameter(param_name)
        row.append(get_param_text(param))

    return row


def write_csv(path, headers, rows):
    with io.open(path, 'wb') as csv_file:
        csv_file.write(b'\xef\xbb\xbf')
        writer = csv.writer(csv_file)
        writer.writerow([to_text(header).encode('utf-8') for header in headers])
        for row in rows:
            writer.writerow([to_text(value).encode('utf-8') for value in row])


def main():
    views = collect_views()
    if not views:
        forms.alert('No project views were found.', exitscript=True)
        return

    param_names = collect_param_names(views)
    default_checked = [
        'View Name',
        'Title on Sheet',
        'View Template',
        'Detail Number',
        'Sheet Number',
        'Sheet Name',
        'Scale',
        'Discipline',
        'Sub-Discipline'
    ]

    options = [ParamOption(name, checked=(name in default_checked)) for name in param_names]
    selected_params = forms.SelectFromList.show(
        options,
        multiselect=True,
        title='Select view parameters to export',
        button_name='Export'
    )

    if not selected_params:
        forms.alert('No parameters selected.', exitscript=True)
        return

    system_headers = ['ElementId', 'UniqueId', 'View Type', 'Current View Name']
    headers = system_headers + list(selected_params)

    csv_path = forms.save_file(file_ext='csv', title='Export views CSV')
    if not csv_path:
        forms.alert('Export cancelled.', exitscript=True)
        return

    try:
        rows = [build_row(view, selected_params) for view in views]
        write_csv(csv_path, headers, rows)
        forms.alert(
            'Exported {} views to:\n{}'.format(len(rows), csv_path),
            title='Export complete',
            warn_icon=False
        )
    except Exception as err:
        forms.alert('Export failed:\n{}'.format(err), exitscript=True)


if __name__ == '__main__':
    main()
