# -*- coding: utf-8 -*-
from __future__ import print_function
from pyrevit import forms, revit, DB
import csv

doc = revit.doc

SYSTEM_COLUMNS = ['ElementId', 'UniqueId', 'View Type', 'Current View Name']
VIEW_NAME_COLUMNS = ['View Name']


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


def normalize(value):
    text = to_text(value)
    return text.replace(u'\u3000', u' ').strip()


def read_csv_any(path):
    encodings = ('utf-8-sig', 'utf-8', 'cp950')
    last_error = None

    for encoding in encodings:
        try:
            with open(path, 'rb') as csv_file:
                data = csv_file.read()
                text = data.decode(encoding, 'ignore')
                reader = csv.reader(text.splitlines())
                headers = [normalize(header) for header in next(reader)]

                for index, header in enumerate(headers):
                    if not header:
                        headers[index] = 'Column{}'.format(index + 1)

                rows = []
                for row in reader:
                    item = {}
                    for index, header in enumerate(headers):
                        item[header] = normalize(row[index]) if index < len(row) else u''
                    rows.append(item)

                return headers, rows, encoding
        except StopIteration:
            return [], [], encoding
        except Exception as error:
            last_error = error

    raise Exception('Could not read CSV as utf-8-sig, utf-8, or cp950: {}'.format(last_error))


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


def get_view_by_unique_id(unique_id):
    if not unique_id:
        return None
    try:
        view = doc.GetElement(unique_id)
        return view if isinstance(view, DB.View) and is_exportable_view(view) else None
    except:
        return None


def get_view_by_element_id(element_id_text):
    if not element_id_text:
        return None
    try:
        element_id = DB.ElementId(int(element_id_text))
        view = doc.GetElement(element_id)
        return view if isinstance(view, DB.View) and is_exportable_view(view) else None
    except:
        return None


def get_row_view(row):
    view = get_view_by_unique_id(row.get('UniqueId'))
    if view:
        return view
    return get_view_by_element_id(row.get('ElementId'))


def get_param_text(param):
    if not param or not param.HasValue:
        return u''
    try:
        storage_type = param.StorageType
        if storage_type == DB.StorageType.String:
            return normalize(param.AsString())
        if storage_type == DB.StorageType.Integer:
            value_string = param.AsValueString()
            return normalize(value_string) if value_string else normalize(param.AsInteger())
        if storage_type == DB.StorageType.Double:
            value_string = param.AsValueString()
            return normalize(value_string) if value_string else normalize(param.AsDouble())
        if storage_type == DB.StorageType.ElementId:
            element_id = param.AsElementId()
            if element_id and element_id.IntegerValue > 0:
                element = doc.GetElement(element_id)
                return normalize(element.Name) if element else normalize(element_id.IntegerValue)
            return u''
        return normalize(param.AsValueString())
    except:
        return u''


def set_param_value(param, value):
    if param.IsReadOnly:
        raise Exception('Parameter is read-only.')

    storage_type = param.StorageType
    if storage_type == DB.StorageType.String:
        param.Set(value)
    elif storage_type == DB.StorageType.Integer:
        if value:
            param.Set(int(value))
    elif storage_type == DB.StorageType.Double:
        if value:
            try:
                if not param.SetValueString(value):
                    param.Set(float(value))
            except:
                param.Set(float(value))
    elif storage_type == DB.StorageType.ElementId:
        if value:
            param.Set(DB.ElementId(int(value)))
    else:
        param.Set(value)


def update_view_name(view, value):
    value = normalize(value)
    if not value or normalize(view.Name) == value:
        return False
    view.Name = value
    return True


def update_view_param(view, column_name, value):
    value = normalize(value)

    if column_name in VIEW_NAME_COLUMNS:
        return update_view_name(view, value)

    param = view.LookupParameter(column_name)
    if not param:
        raise Exception('Parameter was not found on this view.')

    old_value = get_param_text(param)
    if old_value == value:
        return False

    set_param_value(param, value)
    return True


class ColumnOption(forms.TemplateListItem):
    @property
    def name(self):
        return self.item


def main():
    csv_path = forms.pick_file(file_ext='csv', title='Select edited views CSV')
    if not csv_path:
        forms.alert('Import cancelled.')
        return

    try:
        headers, rows, used_encoding = read_csv_any(csv_path)
    except Exception as error:
        forms.alert('CSV read failed:\n{}'.format(error), exitscript=True)
        return

    if not headers or not rows:
        forms.alert('CSV is empty.', exitscript=True)
        return

    if 'UniqueId' not in headers and 'ElementId' not in headers:
        forms.alert('CSV must contain UniqueId or ElementId.', exitscript=True)
        return

    candidates = [header for header in headers if header not in SYSTEM_COLUMNS]
    if not candidates:
        forms.alert('No editable columns were found.', exitscript=True)
        return

    options = [ColumnOption(header, checked=(header not in ['Current View Name'])) for header in candidates]
    target_columns = forms.SelectFromList.show(
        options,
        multiselect=True,
        title='Select view columns to import',
        button_name='Import'
    )

    if not target_columns:
        forms.alert('No columns selected.', exitscript=True)
        return

    updated = 0
    matched = 0
    not_found = 0
    skipped = 0
    failures = []

    transaction = DB.Transaction(doc, 'Import views CSV')
    transaction.Start()
    try:
        for row in rows:
            view = get_row_view(row)
            if not view:
                not_found += 1
                continue

            matched += 1
            row_updated = False

            for column_name in target_columns:
                if column_name not in row:
                    continue
                try:
                    if update_view_param(view, column_name, row.get(column_name)):
                        row_updated = True
                except Exception as error:
                    failures.append((to_text(view.Id.IntegerValue), column_name, to_text(error)))

            if row_updated:
                updated += 1
            else:
                skipped += 1

        transaction.Commit()
    except Exception as error:
        transaction.RollBack()
        forms.alert('Import failed and changes were rolled back:\n{}'.format(error), exitscript=True)
        return

    message = [
        'Import complete. CSV encoding: {}'.format(used_encoding),
        '- Matched views: {}'.format(matched),
        '- Updated views: {}'.format(updated),
        '- Unchanged views: {}'.format(skipped),
        '- Views not found: {}'.format(not_found)
    ]

    if failures:
        message.append('- Failed field updates: {}'.format(len(failures)))
        message.append('  First failure: {}'.format(failures[0]))

    forms.alert('\n'.join(message), title='Import views CSV', warn_icon=False)


if __name__ == '__main__':
    main()
