# -*- coding: utf-8 -*-
"""Export Revit painted faces to OBJ/MTL for Rhino."""

import codecs
import csv
import os
import re
import traceback

from Autodesk.Revit.DB import (
    CategoryType,
    Color,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    Options,
    Solid,
    Transform,
    ViewDetailLevel,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms, revit, script


doc = revit.doc
uidoc = revit.uidoc

MM_PER_FOOT = 304.8
MIN_FACE_AREA = 1e-9


class ExportError(Exception):
    """Raised when painted faces cannot be exported."""


class MaterialInfo(object):
    """Minimal material data for OBJ MTL export."""

    def __init__(self, key, name, color, transparency):
        self.key = key
        self.name = name
        self.color = color
        self.transparency = transparency


class FaceMesh(object):
    """Export data for one painted Revit face."""

    def __init__(self, element_id, element_name, category_name, material_key, material_name, material_source, triangles, boundary_loops, area):
        self.element_id = element_id
        self.element_name = element_name
        self.category_name = category_name
        self.material_key = material_key
        self.material_name = material_name
        self.material_source = material_source
        self.triangles = triangles
        self.boundary_loops = boundary_loops
        self.area = area


def safe_filename(value):
    """Return a file-system safe name."""
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "_", value or "painted_faces")
    cleaned = cleaned.strip(" .")
    return cleaned or "painted_faces"


def safe_obj_name(value):
    """Return an OBJ-safe object/material token."""
    cleaned = re.sub(r"[^0-9A-Za-z_\\-]+", "_", value or "Material")
    cleaned = cleaned.strip("_")
    return cleaned or "Material"


def get_element_name(element):
    """Return a display name for an element."""
    try:
        if element.Name:
            return element.Name
    except Exception:
        pass
    return "Element_{}".format(element.Id.IntegerValue)


def get_category_name(element):
    """Return element category name."""
    try:
        if element.Category:
            return element.Category.Name
    except Exception:
        pass
    return "No Category"


def get_material_color(material):
    """Return a material color, falling back to neutral gray."""
    try:
        color = material.Color
        if color and color.IsValid:
            return color
    except Exception:
        pass
    return Color(180, 180, 180)


def get_material_info(material_id, material_cache):
    """Collect and cache material information by ElementId."""
    mat_int = material_id.IntegerValue
    if mat_int in material_cache:
        return material_cache[mat_int]

    material = doc.GetElement(material_id)
    if material is None:
        name = "Paint_Material_{}".format(mat_int)
        color = Color(180, 180, 180)
        transparency = 0
    else:
        name = material.Name
        color = get_material_color(material)
        try:
            transparency = int(material.Transparency)
        except Exception:
            transparency = 0

    key = "{}_{}".format(safe_obj_name(name), mat_int)
    info = MaterialInfo(key, name, color, transparency)
    material_cache[mat_int] = info
    return info


def is_valid_material_id(material_id):
    """Return True for usable Revit material ids."""
    return material_id and material_id != ElementId.InvalidElementId and material_id.IntegerValue > 0


def get_painted_material_id(element_id, face):
    """Return painted material id for a face, or None when it is not painted."""
    try:
        if hasattr(doc, "IsPainted") and not doc.IsPainted(element_id, face):
            return None
    except Exception:
        pass

    try:
        material_id = doc.GetPaintedMaterial(element_id, face)
        if is_valid_material_id(material_id):
            return material_id
    except Exception:
        return None
    return None


def get_face_material_id(element_id, face, is_split_region):
    """Return the export material id and source for a painted face or split-face region."""
    material_id = get_painted_material_id(element_id, face)
    if is_valid_material_id(material_id):
        return material_id, "paint"

    # Split Face regions expose their independently painted material through
    # MaterialElementId in some Revit API contexts.
    if is_split_region:
        try:
            material_id = face.MaterialElementId
            if is_valid_material_id(material_id):
                return material_id, "split_region_material"
        except Exception:
            pass

    return None, None


def transform_point(point, transform):
    """Transform a Revit XYZ point and convert feet to millimeters."""
    try:
        point = transform.OfPoint(point)
    except Exception:
        pass
    return (point.X * MM_PER_FOOT, point.Y * MM_PER_FOOT, point.Z * MM_PER_FOOT)


def points_close(point_a, point_b, tolerance):
    """Return True when two exported point tuples are nearly equal."""
    return (
        abs(point_a[0] - point_b[0]) <= tolerance and
        abs(point_a[1] - point_b[1]) <= tolerance and
        abs(point_a[2] - point_b[2]) <= tolerance
    )


def triangulate_face(face, transform):
    """Triangulate one Revit face into OBJ-ready triangle coordinate tuples."""
    mesh = face.Triangulate()
    triangles = []
    for index in range(mesh.NumTriangles):
        tri = mesh.get_Triangle(index)
        p0 = transform_point(tri.get_Vertex(0), transform)
        p1 = transform_point(tri.get_Vertex(1), transform)
        p2 = transform_point(tri.get_Vertex(2), transform)
        triangles.append((p0, p1, p2))
    return triangles


def curve_points_on_face(edge, face):
    """Return tessellated points for an edge following the given face."""
    try:
        curve = edge.AsCurveFollowingFace(face)
    except Exception:
        curve = edge.AsCurve()
    try:
        return list(curve.Tessellate())
    except Exception:
        return [curve.GetEndPoint(0), curve.GetEndPoint(1)]


def extract_face_boundary_loops(face, transform):
    """Extract painted face edge loops as closed 3D polylines in millimeters."""
    loops = []
    for edge_loop in face.EdgeLoops:
        points = []
        for edge in edge_loop:
            curve_points = curve_points_on_face(edge, face)
            for point in curve_points:
                exported = transform_point(point, transform)
                if not points or not points_close(points[-1], exported, 0.001):
                    points.append(exported)

        if len(points) >= 2 and points_close(points[0], points[-1], 0.001):
            points.pop()
        if len(points) >= 3:
            loops.append(points)
    return loops


def iter_export_faces(face):
    """Yield split-face regions when present; otherwise yield the original face."""
    try:
        if face.HasRegions:
            regions = list(face.GetRegions())
            if regions:
                for region in regions:
                    yield region, True
                return
    except Exception:
        pass

    yield face, False


def iter_geometry_objects(geometry_element, transform):
    """Yield solids from nested Revit geometry."""
    if geometry_element is None:
        return

    for obj in geometry_element:
        if isinstance(obj, Solid):
            if obj.Faces and obj.Faces.Size > 0:
                yield obj, transform
        elif isinstance(obj, GeometryInstance):
            try:
                instance_geometry = obj.GetInstanceGeometry()
                for item in iter_geometry_objects(instance_geometry, Transform.Identity):
                    yield item
            except Exception:
                try:
                    symbol_geometry = obj.GetSymbolGeometry()
                    nested_transform = transform.Multiply(obj.Transform)
                    for item in iter_geometry_objects(symbol_geometry, nested_transform):
                        yield item
                except Exception:
                    pass


def collect_candidate_elements(scope):
    """Collect model elements from the chosen export scope."""
    if scope == "selected":
        ids = list(uidoc.Selection.GetElementIds())
        if not ids:
            raise ExportError("No selected elements. Select elements first or export the active view.")
        return [doc.GetElement(eid) for eid in ids if doc.GetElement(eid) is not None]

    if scope == "active_view":
        view = doc.ActiveView
        elements = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    else:
        elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

    output = []
    for element in elements:
        try:
            if element.Category is not None and element.Category.CategoryType == CategoryType.Model:
                output.append(element)
        except Exception:
            pass
    return output


def choose_scope():
    """Ask whether to export selected elements or active-view visible elements."""
    selected_count = uidoc.Selection.GetElementIds().Count
    options = []
    if selected_count > 0:
        options.append("Selected elements ({})".format(selected_count))
    options.append("Visible elements in active view")
    options.append("Entire model")

    picked = forms.SelectFromList.show(
        options,
        title="Export Painted Faces - Scope",
        button_name="Export"
    )
    if not picked:
        script.exit()
    if picked.startswith("Selected"):
        return "selected"
    if picked.startswith("Entire"):
        return "entire_model"
    return "active_view"


def choose_export_format():
    """Ask which file format to export."""
    options = [
        "DXF boundaries only (recommended for Rhino editing)",
        "OBJ mesh only (triangulated reference)",
        "Both DXF boundaries and OBJ mesh"
    ]
    picked = forms.SelectFromList.show(
        options,
        title="Export Painted Faces - Format",
        button_name="Continue"
    )
    if not picked:
        script.exit()
    if picked.startswith("DXF"):
        return "dxf"
    if picked.startswith("OBJ"):
        return "obj"
    return "both"


def choose_output_paths():
    """Ask for output folder and base filename."""
    folder = forms.pick_folder(title="Choose export folder")
    if not folder:
        script.exit()

    default_name = safe_filename("{}_painted_faces".format(doc.Title))
    name = forms.ask_for_string(
        title="Export File Name",
        prompt="File name without extension",
        default=default_name
    )
    if not name:
        script.exit()

    base = safe_filename(name)
    obj_path = os.path.join(folder, base + ".obj")
    mtl_path = os.path.join(folder, base + ".mtl")
    dxf_path = os.path.join(folder, base + ".dxf")
    csv_path = os.path.join(folder, base + "_summary.csv")
    return obj_path, mtl_path, dxf_path, csv_path, base + ".mtl"


def make_geometry_options():
    """Create geometry options for face/reference access."""
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = False
    options.DetailLevel = ViewDetailLevel.Fine
    try:
        options.View = doc.ActiveView
    except Exception:
        pass
    return options


def collect_painted_faces(elements, include_mesh):
    """Collect all painted faces and convert them to export data."""
    options = make_geometry_options()
    material_cache = {}
    face_meshes = []
    checked_faces = 0
    skipped_faces = 0
    skipped_elements = 0

    for element in elements:
        try:
            geometry = element.get_Geometry(options)
        except Exception:
            skipped_elements += 1
            continue

        try:
            solids = list(iter_geometry_objects(geometry, Transform.Identity))
        except Exception:
            skipped_elements += 1
            continue

        for solid, transform in solids:
            for host_face in solid.Faces:
                checked_faces += 1
                for face, is_split_region in iter_export_faces(host_face):
                    try:
                        if face.Area <= MIN_FACE_AREA:
                            skipped_faces += 1
                            continue

                        material_id, material_source = get_face_material_id(element.Id, face, is_split_region)
                        if not is_valid_material_id(material_id):
                            continue

                        material_info = get_material_info(material_id, material_cache)
                        boundary_loops = extract_face_boundary_loops(face, transform)
                        if not boundary_loops:
                            skipped_faces += 1
                            continue
                        triangles = triangulate_face(face, transform) if include_mesh else []
                        if include_mesh and not triangles:
                            skipped_faces += 1
                            continue

                        face_meshes.append(FaceMesh(
                            element.Id.IntegerValue,
                            get_element_name(element),
                            get_category_name(element),
                            material_info.key,
                            material_info.name,
                            material_source,
                            triangles,
                            boundary_loops,
                            face.Area
                        ))
                    except Exception:
                        skipped_faces += 1

    return face_meshes, material_cache, checked_faces, skipped_faces, skipped_elements


def write_dxf_pair(dxf, code, value):
    """Write a DXF group code pair."""
    dxf.write("{}\n{}\n".format(code, value))


def write_dxf_header(dxf):
    """Write a minimal DXF R12 header."""
    write_dxf_pair(dxf, 0, "SECTION")
    write_dxf_pair(dxf, 2, "HEADER")
    write_dxf_pair(dxf, 9, "$ACADVER")
    write_dxf_pair(dxf, 1, "AC1009")
    write_dxf_pair(dxf, 9, "$INSUNITS")
    write_dxf_pair(dxf, 70, 4)
    write_dxf_pair(dxf, 0, "ENDSEC")


def write_dxf_tables(dxf, material_infos):
    """Write layer table, one layer per painted material."""
    write_dxf_pair(dxf, 0, "SECTION")
    write_dxf_pair(dxf, 2, "TABLES")
    write_dxf_pair(dxf, 0, "TABLE")
    write_dxf_pair(dxf, 2, "LAYER")
    write_dxf_pair(dxf, 70, len(material_infos))
    for key in sorted(material_infos.keys()):
        info = material_infos[key]
        write_dxf_pair(dxf, 0, "LAYER")
        write_dxf_pair(dxf, 2, info.key)
        write_dxf_pair(dxf, 70, 0)
        write_dxf_pair(dxf, 62, 7)
        write_dxf_pair(dxf, 6, "CONTINUOUS")
    write_dxf_pair(dxf, 0, "ENDTAB")
    write_dxf_pair(dxf, 0, "ENDSEC")


def write_dxf_polyline(dxf, layer, points):
    """Write a closed 3D polyline loop."""
    write_dxf_pair(dxf, 0, "POLYLINE")
    write_dxf_pair(dxf, 8, layer)
    write_dxf_pair(dxf, 66, 1)
    write_dxf_pair(dxf, 70, 9)

    for point in points:
        write_dxf_pair(dxf, 0, "VERTEX")
        write_dxf_pair(dxf, 8, layer)
        write_dxf_pair(dxf, 10, "{:.6f}".format(point[0]))
        write_dxf_pair(dxf, 20, "{:.6f}".format(point[1]))
        write_dxf_pair(dxf, 30, "{:.6f}".format(point[2]))
        write_dxf_pair(dxf, 70, 32)

    write_dxf_pair(dxf, 0, "SEQEND")
    write_dxf_pair(dxf, 8, layer)


def write_dxf(dxf_path, face_meshes, material_infos):
    """Write painted face boundary loops as DXF 3D polylines."""
    with codecs.open(dxf_path, "w", "ascii", "ignore") as dxf:
        write_dxf_header(dxf)
        write_dxf_tables(dxf, material_infos)
        write_dxf_pair(dxf, 0, "SECTION")
        write_dxf_pair(dxf, 2, "ENTITIES")

        for face_mesh in face_meshes:
            for loop in face_mesh.boundary_loops:
                write_dxf_polyline(dxf, face_mesh.material_key, loop)

        write_dxf_pair(dxf, 0, "ENDSEC")
        write_dxf_pair(dxf, 0, "EOF")


def write_mtl(mtl_path, material_infos):
    """Write an OBJ material library."""
    with codecs.open(mtl_path, "w", "utf-8") as mtl:
        mtl.write("# Revit painted face materials\n")
        for key in sorted(material_infos.keys()):
            info = material_infos[key]
            r = info.color.Red / 255.0
            g = info.color.Green / 255.0
            b = info.color.Blue / 255.0
            alpha = max(0.0, min(1.0, 1.0 - (float(info.transparency) / 100.0)))
            mtl.write("\n")
            mtl.write("newmtl {}\n".format(info.key))
            mtl.write("Ka {:.6f} {:.6f} {:.6f}\n".format(r, g, b))
            mtl.write("Kd {:.6f} {:.6f} {:.6f}\n".format(r, g, b))
            mtl.write("Ks 0.000000 0.000000 0.000000\n")
            mtl.write("d {:.6f}\n".format(alpha))
            mtl.write("illum 2\n")


def write_obj(obj_path, mtl_filename, face_meshes):
    """Write painted face meshes to OBJ."""
    vertex_index = 1
    with codecs.open(obj_path, "w", "utf-8") as obj:
        obj.write("# Exported from Revit painted faces\n")
        obj.write("# Units: millimeters\n")
        obj.write("mtllib {}\n".format(mtl_filename))

        for face_index, face_mesh in enumerate(face_meshes, 1):
            group_name = "E{}_F{}_{}".format(
                face_mesh.element_id,
                face_index,
                safe_obj_name(face_mesh.material_name)
            )
            obj.write("\n")
            obj.write("g {}\n".format(group_name))
            obj.write("usemtl {}\n".format(face_mesh.material_key))

            for tri in face_mesh.triangles:
                for point in tri:
                    obj.write("v {:.6f} {:.6f} {:.6f}\n".format(point[0], point[1], point[2]))
                obj.write("f {} {} {}\n".format(vertex_index, vertex_index + 1, vertex_index + 2))
                vertex_index += 3


def write_summary(csv_path, face_meshes):
    """Write a CSV summary for checking exported faces."""
    with open(csv_path, "wb") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow([
            "element_id",
            "element_name",
            "category",
            "material",
            "material_layer",
            "material_source",
            "boundary_loop_count",
            "triangle_count",
            "area_m2"
        ])
        for face_mesh in face_meshes:
            writer.writerow([
                face_mesh.element_id,
                face_mesh.element_name.encode("utf-8"),
                face_mesh.category_name.encode("utf-8"),
                face_mesh.material_name.encode("utf-8"),
                face_mesh.material_key,
                face_mesh.material_source,
                len(face_mesh.boundary_loops),
                len(face_mesh.triangles),
                "{:.6f}".format(face_mesh.area * 0.09290304)
            ])


def main():
    scope = choose_scope()
    export_format = choose_export_format()
    elements = collect_candidate_elements(scope)
    if not elements:
        raise ExportError("No candidate elements found.")

    obj_path, mtl_path, dxf_path, csv_path, mtl_filename = choose_output_paths()

    include_mesh = export_format in ("obj", "both")
    include_dxf = export_format in ("dxf", "both")
    face_meshes, material_cache, checked_faces, skipped_faces, skipped_elements = collect_painted_faces(elements, include_mesh)
    if not face_meshes:
        raise ExportError(
            "No painted faces found.\n\n"
            "This exporter only exports faces painted with Revit's Paint tool. "
            "Object type materials or layer materials are not included."
        )

    material_infos = {}
    for mat_id in material_cache:
        info = material_cache[mat_id]
        material_infos[info.key] = info

    output_lines = []
    if include_mesh:
        write_mtl(mtl_path, material_infos)
        write_obj(obj_path, mtl_filename, face_meshes)
        output_lines.append("OBJ: {}".format(obj_path))
        output_lines.append("MTL: {}".format(mtl_path))
    if include_dxf:
        write_dxf(dxf_path, face_meshes, material_infos)
        output_lines.append("DXF: {}".format(dxf_path))
    write_summary(csv_path, face_meshes)
    output_lines.append("Summary: {}".format(csv_path))

    triangle_count = sum([len(face_mesh.triangles) for face_mesh in face_meshes])
    loop_count = sum([len(face_mesh.boundary_loops) for face_mesh in face_meshes])
    message = "\n".join([
        "Export Painted Faces complete",
        "",
    ] + output_lines + [
        "",
        "Elements scanned: {}".format(len(elements)),
        "Faces checked: {}".format(checked_faces),
        "Painted faces exported: {}".format(len(face_meshes)),
        "Boundary loops: {}".format(loop_count),
        "Triangles: {}".format(triangle_count if include_mesh else "not exported"),
        "Materials: {}".format(len(material_infos)),
        "Skipped faces: {}".format(skipped_faces),
        "Skipped elements: {}".format(skipped_elements),
        "",
        "Rhino import note: coordinates are exported in millimeters.",
        "Use DXF for editable boundaries. OBJ is triangulated mesh for visual/material reference."
    ])
    forms.alert(message, title="Export Painted Faces")


try:
    main()
except OperationCanceledException:
    script.exit()
except ExportError as err:
    forms.alert(str(err), title="Export Painted Faces", exitscript=True)
except Exception as err:
    forms.alert(
        "{}\n\n{}".format(err, traceback.format_exc()),
        title="Export Painted Faces",
        exitscript=True
    )
