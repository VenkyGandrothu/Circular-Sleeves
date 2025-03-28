from traceback import print_tb

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, StorageType, ElementId, Transaction,
    FamilyInstance, LocationPoint, Structure, UV
)
from Autodesk.Revit.DB import Transaction as DBTransaction
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommandLinkId
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import UIDocument

import math

doc = revit.doc

element_collector = DB.FilteredElementCollector(doc).OfCategory(
    DB.BuiltInCategory.OST_MechanicalEquipment
).WhereElementIsNotElementType()
wall_collector = DB.FilteredElementCollector(doc).OfCategory(
    DB.BuiltInCategory.OST_Walls
).WhereElementIsNotElementType()
beam_collector = DB.FilteredElementCollector(doc).OfCategory(
    DB.BuiltInCategory.OST_StructuralFraming
).WhereElementIsNotElementType()

def do_bounding_boxes_intersect(bbox1, bbox2):
    return (bbox1.Min.X <= bbox2.Max.X and bbox1.Max.X >= bbox2.Min.X and
            bbox1.Min.Y <= bbox2.Max.Y and bbox1.Max.Y >= bbox2.Min.Y and
            bbox1.Min.Z <= bbox2.Max.Z and bbox1.Max.Z >= bbox2.Min.Z)

def get_opposite_face_of_equipment(equip_element):
    # Get the bounding box of the equipment
    equip_bbox = equip_element.get_BoundingBox(None)
    if equip_bbox is None:
        return None

    # Determine the "far end" of the mechanical equipment along the Z-axis (vertical axis)
    min_point = equip_bbox.Min
    max_point = equip_bbox.Max
    far_end_point = max_point if max_point.Z > min_point.Z else min_point

    return far_end_point

def find_intersecting_face_based_on_far_end(geo_element, far_end_point, equip_bbox, base_tolerance=0.2):
    diameter = max(equip_bbox.Max.X - equip_bbox.Min.X, equip_bbox.Max.Y - equip_bbox.Min.Y)
    threshold = 0.3937  # 120 mm in feet
    tolerance = base_tolerance * (diameter / threshold) if diameter > threshold else base_tolerance

    closest_face = None
    min_distance = float('inf')
    for geo_obj in geo_element:
        if isinstance(geo_obj, DB.Solid):
            for face in geo_obj.Faces:
                proj = face.Project(far_end_point)
                if proj:
                    distance = proj.Distance
                    if distance < min_distance and distance <= tolerance:
                        min_distance = distance
                        closest_face = face
    return closest_face

def find_intersecting_face(geo_element, point, tolerance=0.2):
    closest_face = None
    min_distance = float('inf')
    for geo_obj in geo_element:
        if isinstance(geo_obj, DB.Solid):
            for face in geo_obj.Faces:
                proj = face.Project(point)
                if proj:
                    if proj.Distance < min_distance:
                        min_distance = proj.Distance
                        closest_face = face
    if closest_face is None or min_distance > tolerance:
        sample_uvs = [DB.UV(u, v) for u in [0.2, 0.4, 0.6, 0.8] for v in [0.2, 0.4, 0.6, 0.8]]
        for geo_obj in geo_element:
            if isinstance(geo_obj, DB.Solid):
                for face in geo_obj.Faces:
                    if isinstance(face, DB.PlanarFace):
                        for uv in sample_uvs:
                            sample_pt = face.Evaluate(uv)
                            normal = face.ComputeNormal(uv)
                            dist = abs((point - sample_pt).DotProduct(normal))
                            if dist < tolerance and dist < min_distance:
                                min_distance = dist
                                closest_face = face
    return closest_face

family_collector = DB.FilteredElementCollector(doc).OfClass(Family).WhereElementIsNotElementType()
family_symbols_dict = {}
for family in family_collector:
    if family.Name == "ADR-10D SLEEVE CUTOUT-":
        target_family = family
        if hasattr(target_family, 'GetFamilySymbolIds'):
            symbol_ids = target_family.GetFamilySymbolIds()
            family_symbols_dict[family.Name] = [str(sid) for sid in symbol_ids] if symbol_ids else []

def start_drag_select_mode_and_finish():
    try:
        uidoc = revit.uidoc
        doc = revit.doc
        selected_elements = uidoc.Selection.PickElementsByRectangle("Select elements by dragging a region")
        if selected_elements:
            dialog = TaskDialog("Finish Selection")
            dialog.MainInstruction = "You have selected elements. Do you want to process them?"
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Yes, Process them")
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "No, Don't Process them")
            dialog.DefaultButton = TaskDialogResult.CommandLink1
            result = dialog.Show()
            if result == TaskDialogResult.CommandLink1:
                process_selected_elements(selected_elements)
            else:
                TaskDialog.Show("Selection Cancelled", "The selection was cancelled.")
        else:
            TaskDialog.Show("No Elements Selected", "No elements were selected, operation aborted.")
    except Exception as e:
        TaskDialog.Show("Selection Error", "Error occurred during selection: {}".format(str(e)))

family_instance_data = {}

def process_selected_elements(selected_elements):
    global family_instance_data
    for element in selected_elements:
        element_id = element.Id
        family_name = "Not a FamilyInstance"
        location_point = None
        sleeve_length = "No Sleeve Length"
        sleeve_cod = "No Sleeve COD"
        sleeve_diameter = "No Sleeve Diameter"
        intersecting_elements = []

        if isinstance(element, DB.FamilyInstance):
            try:
                family_name = element.Symbol.Family.Name if element.Symbol.Family else "No Family"
                loc = element.Location
                if isinstance(loc, DB.LocationPoint):
                    location_point = loc.Point

                sleeve_length_param = element.LookupParameter("Sleeve Length")
                if sleeve_length_param:
                    sleeve_length = sleeve_length_param.AsValueString() or "No Value"

                sleeve_cod_param = element.LookupParameter("Sleeve (COD)")
                if sleeve_cod_param:
                    sleeve_cod = sleeve_cod_param.AsValueString() or "No Value"

                sleeve_diameter_param = element.Symbol.LookupParameter("Sleeve Diameter")
                if sleeve_diameter_param:
                    sleeve_diameter = sleeve_diameter_param.AsDouble() * 304.8
                    sleeve_diameter = "{:.2f} mm".format(sleeve_diameter)
                else:
                    sleeve_diameter = "Sleeve Diameter Not Found"

                equip_bbox = element.get_BoundingBox(None)
                for wall in wall_collector:
                    wall_bbox = wall.get_BoundingBox(None)
                    if wall_bbox and equip_bbox and do_bounding_boxes_intersect(equip_bbox, wall_bbox):
                        intersecting_elements.append({'id': wall.Id, 'type': 'Wall'})
                for beam in beam_collector:
                    beam_bbox = beam.get_BoundingBox(None)
                    if beam_bbox and equip_bbox and do_bounding_boxes_intersect(equip_bbox, beam_bbox):
                        intersecting_elements.append({'id': beam.Id, 'type': 'Beam'})

            except Exception as e:
                pass

        family_instance_data[element_id] = {
            'family_name': family_name,
            'location': location_point,
            'sleeve_length': sleeve_length,
            'sleeve_cod': sleeve_cod,
            'sleeve_diameter': sleeve_diameter,
            'intersecting_elements': intersecting_elements
        }

    # End of process_selected_elements (debugging output removed)

start_drag_select_mode_and_finish()
level_collector = DB.FilteredElementCollector(doc).OfClass(DB.Level)
levels_dict = {lvl.Id: lvl for lvl in level_collector}

if family_symbols_dict:
    for fname, symbol_ids in family_symbols_dict.items():
        if symbol_ids:
            first_symbol_id = ElementId(int(symbol_ids[0]))
            first_symbol = doc.GetElement(first_symbol_id)
            is_face_based = (
                first_symbol.Family.get_Parameter(DB.BuiltInParameter.FAMILY_WORK_PLANE_BASED).AsInteger() == 1
            )
            if not first_symbol.IsActive:
                t_act = DBTransaction(doc, "Activate Family Symbol")
                t_act.Start()
                first_symbol.Activate()
                t_act.Commit()

            def place_family_instance_at_location(equip_element, first_symbol, face, location_point):
                adjusted_location = location_point
                face_normal = face.ComputeNormal(DB.UV(0.5, 0.5))
                reference_direction = face_normal.CrossProduct(DB.XYZ.BasisX)
                if reference_direction.IsZeroLength():
                    reference_direction = face_normal.CrossProduct(DB.XYZ.BasisY)
                reference_direction = reference_direction.Normalize()
                new_instance = doc.Create.NewFamilyInstance(face.Reference, adjusted_location, reference_direction, first_symbol)
                return new_instance

            placed_instance_count = 0
            with revit.Transaction("Place Family Instances"):
                for element in element_collector:
                    if isinstance(element, DB.FamilyInstance):
                        try:
                            loc = element.Location
                            if not isinstance(loc, DB.LocationPoint):
                                continue
                            location_point = loc.Point
                            host_data = family_instance_data.get(element.Id, {})
                            sleeve_diameter = host_data.get('sleeve_diameter')
                            intersections = host_data.get('intersecting_elements', [])
                            if not intersections:
                                continue
                            if isinstance(sleeve_diameter, str) and "mm" in sleeve_diameter:
                                sleeve_diameter = float(sleeve_diameter.replace(" mm", ""))
                            else:
                                sleeve_diameter = 0.0

                            instance_placed = False
                            for intersect in intersections:
                                if intersect['type'] == 'Beam':
                                    host = doc.GetElement(intersect['id'])
                                    if host:
                                        beam_type = doc.GetElement(host.GetTypeId())
                                        if beam_type:
                                            width_param = beam_type.LookupParameter("b")
                                            if not width_param:
                                                width_param = beam_type.LookupParameter("B")
                                            if width_param and width_param.StorageType == StorageType.Double:
                                                beam_width = width_param.AsDouble() * 304.8  # in mm
                                            else:
                                                continue
                                        else:
                                            continue
                                        geom_options = DB.Options()
                                        geom_options.ComputeReferences = True
                                        geo_element = host.get_Geometry(geom_options)
                                        equip_bbox = element.get_BoundingBox(None)
                                        far_end_point = get_opposite_face_of_equipment(element)
                                        face = find_intersecting_face_based_on_far_end(geo_element, far_end_point, equip_bbox)
                                        if face and face.Reference:
                                            new_instance = place_family_instance_at_location(element, first_symbol, face, location_point)
                                            for param in new_instance.Parameters:
                                                param_name = param.Definition.Name
                                                if "Length" in param_name and param.StorageType == StorageType.Double:
                                                    param.Set(beam_width / 304.8)  # Convert mm to feet.
                                                if "Outer Diameter" in param_name and param.StorageType == StorageType.Double:
                                                    param.Set((sleeve_diameter + 2) / 304.8)
                                            placed_instance_count += 1
                                            instance_placed = True
                                            break
                            if not instance_placed and intersections:
                                fallback_host = doc.GetElement(intersections[0]['id'])
                                new_instance = doc.Create.NewFamilyInstance(
                                    far_end_point, first_symbol, fallback_host, DB.Structure.StructuralType.NonStructural
                                )
                                placed_instance_count += 1
                        except Exception as e:
                            pass
            TaskDialog.Show("Sleeves Placement", "Total Sleeves Placed: {}".format(placed_instance_count))
