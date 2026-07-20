# Export Painted Faces

Exports faces painted with Revit's Paint tool to Rhino-readable files.

The recommended Rhino-editing workflow is DXF boundary export. OBJ is kept as an optional visual/material reference because OBJ faces are triangulated mesh.

## Output

- `.dxf` painted-face boundary curves
- `.obj` mesh model, optional triangulated reference
- `.mtl` material library for OBJ
- `_summary.csv` face/material summary

Rhino can open the DXF or OBJ directly. The OBJ references the MTL file, so keep both files in the same folder.

## What Gets Exported

The tool only exports faces that have a material assigned by Revit's Paint tool.

It does not export:

- normal wall/floor type materials
- category materials
- unpainted faces
- texture image assets from Appearance settings

The MTL file includes material names, diffuse colors, and transparency where available.

The DXF file puts each Paint material on its own layer. Layer names are ASCII-safe material keys; the CSV keeps the original Revit material names.

For Revit Split Face regions, the exporter reads `Face.GetRegions()` and exports each region boundary separately. Region material is read from Paint when available, with `Face.MaterialElementId` as a fallback because Split Face paint can be exposed there by the Revit API.

## Workflow

1. Run `Export Painted Faces`.
2. Choose export scope:
   - selected elements
   - visible elements in the active view
   - entire model
3. Choose export format:
   - DXF boundaries only, recommended for Rhino editing
   - OBJ mesh only, triangulated reference
   - both DXF and OBJ
4. Pick an output folder.
5. Enter a file name.
6. Open the DXF in Rhino for boundary editing, or OBJ only as mesh reference.

## Notes

- Coordinates are exported in millimeters.
- Each painted Revit face edge loop becomes a closed 3D polyline in the DXF.
- Each painted Revit face becomes a triangulated mesh group in the OBJ.
- Split Face regions are exported as separate boundaries instead of one large host face.
- Curved painted faces are triangulated.
- Curved painted face boundaries are tessellated into polyline segments in the DXF.
- For Rhino surface reconstruction, import DXF, select loops by material layer, then use `Join` and `PlanarSrf`.
- If a face is painted but Revit does not expose it through `GetPaintedMaterial`, it may be skipped.
