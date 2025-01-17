using CADability.Attribute;
using CADability.GeoObject;
using System;
using System.Collections.Generic;
using netDxf;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using ACadSharp.XData;
using CSMath;
using Color = System.Drawing.Color;
using Layer = ACadSharp.Tables.Layer;
using Transparency = ACadSharp.Transparency;

namespace CADability.DXF
{
    [Obsolete("ACadSharp export is experimental.")]
    public class Export2
    {
        private CadDocument doc;
        Dictionary<LinePattern, LineType> createdLinePatterns;
        int anonymousBlockCounter;
        double triangulationPrecision = 0.1;
        public Export2(ACadVersion version)
        {
            doc = new CadDocument(version);
            createdLinePatterns = new Dictionary<LinePattern, LineType>();
        }
        // Can we export to bytes/steams in Acadsharp?
        // public byte[] WriteToByteArray(Project toExport)
        // {
        //     var memoryStream = new System.IO.MemoryStream();
        //     Model modelSpace = null;
        //     if (toExport.GetModelCount() == 1) modelSpace = toExport.GetModel(0);
        //     else
        //     {
        //         modelSpace = toExport.FindModel("*Model_Space");
        //     }
        //     if (modelSpace == null) modelSpace = toExport.GetActiveModel();
        //     for (int i = 0; i < modelSpace.Count; i++)
        //     {
        //         EntityObject[] entity = GeoObjectToEntity(modelSpace[i]);
        //         if (entity != null) doc.Entities.Add(entity);
        //     }
        //     doc.Save(memoryStream);
        //     return memoryStream.ToArray();
        // }

        public void WriteToFile(Project toExport, string filename)
        {
            Model modelSpace;
            if (toExport.GetModelCount() == 1) modelSpace = toExport.GetModel(0);
            else
            {
                modelSpace = toExport.FindModel("*Model_Space");
            }
            if (modelSpace == null) modelSpace = toExport.GetActiveModel();
            GeoObjectList geoObjects = new GeoObjectList();
            List<Face> faces = new List<Face>(); // all top level faces collected
            for (int i = 0; i < modelSpace.Count; i++)
            {
                if (modelSpace[i] is Face face) faces.Add(face.Clone() as Face); // Clone() to keep ownership
                else geoObjects.Add(modelSpace[i]);
            }
            if (faces.Count > 0) geoObjects.Add(Shell.FromFaces(faces.ToArray())); // this shell is only to combine same color faces to a single mesh
            for (int i = 0; i < geoObjects.Count; i++)
            {
                Entity[] entities = GeoObjectToEntity(geoObjects[i]);
                if (entities != null)
                {
                    foreach (Entity entity in entities)
                    {
                        doc.Entities.Add(entity);
                    }
                }
            }
            using (DxfWriter writer = new DxfWriter(filename, doc, false))
            {
                writer.Write();
            }
        }

        private Entity[] GeoObjectToEntity(IGeoObject geoObject)
        {
            Entity entity = null;
            Entity[] entities = null;
            switch (geoObject)
            {
                case GeoObject.Point point: entity = ExportPoint(point); break;
                case GeoObject.Line line: entity = ExportLine(line); break;
                case GeoObject.Ellipse elli: entity = ExportEllipse(elli); break;
                case GeoObject.Polyline polyline: entity = ExportPolyline(polyline); break;
                case BSpline bspline: entity = ExportBSpline(bspline); break;
                case Path path:
                    if (Settings.GlobalSettings.GetBoolValue("DxfExport.ExportPathsAsBlocks", true))
                    {
                        entity = ExportPath(path);
                    }
                    else
                    {
                        entities = ExportPathWithoutBlock(path);
                    }
                    break;
                case Text text: entity = ExportText(text); break;
                case Block block: entity = ExportBlock(block); break;
                case Face face: entity = ExportFace(face); break;
                case Shell shell: entities = ExportShell(shell); break;
                case GeoObject.Solid solid: entities = ExportShell(solid.Shells[0]); break;
            }
            if (entity != null)
            {
                SetAttributes(entity, geoObject);
                SetUserData(entity, geoObject);
                return new Entity[] { entity };
            }
            if (entities != null)
            {
                for (int i = 0; i < entities.Length; i++)
                {
                    if (geoObject.Layer == null)
                    {
                        entities[i].Layer = Layer.Default;
                        continue;
                    }
                    if (!doc.Layers.TryGetValue(geoObject.Layer.Name, out Layer layer))
                    {
                        layer = new Layer(geoObject.Layer.Name);
                        doc.Layers.Add(layer);
                        //createdLayers[geoObject.Layer] = layer;
                    }
                    entities[i].Layer = layer;
                }
                return entities;
            }
            return null;
        }

        private void SetUserData(Entity entity, IGeoObject go)
        {
            if (entity is null || go is null || go.UserData is null || go.UserData.Count == 0)
                return;

            foreach (KeyValuePair<string, object> de in go.UserData)
            {
                if (de.Value is ExtendedEntityData xData)
                {
                    
                    if (!doc.AppIds.TryGetValue(xData.ApplicationName, out AppId xDataApp))
                    {
                        xDataApp = new AppId(xData.ApplicationName);
                        doc.AppIds.Add(xDataApp);
                    }
                    var xDataRecords = new List<ExtendedDataRecord>();
                    foreach (KeyValuePair<XDataCode, object> item in xData.Data)
                    {
                        // Convert from netDxf XDataCode enum to preserve backward compatibility
                        DxfCode code = (DxfCode)item.Key;
                        ExtendedDataRecord newValue;
                        // TODO: Should we use explicit cast with try/catch or safe cast and return null?
                        // TODO: Break this off to a SetXData() method for better readability
                        try
                        {
                            switch (code)
                            {
                                case DxfCode.ExtendedDataAsciiString:
                                    newValue = new ExtendedDataString((string)item.Value);
                                    break;
                                case DxfCode.ExtendedDataControlString:
                                    newValue = new ExtendedDataControlString((bool)item.Value);
                                    break;
                                case DxfCode.ExtendedDataLayerName:
                                    newValue = new ExtendedDataLayer((ulong)item.Value);
                                    break;
                                case DxfCode.ExtendedDataBinaryChunk:
                                    newValue = new ExtendedDataBinaryChunk((byte[])item.Value);
                                    break;
                                case DxfCode.ExtendedDataHandle:
                                    newValue = new ExtendedDataHandle((ulong)item.Value);
                                    break;
                                // xData coordinates are saved as individual axes but each takes a full XYZ
                                // Setting unused axes to 0?
                                case DxfCode.ExtendedDataXCoordinate:
                                    newValue = new ExtendedDataCoordinate(new XYZ((double)item.Value, 0, 0));
                                    break;
                                case DxfCode.ExtendedDataYCoordinate:
                                    newValue = new ExtendedDataCoordinate(new XYZ(0, (double)item.Value, 0));
                                    break;
                                case DxfCode.ExtendedDataZCoordinate:
                                    newValue = new ExtendedDataCoordinate(new XYZ(0, 0, (double)item.Value));
                                    break;
                                case DxfCode.ExtendedDataWorldXCoordinate:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ((double)item.Value, 0, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldYCoordinate:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, (double)item.Value, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldZCoordinate:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, 0, (double)item.Value));
                                    break;
                                case DxfCode.ExtendedDataWorldXDisp:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ((double)item.Value, 0, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldYDisp:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, (double)item.Value, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldZDisp:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, 0, (double)item.Value));
                                    break;
                                case DxfCode.ExtendedDataWorldXDir:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ((double)item.Value, 0, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldYDir:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, (double)item.Value, 0));
                                    break;
                                case DxfCode.ExtendedDataWorldZDir:
                                    newValue = new ExtendedDataWorldCoordinate(new XYZ(0, 0, (double)item.Value));
                                    break;
                                case DxfCode.ExtendedDataReal:
                                    newValue = new ExtendedDataReal((double)item.Value);
                                    break;
                                case DxfCode.ExtendedDataDist:
                                    newValue = new ExtendedDataDistance((double)item.Value);
                                    break;
                                case DxfCode.ExtendedDataScale:
                                    newValue = new ExtendedDataScale((double)item.Value);
                                    break;
                                case DxfCode.ExtendedDataInteger16:
                                    newValue = new ExtendedDataInteger16((short)item.Value);
                                    break;
                                case DxfCode.ExtendedDataInteger32:
                                    newValue = new ExtendedDataInteger32((int)item.Value);
                                    break;
                                default:
                                    // How gracefully should we fail on unknown DxfCodes?
#if(DEBUG)
                                    throw (new Exception("Unknown DxfCode: " + code));
#endif
                            }
                        }
                        catch (InvalidCastException e)
                        {
#if(DEBUG)
                            throw (e);
#endif
                        }

                        if (newValue != null)
                        {
                            xDataRecords.Add(newValue);
                        }
                    }
                    entity.ExtendedData.Add(xDataApp, xDataRecords);
                }
                else
                {
                    var xDataApp = new AppId(de.Key);
                    var xDataRecords = new List<ExtendedDataRecord>();
                
                    ExtendedDataRecord record = null;
                
                    switch (de.Value)
                    {
                        case string strValue:
                            record = new ExtendedDataString(strValue);
                            break;
                        case short shrValue:
                            record = new ExtendedDataInteger16(shrValue);
                            break;
                        case int intValue:
                            record = new ExtendedDataInteger32(intValue);
                            break;
                        case double dblValue:
                            record = new ExtendedDataReal(dblValue);
                            break;
                        case byte[] bytValue:
                            record = new ExtendedDataBinaryChunk(bytValue);
                            break;
                    }

                    if (record != null)
                    {
                        xDataRecords.Add(record);
                        entity.ExtendedData.Add(xDataApp, xDataRecords);
                    }
                }
            }
        }

        private Entity[] ExportShell(Shell shell)
        {
            // TODO: Export as full multi-face polyfacemesh instead of array of single-face polyfacemeshes?
            // Look at Shell type and see how it can better be translated to a Dxf entitity
            if (Settings.GlobalSettings.GetBoolValue("DxfImport.SingleMeshPerFace", false))
            {
                List<Entity> res = new List<Entity>();
                for (int i = 0; i < shell.Faces.Length; i++)
                {
                    Entity mesh = ExportFace(shell.Faces[i]);
                    if (mesh != null) res.Add(mesh);
                }
                return res.ToArray();
            }
            else
            {
                List<Entity> res = new List<Entity>();
                Dictionary<int, (List<Vector3>, List<short[]>)> mesh = new Dictionary<int, (List<Vector3>, List<short[]>)>();
                for (int i = 0; i < shell.Faces.Length; i++)
                {
                    CollectMeshByColor(mesh, shell.Faces[i]);
                }
                foreach (KeyValuePair<int, (List<Vector3> vertices, List<short[]> indices)> item in mesh)
                {
                    PolyfaceMesh pfm = new PolyfaceMesh();
                    foreach (var vert in item.Value.vertices)
                    {
                        pfm.Vertices.Add(new Vertex3D() { Location = new XYZ(vert.X, vert.Y, vert.Z) });
                    }
                    VertexFaceRecord faceRecord = new VertexFaceRecord();
                    if (item.Value.indices.Count >= 3)
                    {
                        faceRecord.Index1 = item.Value.indices[0][0];
                        faceRecord.Index2 = item.Value.indices[0][1];
                        faceRecord.Index3 = item.Value.indices[0][2];
                    }

                    if (item.Value.indices.Count == 4)
                    {
                        faceRecord.Index4 = item.Value.indices[0][3];
                    } pfm.Faces.Add(faceRecord);
                    SetColor(pfm, item.Key);
                    res.Add(pfm);
                }
                return res.ToArray();
            }
        }
        private Entity ExportFace(Face face)
        {
            // TODO: Lots of duplicate code in exporting meshes and polyfacemeshes. Condense to new methods.
            if (Settings.GlobalSettings.GetBoolValue("DxfImport.UseMesh", false))
            {
                if (face.Surface is PlaneSurface)
                {
                    if (face.OutlineEdges.Length == 4 && face.OutlineEdges[0].Curve3D is GeoObject.Line && face.OutlineEdges[1].Curve3D is GeoObject.Line && face.OutlineEdges[2].Curve3D is GeoObject.Line && face.OutlineEdges[3].Curve3D is GeoObject.Line)
                    {
                        // 4 lines, export as a simple PolyfaceMesh with 4 vertices
                        List<XYZ> vertices = new List<XYZ>();
                        for (int i = 0; i < 4; i++)
                        {
                            vertices.Add(new XYZ(face.OutlineEdges[i].StartVertex(face).Position.x, face.OutlineEdges[i].StartVertex(face).Position.y, face.OutlineEdges[i].StartVertex(face).Position.z));
                        }
                        Mesh res = new Mesh()
                        {
                            Vertices = vertices,
                            Faces = new List<int[]> { new int[] { 0, 1, 2, 3 } },
                        };
                        SetAttributes(res, face);
                        return res;
                    }
                }
                {
                    face.GetTriangulation(triangulationPrecision, out GeoPoint[] trianglePoint, out _, out int[] triangleIndex, out _);
                    List<int[]> indices = new List<int[]>();
                    for (int i = 0; i < triangleIndex.Length - 2; i += 3)
                    {   // it is strange, but the indices must be +1
                        indices[i / 3] = new int[] { triangleIndex[i], triangleIndex[i + 1], triangleIndex[i + 2] };
                    }
                    List<XYZ> vertices = new List<XYZ>();
                    foreach (var item in trianglePoint)
                    {
                        vertices.Add(new XYZ(item.x, item.y, item.z));
                    }

                    Mesh res = new Mesh()
                    {
                        Vertices = vertices,
                        Faces = indices,
                    };
                    
                    SetAttributes(res, face);
                    return res;
                }
            }
            else
            {
                if (face.Surface is PlaneSurface)
                {
                    if (face.OutlineEdges.Length == 4 && face.OutlineEdges[0].Curve3D is GeoObject.Line && face.OutlineEdges[1].Curve3D is GeoObject.Line && face.OutlineEdges[2].Curve3D is GeoObject.Line && face.OutlineEdges[3].Curve3D is GeoObject.Line)
                    {
                        PolyfaceMesh res = new PolyfaceMesh();
                        // 4 lines, export as a simple PolyfaceMesh with 4 vertices
                        foreach (var item in face.OutlineEdges)
                        {
                            var vert = new Vertex3D();
                            vert.Location = new XYZ(item.StartVertex(face).Position.x, item.StartVertex(face).Position.y, item.StartVertex(face).Position.z);
                            res.Vertices.Add(vert);
                        }
                        var vtexFaceRecord = new VertexFaceRecord()
                        {
                            Index1 = 1,
                            Index2 = 2,
                            Index3 = 3,
                            Index4 = 4,
                        };
                        res.Faces.Add(vtexFaceRecord);
                        SetAttributes(res, face);
                        return res;
                    }
                }
                {
                    PolyfaceMesh res = new PolyfaceMesh();
                    face.GetTriangulation(triangulationPrecision, out GeoPoint[] trianglePoint, out _, out int[] triangleIndex, out _);
                    List<int[]> indices = new List<int[]>();
                    for (int i = 0; i < triangleIndex.Length - 2; i += 3)
                    {   // it is strange, but the indices must be +1
                        indices[i / 3] = new int[] { triangleIndex[i], triangleIndex[i + 1], triangleIndex[i + 2] };
                    }
                    foreach (var item in trianglePoint)
                    {
                        res.Vertices.Add(new Vertex3D() { Location = new XYZ(item.x, item.y, item.z) });
                    }

                    var vtexFaceRecord = new VertexFaceRecord()
                    {
                        Index1 = 1,
                        Index2 = 2,
                        Index3 = 3,
                    };
                    
                    res.Faces.Add(vtexFaceRecord);
                    SetAttributes(res, face);
                    return res;
                }
            }
        }
        private void CollectMeshByColor(Dictionary<int, (List<Vector3>, List<short[]>)> mesh, Face face)
        {
            if (!mesh.TryGetValue(face.ColorDef.Color.ToArgb(), out (List<Vector3> vertices, List<short[]> indices) mc))
            {
                mesh[face.ColorDef.Color.ToArgb()] = mc = (new List<Vector3>(), new List<short[]>());
            }
            short offset = (short)(mc.vertices.Count + 1);
            face.GetTriangulation(triangulationPrecision, out GeoPoint[] trianglePoint, out _, out int[] triangleIndex, out _);
            short[][] indices = new short[triangleIndex.Length / 3][];
            for (int i = 0; i < triangleIndex.Length - 2; i += 3)
            {   // it is strange, but the indices must be +1
                indices[i / 3] = new short[] { (short)(triangleIndex[i] + offset), (short)(triangleIndex[i + 1] + offset), (short)(triangleIndex[i + 2] + offset) };
            }
            mc.indices.AddRange(indices);
            Vector3[] vertices = new Vector3[trianglePoint.Length];
            for (int i = 0; i < trianglePoint.Length; i++)
            {
                vertices[i] = Vector3(trianglePoint[i]);
            }
            mc.vertices.AddRange(vertices);
        }
        private TextEntity ExportText(Text textGeoObj)
        {
            string textValue = textGeoObj.TextString;
            // TODO: Use Text.FourPoints to use TextEntity.AlignmentPoint for rotation?
            FontFlags fontFlags = FontFlags.Regular;
            if (textGeoObj.ColorDef == null) textGeoObj.ColorDef = new ColorDef();
            
            if (textGeoObj.Bold) fontFlags |= FontFlags.Bold;
            if (textGeoObj.Italic) fontFlags |= FontFlags.Italic;
            if (textGeoObj.Underline) textValue = textValue.Insert(0, "%%U");
            if (textGeoObj.Strikeout) textValue = textValue.Insert(0, "%%K");
            // TODO: Calculate oblique angle from glyph direction and line direction

            TextStyle textStyleDxf = new TextStyle(textGeoObj.Font)
            {
                Filename = textGeoObj.Font,
            };

            if (System.IO.Path.GetExtension(textGeoObj.Font).ToLower() == ".ttf")
            {
                textStyleDxf.TrueType = fontFlags;
            }

            foreach (TextStyle textStyle in doc.TextStyles)
            {
                if (textStyleDxf.Filename == textStyle.Filename && textStyleDxf.TrueType == textStyle.TrueType)
                {
                    textStyleDxf = textStyle;
                    break;
                }
            }

            switch (textGeoObj.ColorDef.Source)
            {
                case ColorDef.ColorSource.fromObject:
                    break;
                case ColorDef.ColorSource.fromParent:
                    break;
                case ColorDef.ColorSource.fromName:
                    break;
                case ColorDef.ColorSource.fromStyle:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            
            TextHorizontalAlignment horizontalAignment = TextHorizontalAlignment.Left;
            TextVerticalAlignmentType verticalAlignment = TextVerticalAlignmentType.Baseline;

            switch (textGeoObj.LineAlignment)
            {
                case Text.LineAlignMode.Left:
                    horizontalAignment = TextHorizontalAlignment.Left;
                    break;
                case Text.LineAlignMode.Center:
                    horizontalAignment = TextHorizontalAlignment.Center;
                    break;
                case Text.LineAlignMode.Right:
                    horizontalAignment = TextHorizontalAlignment.Right;
                    break;
            }

            switch (textGeoObj.Alignment)
            {
                case Text.AlignMode.Baseline:
                    verticalAlignment = TextVerticalAlignmentType.Baseline;
                    break;
                case Text.AlignMode.Bottom:
                    verticalAlignment = TextVerticalAlignmentType.Bottom;
                    break;
                case Text.AlignMode.Center:
                    verticalAlignment = TextVerticalAlignmentType.Middle;
                    break;
                case Text.AlignMode.Top:
                    verticalAlignment = TextVerticalAlignmentType.Top;
                    break;
            }
            
            TextEntity textEnt = new TextEntity()
            {
                Value = textValue,
                InsertPoint = new XYZ(textGeoObj.Location.x, textGeoObj.Location.y, textGeoObj.Location.z),
                AlignmentPoint = new XYZ(textGeoObj.Location.x, textGeoObj.Location.y, textGeoObj.Location.z),
                Normal = new XYZ(textGeoObj.Plane.Normal.x, textGeoObj.Plane.Normal.y, textGeoObj.Plane.Normal.z),
                Height = textGeoObj.TextSize,
                Rotation = new Angle(textGeoObj.LineDirection.To2D()),
                VerticalAlignment = verticalAlignment,
                HorizontalAlignment = horizontalAignment,
                Style = textStyleDxf,
            };
            return textEnt;

            // // Old code below. Preserved for troubleshooting positioning
            // System.Drawing.FontStyle fs = System.Drawing.FontStyle.Regular;
            // if (textGeoObj.Bold) fs |= System.Drawing.FontStyle.Bold;
            // if (textGeoObj.Italic) fs |= System.Drawing.FontStyle.Italic;
            // System.Drawing.Font font = new System.Drawing.Font(textGeoObj.Font, 1000.0f, fs);
            // // Old font contructor. May have useful signature
            // //netDxf.Entities.Text res = new netDxf.Entities.Text(text.TextString, Vector2.Zero, text.TextSize * 1000 / font.Height, new TextStyle(text.Font, text.Font + ".ttf"));
            // ModOp toText = ModOp.Fit(GeoPoint.Origin, new GeoVector[] { GeoVector.XAxis, GeoVector.YAxis, GeoVector.ZAxis }, textGeoObj.Location, new GeoVector[] { textGeoObj.LineDirection.Normalized, textGeoObj.GlyphDirection.Normalized, textGeoObj.LineDirection.Normalized ^ textGeoObj.GlyphDirection.Normalized });
            // res.TransformBy(Matrix4(toText)); // easier than setting normal and rotation
            // return res;
        }

        private Insert ExportBlock(Block blk)
        {
            string name = blk.Name;
            // What's TableObject. Do I need to validate names?
            char[] invalidCharacters = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', ';', ',', '=', '`' };
            if (name == null || doc.BlockRecords.Contains(name) || name.IndexOfAny(invalidCharacters) > -1) name = GetNextAnonymousBlockName();
            BlockRecord acBlock = new BlockRecord(name);
            foreach (IGeoObject entity in blk.Children)
            {
                acBlock.Entities.AddRange(GeoObjectToEntity(entity));
            }
            doc.BlockRecords.Add(acBlock);
            return new Insert(acBlock);
        }
        private Insert ExportPath(Path path)
        {
            BlockRecord acBlock = new BlockRecord(GetNextAnonymousBlockName());
            foreach (ICurve curve in path.Curves)
            {
                Entity[] curveEntity = GeoObjectToEntity(curve as IGeoObject);
                if (curve != null) acBlock.Entities.AddRange(curveEntity);
            }
            doc.BlockRecords.Add(acBlock);
            return new Insert(acBlock);
        }

        private Entity[] ExportPathWithoutBlock(Path path)
        {
            List<Entity> entities = new List<Entity>();
            for (int i = 0; i < path.Curves.Length; i++)
            {
                Entity[] curve = GeoObjectToEntity(path.Curves[i] as IGeoObject);
                if (curve != null) entities.AddRange(curve);
            }
            return entities.ToArray();
        }
        
        private Spline ExportBSpline(BSpline bspline)
        {
            Spline spline = new Spline();

            if (bspline.ThroughPointCount > 0)
            {
                foreach (GeoPoint fitPoint in bspline.ThroughPoint)
                {
                    spline.FitPoints.Add(new XYZ(fitPoint.x, fitPoint.y, fitPoint.z));
                }

            }
            foreach (GeoPoint controlPoint in bspline.Poles)
            {
                spline.ControlPoints.Add(new XYZ(controlPoint.x, controlPoint.y, controlPoint.z));
            }

            foreach (double weight in bspline.Weights)
            {
                spline.Weights.Add(weight);
            }

            for (int i = 0; i < bspline.Knots.Length; i++)
            {
                for (int j = 0; j < bspline.Multiplicities[i]; j++) spline.Knots.Add(bspline.Knots[i]);
            }

            return spline;
        }

        private Polyline3D ExportPolyline(GeoObject.Polyline goPolyline)
        {
            //TODO: Check if a new method for Polyline2D (old LwPolyline) is necessary
            Polyline3D polyline = new Polyline3D();
            for (int i = 0; i < goPolyline.Vertices.Length; i++)
            {
                Vertex3D vert = new Vertex3D();
                vert.Location = new XYZ(goPolyline.Vertices[i].x, goPolyline.Vertices[i].y, goPolyline.Vertices[i].z);
                polyline.Vertices.Add(vert);
            }

            if (goPolyline.IsClosed)
            {
                Vertex2D endVert = new Vertex2D();
                endVert.Location = new XYZ(goPolyline.Vertices[0].x, goPolyline.Vertices[0].y, goPolyline.Vertices[0].z);
                polyline.Vertices.Add(endVert);
            }
            return polyline;
        }

        private ACadSharp.Entities.Point ExportPoint(GeoObject.Point point)
        {
            return new ACadSharp.Entities.Point(new XYZ(point.Location.x, point.Location.y, point.Location.z));
        }
        private ACadSharp.Entities.Line ExportLine(GeoObject.Line line)
        {
            XYZ startPoint = new XYZ(line.StartPoint.x, line.StartPoint.y, line.StartPoint.z);
            XYZ endPoint = new XYZ(line.EndPoint.x, line.EndPoint.y, line.EndPoint.z);
            return new ACadSharp.Entities.Line(startPoint, endPoint);
        }
        private Entity ExportEllipse(GeoObject.Ellipse elli)
        {
            Entity entity;
            if (elli.IsArc)
            {
                Plane dxfPlane = Import2.Plane(Vector3(elli.Center), Vector3(elli.Plane.Normal));
                // TODO: Test removing this and inverting negative start/eng angle params below
                // This will correct clockwise arcs exporting backwards or with an inverted Z normal
                if (!elli.CounterClockWise) (elli.StartPoint, elli.EndPoint) = (elli.EndPoint, elli.StartPoint);
                if (elli.IsCircle)
                {
                    GeoObject.Ellipse aligned = GeoObject.Ellipse.Construct();
                    aligned.SetArcPlaneCenterStartEndPoint(dxfPlane, dxfPlane.Project(elli.Center), dxfPlane.Project(elli.StartPoint), dxfPlane.Project(elli.EndPoint), dxfPlane, true);

                    //Check if ellipse is actually a closed circle
                    if (Math.Abs(elli.SweepParameter) > Math.PI && Precision.IsEqual(elli.StartPoint, elli.EndPoint))
                    {
                        entity = new Circle
                        {
                            Normal = new XYZ(dxfPlane.Normal.x, dxfPlane.Normal.y, dxfPlane.Normal.z),
                            Center = new XYZ(elli.Center.x, elli.Center.y, elli.Center.z),
                            Radius = elli.Radius
                        };
                    }
                    else
                    {
                        entity = new Arc
                        {
                            Normal = new XYZ(dxfPlane.Normal),
                            Center = new XYZ(aligned.Center),
                            Radius = aligned.Radius,
                            StartAngle = aligned.StartParameter, // Radians
                            EndAngle = aligned.StartParameter + aligned.SweepParameter, // Radians
                        };
                    }
                }
                else
                {
                    ACadSharp.Entities.Ellipse expelli = new ACadSharp.Entities.Ellipse()
                    {
                        Normal = new XYZ(elli.Plane.Normal),
                        Center = new XYZ(elli.Center.x, elli.Center.y, elli.Center.z),
                        RadiusRatio = elli.MinorRadius / elli.MajorRadius,
                        StartParameter = elli.StartParameter,
                        EndParameter = elli.SweepParameter + elli.StartParameter,
                        EndPoint = new XYZ(elli.MajorAxis.x, elli.MajorAxis.y, elli.MajorAxis.z),
                    };
                    // We cannot have mismatched signs on start and end params since DXF has no concept of CW/CCW
                    // Find arcs with mismatched signs and flip the negative value to a positive on the same vector
                    if (Math.Sign(expelli.StartParameter) != Math.Sign(expelli.EndParameter))
                    {
                        if (expelli.StartParameter < 0) expelli.StartParameter += 2 * Math.PI;
                        if (expelli.EndParameter < 0) expelli.EndParameter += 2 * Math.PI;
                    }
                    entity = expelli;
                }
            }
            else
            {
                if (elli.IsCircle)
                {
                    entity = new Circle()
                    {
                        Normal = new XYZ(elli.Normal.x, elli.Normal.y, elli.Normal.z),
                        Center = new XYZ(elli.Center.x, elli.Center.y, elli.Center.z),
                        Radius = elli.Radius
                    };
                }
                else
                {
                    ACadSharp.Entities.Ellipse expelli = new ACadSharp.Entities.Ellipse()
                    {
                        Normal = new XYZ(elli.Plane.Normal.x, elli.Plane.Normal.y, elli.Plane.Normal.z),
                        Center = new XYZ(elli.Center.x, elli.Center.y, elli.Center.z),
                        RadiusRatio = elli.MinorRadius / elli.MajorRadius,
                        StartParameter = elli.StartParameter,
                        EndParameter = elli.SweepParameter + elli.StartParameter,
                        EndPoint = new XYZ(elli.MajorAxis.x, elli.MajorAxis.y, elli.MajorAxis.z),
                    };
                    
                    entity = expelli;
                }
            }
            return entity;
        }

        private void SetColor(Entity entity, int colorIndex)
        {
            if (colorIndex.Equals(Color.White.ToArgb()) || colorIndex.Equals(Color.Black.ToArgb()))
            {
                // White/black index
                entity.Color = new ACadSharp.Color(7);
            }
            else
            {
                if (colorIndex > 0 && colorIndex < 256)
                {
                    entity.Color = new ACadSharp.Color((short)colorIndex);
                }
                else if (colorIndex > 255)
                {
                    entity.Color = ACadSharp.Color.FromTrueColor((uint)colorIndex);
                }
            }
        }
        private void SetAttributes(Entity entity, IGeoObject go)
        {
            if (go is IColorDef cd && cd.ColorDef != null)
            {
                ACadSharp.Color color = new ACadSharp.Color();
                switch (cd.ColorDef.Source)
                {
                    case ColorDef.ColorSource.fromParent:
                        color = new ACadSharp.Color(ACadSharp.Color.ByBlock.Index);
                        entity.Transparency = Transparency.ByBlock;
                        break;
                    case ColorDef.ColorSource.fromStyle:
                        color = new ACadSharp.Color(ACadSharp.Color.ByLayer.Index);
                        entity.Transparency = Transparency.ByLayer;
                        break;
                    // Not sure what fromName is
                    case ColorDef.ColorSource.fromName:
                    case ColorDef.ColorSource.fromObject:
                        //Convert ARGB alpha channel to AutoCAD transparency
                        entity.Transparency = new Transparency((byte)(90 - (short)(cd.ColorDef.Color.A / 255.0 * 90)));
                        if (cd.ColorDef.Color.ToArgb().Equals(Color.White.ToArgb()) || cd.ColorDef.Color.ToArgb().Equals(Color.Black.ToArgb()))
                        {
                            // White/black index
                            color = new ACadSharp.Color(7);
                        }
                        else
                        {
                            color = new ACadSharp.Color(cd.ColorDef.Color.R, cd.ColorDef.Color.G, cd.ColorDef.Color.B);
                        }
                        break;
                }
                entity.Color = color;
            }
            if (go.Layer != null)
            {
                if (!doc.Layers.TryGetValue(go.Layer.Name, out Layer layer))
                {
                    // TODO: Copy more layer data like color?
                    layer = new Layer(go.Layer.Name);
                    doc.Layers.Add(layer);
                    //createdLayers[go.Layer] = layer;
                }
                entity.Layer = layer;
            }
            if (go is ILinePattern lp && lp.LinePattern != null)
            {
                if (!doc.LineTypes.TryGetValue(lp.LinePattern.Name, out LineType linetype))
                {
                    linetype = new LineType(lp.LinePattern.Name);
                    if (lp.LinePattern.Pattern != null)
                    {
                        for (int i = 0; i < lp.LinePattern.Pattern.Length; i++)
                        {
                            LineType.Segment ls = new LineType.Segment();
                            if ((i & 0x01) == 0) ls.Length = lp.LinePattern.Pattern[i];
                            else ls.Length = -lp.LinePattern.Pattern[i];
                            linetype.AddSegment(ls); 
                        }
                    }
                    doc.LineTypes.Add(linetype);
                    createdLinePatterns[lp.LinePattern] = linetype;
                }
                entity.LineType = linetype;
            }
            if (go is ILineWidth lw && lw.LineWidth != null)
            {
                double minError = double.MaxValue;
                LineweightType found = LineweightType.Default;
                foreach (LineweightType lwe in Enum.GetValues(typeof(LineweightType)))
                {
                    double err = Math.Abs(((int)lwe) / 100.0 - lw.LineWidth.Width);
                    if (err < minError)
                    {
                        minError = err;
                        found = lwe;
                    }
                }
                entity.LineWeight = found;
            }
        }
        
        private double CalcStartEndAngle(double startEndParameter, double majorAxis, double minorAxis)
        {
            double a = majorAxis * 0.5d;
            double b = minorAxis * 0.5d;
            Vector2 startPoint = new Vector2(a * Math.Cos(startEndParameter), b * Math.Sin(startEndParameter));
            return Vector2.Angle(startPoint) * netDxf.MathHelper.RadToDeg;
        }
        private Vector3 Vector3(GeoPoint p)
        {
            return new Vector3(p.x, p.y, p.z);
        }

        private Vector3 Vector3(GeoVector p)
        {
            return new Vector3(p.x, p.y, p.z);
        }

        private Vector3 Vector3(XYZ p)
        {
            return new Vector3(p.X, p.Y, p.Z);
        }
        private string GetNextAnonymousBlockName()
        {
            return "AnonymousBlock" + (++anonymousBlockCounter).ToString();
        }
    }
}