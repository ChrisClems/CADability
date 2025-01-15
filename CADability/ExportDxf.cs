using CADability.Attribute;
using CADability.GeoObject;
using MathNet.Numerics.Statistics.Mcmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Xml.XPath;
using netDxf;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using ACadSharp.XData;
using CSMath;
using Color = System.Drawing.Color;
using Layer = ACadSharp.Tables.Layer;

namespace CADability.DXF
{
    public class Export
    {
        private CadDocument doc;
        private Dictionary<CADability.Attribute.Layer, ACadSharp.Tables.Layer> createdLayers;
        Dictionary<CADability.Attribute.LinePattern, ACadSharp.Tables.LineType> createdLinePatterns;
        int anonymousBlockCounter = 0;
        double triangulationPrecision = 0.1;
        public Export(ACadSharp.ACadVersion version)
        {
            doc = new CadDocument(version);
            createdLayers = new Dictionary<Attribute.Layer, ACadSharp.Tables.Layer>();
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
            Model modelSpace = null;
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
                case GeoObject.BSpline bspline: entity = ExportBSpline(bspline); break;
                case GeoObject.Path path:
                    if (Settings.GlobalSettings.GetBoolValue("DxfExport.ExportPathsAsBlocks", true))
                    {
                        entity = ExportPath(path);
                    }
                    else
                    {
                        entities = ExportPathWithoutBlock(path);
                    }
                    break;
                case GeoObject.Text text: entity = ExportText(text); break;
                case GeoObject.Block block: entity = ExportBlock(block); break;
                case GeoObject.Face face: entity = ExportFace(face); break;
                case GeoObject.Shell shell: entities = ExportShell(shell); break;
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
                        entities[i].Layer = ACadSharp.Tables.Layer.Default;
                        continue;
                    }
                    if (!createdLayers.TryGetValue(geoObject.Layer, out ACadSharp.Tables.Layer layer))
                    {
                        layer = new ACadSharp.Tables.Layer(geoObject.Layer.Name);
                        doc.Layers.Add((Layer)layer);
                        createdLayers[geoObject.Layer] = layer;
                    }
                    entities[i].Layer = layer;
                }
                return entities;
            }
            return null;
        }

        private void SetUserData(ACadSharp.Entities.Entity entity, IGeoObject go)
        {
            if (entity is null || go is null || go.UserData is null || go.UserData.Count == 0)
                return;

            foreach (KeyValuePair<string, object> de in go.UserData)
            {
                var exDataApp = new AppId("TestAppId");
                var xDString = new ExtendedDataString("test");
                var xDList = new List<ExtendedDataString>(){ xDString };
                entity.ExtendedData.Add(exDataApp, xDList);
                
                // TODO: Do a deep dive into xData export.
                if (de.Value is ExtendedEntityData xData)
                {
                    var exDataApp
                    XData data = new XData(registry);
                
                    foreach (KeyValuePair<DxfCode, object> item in xData.Data)
                    {
                        DxfCode code = item.Key;
                        //XDataCode code = item.Key;
                        object newValue = null;
                        //Make the export more robust to wrong XDataCode entries.
                        //Try to fit the data into an existing datatype. Otherwise ignore entry.
                        // TODO: This looks wildly incomplete. Run through thorough tests
                        switch (code)
                        {
                            case DxfCode.Int16:
                                if (item.Value is short int16val_1)
                                    newValue = int16val_1;
                                else if (item.Value is int int32val_1 && int32val_1 >= Int16.MinValue && int32val_1 <= Int16.MaxValue)
                                    newValue = Convert.ToInt16(int32val_1);
                                else if (item.Value is long int64val_1 && int64val_1 >= Int16.MinValue && int64val_1 <= Int16.MaxValue)
                                    newValue = Convert.ToInt16(int64val_1);
                                break;
                            case DxfCode.Int32:
                                if (item.Value is short int16val_2)
                                    newValue = Convert.ToInt32(int16val_2);
                                else if (item.Value is int int32val_2)
                                    newValue = int32val_2;
                                else if (item.Value is long int64val_2 && int64val_2 >= Int32.MinValue && int64val_2 <= Int32.MaxValue)
                                    newValue = Convert.ToInt32(int64val_2);
                                break;
                            default:
                                newValue = item.Value;
                                break;
                        }
                        
                        if (newValue != null)
                        {
                            xData.Data.Add(code, newValue);
                            data.Records.AddRange(new ExtendedDataRecord<>()[] {});
                            XDataRecord record = new XDataRecord(code, newValue);
                            data.XDataRecord.Add(record);
                        }
                    }
                    if (data.XDataRecord.Count > 0)
                        entity.XData.Add(data);
                }
                else
                {
                    ApplicationRegistry registry = new ApplicationRegistry(ApplicationRegistry.DefaultName);
                    XData data = new XData(registry);
                
                    XDataRecord record = null;
                
                    switch (de.Value)
                    {
                        case string strValue:
                            record = new XDataRecord(XDataCode.String, strValue);
                            break;
                        case short shrValue:
                            record = new XDataRecord(XDataCode.Int16, shrValue);
                            break;
                        case int intValue:
                            record = new XDataRecord(XDataCode.Int32, intValue);
                            break;
                        case double dblValue:
                            record = new XDataRecord(XDataCode.Real, dblValue);
                            break;
                        case byte[] bytValue:
                            record = new XDataRecord(XDataCode.BinaryData, bytValue);
                            break;
                    }
                
                    if (record != null)
                        data.XDataRecord.Add(record);
                }
            }
        }

        private Entity[] ExportShell(GeoObject.Shell shell)
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
                        faceRecord.Index1 = (short)item.Value.indices[0][0];
                        faceRecord.Index2 = (short)item.Value.indices[0][1];
                        faceRecord.Index3 = (short)item.Value.indices[0][2];
                    }

                    if (item.Value.indices.Count == 4)
                    {
                        faceRecord.Index4 = (short)item.Value.indices[0][3];
                    } pfm.Faces.Add(faceRecord);
                    SetColor(pfm, item.Key);
                    res.Add((Entity)pfm);
                }
                return res.ToArray();
            }
        }
        private ACadSharp.Entities.Entity ExportFace(GeoObject.Face face)
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
                        ACadSharp.Entities.Mesh res = new Mesh()
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
                        ACadSharp.Entities.PolyfaceMesh res = new PolyfaceMesh();
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
                    ACadSharp.Entities.PolyfaceMesh res = new PolyfaceMesh();
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
            face.GetTriangulation(triangulationPrecision, out GeoPoint[] trianglePoint, out GeoPoint2D[] triangleUVPoint, out int[] triangleIndex, out BoundingCube triangleExtent);
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
        private ACadSharp.Entities.TextEntity ExportText(GeoObject.Text textGeoObj)
        {
            string textValue = textGeoObj.TextString;
            // TODO: Use Text.FourPoints to use TextEntity.AlignmentPoint for rotation?
            FontFlags fontFlags = FontFlags.Regular;
            
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
            
            TextEntity textEnt = new ACadSharp.Entities.TextEntity()
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
            SetAttributes(textEnt, textGeoObj);
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

        private ACadSharp.Entities.Insert ExportBlock(GeoObject.Block blk)
        {
            string name = blk.Name;
            // What's TableObject. Do I need to validate names?
            char[] invalidCharacters = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', ';', ',', '=', '`' };
            if (name == null || doc.BlockRecords.Contains(name) || name.IndexOfAny(invalidCharacters) > -1) name = GetNextAnonymousBlockName();
            BlockRecord acBlock = new ACadSharp.Tables.BlockRecord(name);
            foreach (IGeoObject entity in blk.Children)
            {
                acBlock.Entities.AddRange(GeoObjectToEntity(entity));
            }
            doc.BlockRecords.Add(acBlock);
            return new ACadSharp.Entities.Insert(acBlock);
        }
        private ACadSharp.Entities.Insert ExportPath(Path path)
        {
            BlockRecord acBlock = new ACadSharp.Tables.BlockRecord(GetNextAnonymousBlockName());
            foreach (ICurve curve in path.Curves)
            {
                Entity[] curveEntity = GeoObjectToEntity(curve as IGeoObject);
                if (curve != null) acBlock.Entities.AddRange(curveEntity);
            }
            doc.BlockRecords.Add(acBlock);
            return new ACadSharp.Entities.Insert(acBlock);
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
        
        private ACadSharp.Entities.Spline ExportBSpline(BSpline bspline)
        {
            Spline spline = new ACadSharp.Entities.Spline();

            foreach (GeoPoint fitPoint in bspline.ThroughPoint)
            {
                spline.FitPoints.Add(new XYZ(fitPoint.x, fitPoint.y, fitPoint.z));
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
                for (int j = 0; i < bspline.Multiplicities[i]; j++) spline.Knots.Add(bspline.Knots[i]);
            }

            return spline;
        }

        private ACadSharp.Entities.Polyline3D ExportPolyline(GeoObject.Polyline goPolyline)
        {
            //TODO: Check if a new method for Polyline2D (old LwPolyline) is necessary
            Polyline3D polyline = new ACadSharp.Entities.Polyline3D();
            for (int i = 0; i < goPolyline.Vertices.Length; i++)
            {
                Vertex2D vert = new ACadSharp.Entities.Vertex2D();
                vert.Location = new XYZ(goPolyline.Vertices[i].x, goPolyline.Vertices[i].y, goPolyline.Vertices[i].z);
                polyline.Vertices.Add(vert);
            }

            if (goPolyline.IsClosed)
            {
                Vertex2D endVert = new ACadSharp.Entities.Vertex2D();
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
            ACadSharp.Entities.Entity entity = null;
            if (elli.IsArc)
            {
                Plane dxfPlane = Import.Plane(Vector3(elli.Center), Vector3(new GeoPoint(elli.Plane.Normal.x, elli.Plane.Normal.y, elli.Plane.Normal.z)));
                if (!elli.CounterClockWise) (elli.StartPoint, elli.EndPoint) = (elli.EndPoint, elli.StartPoint);
                // Plane dxfPlane;
                // if (elli.CounterClockWise) dxfPlane = Import.Plane(Vector3(elli.Center), Vector3(elli.Plane.Normal));
                // else dxfPlane = Import.Plane(Vector3(elli.Center), Vector3(-elli.Plane.Normal));
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
                            Normal = new XYZ(dxfPlane.Normal.x, dxfPlane.Normal.y, dxfPlane.Normal.z),
                            Center = new XYZ(aligned.Center.x, aligned.Center.y, aligned.Center.z),
                            Radius = aligned.Radius,
                            StartAngle = aligned.StartParameter / Math.PI * 180,
                            EndAngle = (aligned.StartParameter + aligned.SweepParameter) / Math.PI * 180,
                        };
                    }
                }
                else
                {
                    ACadSharp.Entities.Ellipse expelli = new ACadSharp.Entities.Ellipse()
                    {
                        Normal = new XYZ(dxfPlane.Normal.x, dxfPlane.Normal.y, dxfPlane.Normal.z),
                        Center = new XYZ(elli.Center.x, elli.Center.y, elli.Center.z),
                        RadiusRatio = elli.MajorRadius / elli.MinorRadius,
                        StartParameter = elli.StartParameter,
                        EndParameter = elli.SweepParameter,
                        EndPoint = new XYZ(elli.EndPoint.x, elli.EndPoint.y, elli.EndPoint.z),
                    };
                    entity = expelli;
                    // TODO: Is any of this necessary? Acadsharp Ellipse constructor is fully set
                    // Plane cdbplane = elli.Plane;
                    // GeoVector2D dir = dxfPlane.Project(cdbplane.DirectionX);
                    // SweepAngle rot = new SweepAngle(GeoVector2D.XAxis, dir);
                    // if (elli.SweepParameter < 0)
                    // {   // there are no clockwise oriented ellipse arcs in dxf
                    //     expelli.Rotation = -rot.Degree;
                    //
                    //     double startParameter = elli.StartParameter + elli.SweepParameter + Math.PI;
                    //     expelli.StartAngle = CalcStartEndAngle(startParameter, expelli.MajorAxis, expelli.MinorAxis);
                    //
                    //     double endParameter = elli.StartParameter + Math.PI;
                    //     expelli.EndAngle = CalcStartEndAngle(endParameter, expelli.MajorAxis, expelli.MinorAxis);
                    // }
                    // else
                    // {
                    //     expelli.Rotation = rot.Degree;
                    //     expelli.StartAngle = CalcStartEndAngle(elli.StartParameter, expelli.MajorAxis, expelli.MinorAxis);
                    //
                    //     double endParameter = elli.StartParameter + elli.SweepParameter;
                    //     expelli.EndAngle = CalcStartEndAngle(endParameter, expelli.MajorAxis, expelli.MinorAxis);
                    // }
                }
            }
            else
            {
                if (elli.IsCircle)
                {
                    entity = new ACadSharp.Entities.Circle()
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
                        RadiusRatio = elli.MajorRadius / elli.MinorRadius,
                        StartParameter = elli.StartParameter,
                        EndParameter = elli.SweepParameter,
                        EndPoint = new XYZ(elli.EndPoint.x, elli.EndPoint.y, elli.EndPoint.z),
                    };
                    
                    entity = expelli;
                    // TODO: Is any of this necessary?
                    // Plane dxfplane = Import.Plane(expelli.Center, expelli.Normal); // this plane is not correct, it has to be rotated
                    // Plane cdbplane = elli.Plane;
                    // GeoVector2D dir = dxfplane.Project(cdbplane.DirectionX);
                    // SweepAngle rot = new SweepAngle(GeoVector2D.XAxis, dir);
                    // expelli.Rotation = rot.Degree;
                }
            }
            return entity;
        }

        private double CalcStartEndAngle(double startEndParameter, double majorAxis, double minorAxis)
        {
            double a = majorAxis * 0.5d;
            double b = minorAxis * 0.5d;
            Vector2 startPoint = new Vector2(a * Math.Cos(startEndParameter), b * Math.Sin(startEndParameter));
            return Vector2.Angle(startPoint) * netDxf.MathHelper.RadToDeg;
        }

        // Looks to be a method to make less repeat code in ExportEllipse but it was never implemented
        // private static void SetEllipseParameters(netDxf.Entities.Ellipse ellipse, double startparam, double endparam)
        // {
        //     //CADability: also set the start and end parameter
        //     //ellipse.StartParameter = startparam;
        //     //ellipse.EndParameter = endparam;
        //     if (MathHelper.IsZero(startparam) && MathHelper.IsEqual(endparam, MathHelper.TwoPI))
        //     {
        //         ellipse.StartAngle = 0.0;
        //         ellipse.EndAngle = 0.0;
        //     }
        //     else
        //     {
        //         double a = ellipse.MajorAxis * 0.5;
        //         double b = ellipse.MinorAxis * 0.5;
        //
        //         Vector2 startPoint = new Vector2(a * Math.Cos(startparam), b * Math.Sin(startparam));
        //         Vector2 endPoint = new Vector2(a * Math.Cos(endparam), b * Math.Sin(endparam));
        //
        //         if (Vector2.Equals(startPoint, endPoint))
        //         {
        //             ellipse.StartAngle = 0.0;
        //             ellipse.EndAngle = 0.0;
        //         }
        //         else
        //         {
        //             ellipse.StartAngle = Vector2.Angle(startPoint) * MathHelper.RadToDeg;
        //             ellipse.EndAngle = Vector2.Angle(endPoint) * MathHelper.RadToDeg;
        //         }
        //     }
        // }

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
                        // Convert ARGB alpha channel to AutoCAD transparency
                        entity.Transparency = new Transparency((byte)((short)((cd.ColorDef.Color.A / 255.0) * 90)));
                        if (cd.ColorDef.Color.ToArgb().Equals(Color.White.ToArgb()) || cd.ColorDef.Color.ToArgb().Equals(Color.Black.ToArgb()))
                        {
                            // White/black index
                            color = new ACadSharp.Color(7);
                        }
                        else
                        {
                            color = new ACadSharp.Color(cd.ColorDef.Color.R, cd.ColorDef.Color.G, cd.ColorDef.Color.B);
                        }
                        break;;
                }
                entity.Color = color;
            }
            if (go.Layer != null)
            {
                if (!createdLayers.TryGetValue(go.Layer, out ACadSharp.Tables.Layer layer))
                {
                    // TODO: Copy more layer data like color?
                    layer = new ACadSharp.Tables.Layer(go.Layer.Name);
                    doc.Layers.Add((Layer)layer);
                    createdLayers[go.Layer] = layer;
                }
                entity.Layer = layer;
            }
            if (go is ILinePattern lp && lp.LinePattern != null)
            {
                if (!createdLinePatterns.TryGetValue(lp.LinePattern, out LineType linetype))
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
        private Vector3 Vector3(GeoPoint p)
        {
            return new Vector3(p.x, p.y, p.z);
        }
        private string GetNextAnonymousBlockName()
        {
            return "AnonymousBlock" + (++anonymousBlockCounter).ToString();
        }
    }
}