using System;
using System.Collections.Generic;
using System.ComponentModel;
using ACadSharp;
using CADability.GeoObject;
using CADability.Shapes;
using CADability.Curve2D;
using CADability.Attribute;
#if WEBASSEMBLY
using CADability.WebDrawing;
using Point = CADability.WebDrawing.Point;
#else
using System.Drawing;
#endif
using System.Text;
using System.IO;
using System.Linq;
using System.Numerics;
using ACadSharp.Entities;
using ACadSharp.Objects;
using ACadSharp.Tables;
using ACadSharp.XData;
using CSMath;
using Color = ACadSharp.Color;
using Path = System.IO.Path;
using Polyline2D = ACadSharp.Entities.Polyline2D;
using Solid = ACadSharp.Entities.Solid;

namespace CADability.DXF
{
    // ODAFileConverter "C:\Zeichnungen\DxfDwg\Stahl" "C:\Zeichnungen\DxfDwg\StahlConverted" "ACAD2010" "DWG" "0" "0"
    // only converts whole directories.
    /// <summary>
    /// Imports a DXF file, converts it to a project
    /// </summary>
    public class Import
    {
        private ACadSharp.CadDocument doc;
        private Project project;
        private Dictionary<string, GeoObject.Block> blockTable;
        private Dictionary<ACadSharp.Tables.Layer, ColorDef> layerColorTable;
        private Dictionary<ACadSharp.Tables.Layer, Attribute.Layer> layerTable;
        /// <summary>
        /// Create the Import instance. The document is being read and converted to netDXF objects.
        /// </summary>
        /// <param name="fileName"></param>
        public Import(string fileName)
        {
            using (Stream stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // MathHelper.Epsilon = 1e-8;
                // doc = DxfDocument.Load(stream);
                doc = ACadSharp.IO.DxfReader.Read(stream);
            }
        }
        public static bool CanImportVersion(string fileName)
        {
            // netDxf.Header.DxfVersion ver = DxfDocument.CheckDxfFileVersion(fileName, out bool isBinary);
            // return ver >= netDxf.Header.DxfVersion.AutoCad2000;
            //TODO: Rewrite version check logic for AcadSharp version support
            return true;
        }
        private void FillModelSpace(Model model)
        {
            // netDxf.Blocks.Block modelSpace = doc.Blocks["*Model_Space"];
            // foreach (EntityObject item in modelSpace.Entities)
            // {
            //     IGeoObject geoObject = GeoObjectFromEntity(item);
            //     if (geoObject != null) model.Add(geoObject);
            // }
            // model.Name = "*Model_Space";

            var entities = doc.ModelSpace.Entities;
            foreach (Entity entity in entities)
            {
                IGeoObject geoObject = GeoObjectFromEntity(entity);
                if (geoObject != null) model.Add(geoObject);
            }
            model.Name = "*Model_Space";
        }
        private void FillPaperSpace(Model model)
        {
            // netDxf.Blocks.Block modelSpace = doc.Blocks["*Paper_Space"];
            // foreach (EntityObject item in modelSpace.Entities)
            // {
            //     IGeoObject geoObject = GeoObjectFromEntity(item);
            //     if (geoObject != null) model.Add(geoObject);
            // }
            
            var entities = doc.PaperSpace.Entities;
            foreach (Entity entity in entities)
            {
                IGeoObject geoObject = GeoObjectFromEntity(entity);
                if (geoObject != null) model.Add(geoObject);
            }
            model.Name = "*Paper_Space";
        }
        /// <summary>
        /// creates and returns the project
        /// </summary>
        public Project Project { get => CreateProject(); }
        private Project CreateProject()
        {
            if (doc == null) return null;
            project = Project.CreateSimpleProject();
            blockTable = new Dictionary<string, GeoObject.Block>();
            layerColorTable = new Dictionary<ACadSharp.Tables.Layer, ColorDef>();
            layerTable = new Dictionary<ACadSharp.Tables.Layer, Attribute.Layer>();
            foreach (var item in doc.Layers)
            {
                Attribute.Layer layer = project.LayerList.CreateOrFind(item.Name);
                layerTable[item] = layer;
                Color rgb = new Color(item.Color.R, item.Color.G, item.Color.B);
                Color white = new Color(255, 255, 255);
                Color black = new Color(0, 0, 0);
                if (rgb.Equals(white)) rgb = black;
                System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(255, rgb.R, rgb.G, rgb.B);
                ColorDef cd = project.ColorList.CreateOrFind(item.Name + ":ByLayer", sysColor);
                layerColorTable[item] = cd;
            }
            foreach (var item in doc.LineTypes)
            {
                var segments = item.Segments.ToList();
                List<double> pattern = new List<double>();
                for (int i = 0; i < segments.Count; i++)
                {
                    if (segments[i].Shapeflag == LinetypeShapeFlags.None)
                    {
                        pattern.Add(Math.Abs(segments[i].Length));
                    }
                }
                project.LinePatternList.CreateOrFind(item.Name, pattern.ToArray());
            }
            FillModelSpace(project.GetModel(0));
            Model paperSpace = new Model();
            FillPaperSpace(paperSpace);
            if (paperSpace.Count > 0)
            {
                project.AddModel(paperSpace);
                Model modelSpace = project.GetModel(0);
                if (modelSpace.Count == 0)
                {   // if the modelSpace is empty and the paperSpace contains entities, then show the paperSpace
                    for (int i = 0; i < project.ModelViewCount; ++i)
                    {
                        ProjectedModel pm = project.GetProjectedModel(i);
                        if (pm.Model == modelSpace) pm.Model = paperSpace;
                    }
                }
            }
            doc = null;
            return project;
        }
        private IGeoObject GeoObjectFromEntity(ACadSharp.Entities.Entity item)
        {
            IGeoObject res = null;
            switch (item)
            {
                case ACadSharp.Entities.Line dxfLine: res = CreateLine(dxfLine); break;
                case ACadSharp.Entities.Ray dxfRay: res = CreateRay(dxfRay); break;
                case ACadSharp.Entities.Arc dxfArc: res = CreateArc(dxfArc); break;
                case ACadSharp.Entities.Circle dxfCircle: res = CreateCircle(dxfCircle); break;
                case ACadSharp.Entities.Ellipse dxfEllipse: res = CreateEllipse(dxfEllipse); break;
                case ACadSharp.Entities.Spline dxfSpline: res = CreateSpline(dxfSpline); break;
                case ACadSharp.Entities.Face3D dxfFace: res = CreateFace(dxfFace); break;
                case ACadSharp.Entities.PolyfaceMesh dxfPolyfaceMesh: res = CreatePolyfaceMesh(dxfPolyfaceMesh); break;
                case ACadSharp.Entities.Hatch dxfHatch: res = CreateHatch(dxfHatch); break;
                case ACadSharp.Entities.Solid dxfSolid: res = CreateSolid(dxfSolid); break;
                case ACadSharp.Entities.Insert dxfInsert: res = CreateInsert(dxfInsert); break;
                case ACadSharp.Entities.MLine dxfMLine: res = CreateMLine(dxfMLine); break;
                case ACadSharp.Entities.TextEntity dxfText: res = CreateText(dxfText); break;
                case ACadSharp.Entities.Dimension dxfDimension: res = CreateDimension(dxfDimension); break;
                case ACadSharp.Entities.MText dxfMText: res = CreateMText(dxfMText); break;
                case ACadSharp.Entities.Leader dxfLeader: res = CreateLeader(dxfLeader); break;
                case ACadSharp.Entities.Point dxfPoint: res = CreatePoint(dxfPoint); break;
                case ACadSharp.Entities.Mesh dxfMesh: res = CreateMesh(dxfMesh); break;
                case ACadSharp.Entities.IPolyline polyline: res = CreateExplodedPolyline(polyline); break; // Does this import LwPolyline?
                //case ACadSharp.Entities.LwPolyline lwPolyline: res = CreateExplodedPolyline(lwPolyline); break; // Is this needed?
                // Polyline import as exploded geometry streamlined with CreateExplodedPolyline
                // TODO: Revisit polyline importing for enhancements or potential polyline-to-polyline imports
                //case ACadSharp.Entities.Polyline2D dxfPolyline2D: res = CreatePolyline2D(dxfPolyline2D); break;
                //case ACadSharp.Entities.Polyline3D dxfPolyline3D: res = CreatePolyline3D(dxfPolyline3D); break;
                        
                default:
                    System.Diagnostics.Trace.WriteLine("dxf: not imported: " + item.ToString());
                    break;
            }
            if (res != null)
            {
                SetAttributes(res, item);
                SetUserData(res, item);
                res.IsVisible = !item.IsInvisible;
            }
            return res;
        }
        private static GeoPoint GeoPoint(Vector3 p)
        {
            return new GeoPoint(p.X, p.Y, p.Z);
        }
        private static GeoVector GeoVector(Vector3 p)
        {
            return new GeoVector(p.X, p.Y, p.Z);
        }
        internal static Plane Plane(Vector3 center, Vector3 normal)
        {
            // this is AutoCADs arbitrary axis algorithm we must use here to get the correct plane
            // because sometimes we need the correct x-axis, y-axis orientation
            //Let the world Y axis be called Wy, which is always(0, 1, 0).
            //Let the world Z axis be called Wz, which is always(0, 0, 1).
            //If(abs(Nx) < 1 / 64) and(abs(Ny) < 1 / 64) then
            //     Ax = Wy X N(where “X” is the cross - product operator).
            //Otherwise,
            //     Ax = Wz X N.
            //Scale Ax to unit length.

            GeoVector n = GeoVector(normal);
            GeoVector ax = (Math.Abs(normal.X) < 1.0 / 64 && Math.Abs(normal.Y) < 1.0 / 64) ? CADability.GeoVector.YAxis ^ n : CADability.GeoVector.ZAxis ^ n;
            GeoVector ay = n ^ ax;
            return new Plane(GeoPoint(center), ax, ay);
        }
        private HatchStyleSolid FindOrCreateSolidHatchStyle(Color clr)
        {
            for (int i = 0; i < project.HatchStyleList.Count; i++)
            {
                if (project.HatchStyleList[i] is HatchStyleSolid hss)
                {
                    Color hssColor = new Color(hss.Color.Color.R, hss.Color.Color.G, hss.Color.Color.B);
                    if (hssColor.Equals(clr)) return hss;
                }
            }
            HatchStyleSolid nhss = new HatchStyleSolid();
            nhss.Name = "Solid_" + clr.ToString();
            var sysClr = System.Drawing.Color.FromArgb(255, clr.R, clr.G, clr.B);
            nhss.Color = project.ColorList.CreateOrFind(clr.ToString(), sysClr);
            project.HatchStyleList.Add(nhss);
            return nhss;
        }
        private HatchStyleLines FindOrCreateHatchStyleLines(ACadSharp.Entities.Entity entity, double lineAngle, double lineDistance, double[] dashes)
        {
            var entColor = System.Drawing.Color.FromArgb(255, entity.Color.R, entity.Color.G, entity.Color.B);
            for (int i = 0; i < project.HatchStyleList.Count; i++)
            {
                if (project.HatchStyleList[i] is HatchStyleLines hsl)
                {
                    // TODO: Create PR to produce portable RGB values from ColorDef
                    // Ignore alpha channel. Not supported by ACadSharp
                    entColor = System.Drawing.Color.FromArgb(255, entColor.R, entColor.G, entColor.B);
                    if (hsl.ColorDef.Color.ToArgb() == entColor.ToArgb() && (Math.Abs(hsl.LineAngle) - Math.Abs(lineAngle) < Precision.eps) && Math.Abs(hsl.LineDistance) - Math.Abs(lineDistance) < Precision.eps) return hsl;
                }
            }
            HatchStyleLines nhsl = new HatchStyleLines();
            string name = NewName(entity.Layer.Name, project.HatchStyleList);
            nhsl.Name = name;
            nhsl.LineAngle = lineAngle;
            nhsl.LineDistance = lineDistance;
            nhsl.ColorDef = project.ColorList.CreateOrFind(entity.Color.ToString(), entColor);
            LineweightType lw = entity.LineWeight;
            if (lw == LineweightType.ByLayer) lw = entity.Layer.LineWeight;
            // TODO: Test this with a block to make sure AcadSharp blocks are behaving correctly
            if (lw == LineweightType.ByBlock && entity.Owner is BlockRecord blockRecord)
            {
                foreach (Entity drawingEntity in blockRecord.Document.Entities)
                {
                    if (drawingEntity is Insert insert && insert.Block == blockRecord)
                    {
                        lw = insert.Layer.LineWeight;
                        break;
                    }
                }
            }
            if (lw < 0) lw = 0;
            nhsl.LineWidth = project.LineWidthList.CreateOrFind("DXF_" + lw.ToString(), ((int)lw) / 100.0);
            nhsl.LinePattern = FindOrcreateLinePattern(dashes);
            project.HatchStyleList.Add(nhsl);
            return nhsl;
        }
        private ColorDef FindOrCreateColor(ACadSharp.Color color, ACadSharp.Tables.Layer layer)
        {
            if (color.IsByLayer && layer != null)
            {
                ColorDef res = layerColorTable[layer] as ColorDef;
                if (res != null) return res;
            }
            Color rgb = new Color(color.R, color.G, color.B);
            Color white = new Color(255, 255, 255);
            Color black = new Color(0, 0, 0);
            if (rgb.Equals(white)) rgb = black;
            string colorname = rgb.ToString();
            var sysClr = System.Drawing.Color.FromArgb(255, color.R, color.G, color.B);
            return project.ColorList.CreateOrFind(colorname, sysClr);
        }
        private string NewName(string prefix, IAttributeList list)
        {
            string name = prefix;
            while (list.Find(name) != null)
            {
                string[] parts = name.Split('_');
                if (parts.Length >= 2 && int.TryParse(parts[parts.Length - 1], out int nn))
                {
                    parts[parts.Length - 1] = (nn + 1).ToString();
                    name = parts[0];
                    for (int j = 1; j < parts.Length; j++) name += parts[j];
                }
                else
                {
                    name += "_1";
                }
            }
            return name;
        }
        private LinePattern FindOrcreateLinePattern(double[] dashes, string name = null)
        {
            // in CADability a line pattern always starts with a stroke (dash) followed by a gap (space). In DXF positiv is stroke, negative is gap
            if (dashes.Length == 0)
            {
                for (int i = 0; i < project.LinePatternList.Count; i++)
                {
                    if (project.LinePatternList[i].Pattern == null || project.LinePatternList[i].Pattern.Length == 0) return project.LinePatternList[i];
                }
                return new LinePattern(NewName("DXFpattern", project.LinePatternList));
            }
            if (dashes[0] < 0)
            {
                List<double> pattern = new List<double>(dashes);
                if (pattern[pattern.Count - 1] > 0)
                {
                    pattern.Insert(0, pattern[pattern.Count - 1]);
                    pattern.RemoveAt(pattern.Count - 1);
                }
                else
                {   // a pattern that starts with a gap and ends with a gap, what does this mean?
                    pattern.Insert(0, 0.0);
                }
                if ((pattern.Count & 0x01) != 0) pattern.Add(0.0); // there must be an even number (stroke-gap appear in pairs)
                dashes = pattern.ToArray();
            }
            else if ((dashes.Length & 0x01) != 0)
            {
                List<double> pattern = new List<double>(dashes);
                pattern.Add(0.0);
                dashes = pattern.ToArray();
            }
            return new LinePattern(NewName("DXFpattern", project.LinePatternList), dashes);
        }
        private void SetAttributes(IGeoObject go, ACadSharp.Entities.Entity entity)
        {
            if (go is IColorDef cd) cd.ColorDef = FindOrCreateColor(entity.Color, entity.Layer);
            go.Layer = layerTable[entity.Layer];
            if (go is ILinePattern lp) lp.LinePattern = project.LinePatternList.Find(entity.LineType.Name);
            if (go is ILineWidth ld)
            {
                LineweightType lw = entity.LineWeight;
                if (lw == LineweightType.ByLayer) lw = entity.Layer.LineWeight;
                // Blockrecords are instantiated in the drawing with inserts, which hold the style data
                // Insert must be found to retrieve ByBlock style data
                // There may be a more efficient way to do this than walking all entities
                if (lw == LineweightType.ByBlock && entity.Owner is BlockRecord blockRecord)
                {
                    foreach (Entity drawingEntity in blockRecord.Document.Entities)
                    {
                        if (drawingEntity is Insert insert && insert.Block == blockRecord)
                        {
                            lw = insert.Layer.LineWeight;
                            break;
                        }
                    }
                }
                if (lw < 0) lw = 0;
                ld.LineWidth = project.LineWidthList.CreateOrFind("DXF_" + lw.ToString(), ((int)lw) / 100.0);
            }
        }
        private void SetUserData(IGeoObject go, ACadSharp.Entities.Entity entity)
        {
            foreach (KeyValuePair<AppId, ExtendedData> item in entity.ExtendedData.Entries)
            {
                ExtendedEntityData xdata = new ExtendedEntityData();
                xdata.ApplicationName = item.Value.AppId.Name;

                string name = item.Value.AppId.Name + ":" + item.Key;

                foreach (ExtendedDataRecord record in item.Value.Records)
                {
                    xdata.Data.Add(new KeyValuePair<DxfCode, object>(record.Code, record.Value));
                }

                go.UserData.Add(name, xdata);
            }
            go.UserData["DxfImport.Handle"] = new UserInterface.StringProperty(entity.Handle.ToString(), "DxfImport.Handle");
        }
        private GeoObject.Block FindBlock(ACadSharp.Blocks.Block entity)
        {
            if (!blockTable.TryGetValue(entity.Handle.ToString(), out GeoObject.Block found))
            {
                found = GeoObject.Block.Construct();
                found.Name = entity.Name;
                found.RefPoint = new GeoPoint(entity.BasePoint.X, entity.BasePoint.Y, entity.BasePoint.Z);
                List<CadObject> entities = entity.Reactors.Values.ToList();
                foreach (CadObject obj in entities)
                {
                    if (obj is ACadSharp.Entities.Entity e)
                    {
                        IGeoObject go = GeoObjectFromEntity(e);
                        if (go != null) found.Add(go);
                    }
                }
                blockTable[entity.Handle.ToString()] = found;
            }
            return found;
        }
        private IGeoObject CreateLine(ACadSharp.Entities.Line line)
        {
            GeoObject.Line l = GeoObject.Line.Construct();
            {
                XYZ sp = line.StartPoint;
                XYZ ep = line.EndPoint;
                l.StartPoint = new GeoPoint(sp.X, sp.Y, sp.Z);
                l.EndPoint = new GeoPoint(ep.X, ep.Y, ep.Z);
                double th = line.Thickness;
                GeoVector no = new GeoVector(line.Normal.X, line.Normal.Y, line.Normal.Z);
                if (th != 0.0 && !no.IsNullVector())
                {
                    if (l.Length < Precision.eps)
                    {
                        l.EndPoint += th * no;
                        return l;
                    }
                    else
                    {
                        return Make3D.Extrude(l, th * no, null);
                    }
                }
                return l;
            }
        }
        private IGeoObject CreateRay(ACadSharp.Entities.Ray ray)
        {
            GeoObject.Line l = GeoObject.Line.Construct();
            XYZ sp = ray.StartPoint;
            XYZ dir = ray.Direction;
            l.StartPoint = new GeoPoint(sp.X, sp.Y, sp.Z);
            l.EndPoint = l.StartPoint + new GeoVector(dir.X, dir.Y, dir.Z);
            return l;
        }
        private IGeoObject CreateArc(ACadSharp.Entities.Arc arc)
        {
            GeoObject.Ellipse e = GeoObject.Ellipse.Construct();
            GeoVector nor = new GeoVector(arc.Normal.X, arc.Normal.Y, arc.Normal.Z);
            GeoPoint cnt = new GeoPoint(arc.Center.X, arc.Center.Y, arc.Center.Z);
            Plane plane = new Plane(cnt, nor);
            double start = arc.StartAngle;
            double end = arc.EndAngle;
            double sweep = end - start;
            if (sweep < 0.0) sweep += Math.PI * 2.0;
            //if (sweep < Precision.epsa) sweep = Math.PI * 2.0;
            if (start == end) sweep = 0.0;
            if (start == Math.PI * 2.0 && end == 0.0) sweep = 0.0; // see in modena.dxf
            // Arcs are always counterclockwise, but maybe the normal is (0,0,-1) in 2D drawings.
            e.SetArcPlaneCenterRadiusAngles(plane, cnt, arc.Radius, start, sweep);

            //If an arc is a full circle don't import as ellipse as this will be discarded later by Ellipse.HasValidData() 
            if (e.IsCircle && sweep == 0.0d && Precision.IsEqual(e.StartPoint, e.EndPoint))
            {
                GeoObject.Ellipse circle = GeoObject.Ellipse.Construct();
                circle.SetCirclePlaneCenterRadius(plane, cnt, arc.Radius);
                e = circle;
            }

            double th = arc.Thickness;
            if (th != 0.0 && !nor.IsNullVector())
            {
                return Make3D.Extrude(e, th * nor, null);
            }
            return e;
        }

        private IGeoObject CreateCircle(ACadSharp.Entities.Circle circle)
        {
            GeoObject.Ellipse e = GeoObject.Ellipse.Construct();
            GeoPoint cnt = new GeoPoint(circle.Center.X, circle.Center.Y, circle.Center.Z);
            GeoVector nor = new GeoVector(circle.Normal.X, circle.Normal.Y, circle.Normal.Z);
            Plane plane = new Plane(cnt, nor);
            e.SetCirclePlaneCenterRadius(plane, cnt, circle.Radius);
            double th = circle.Thickness;
            if (th != 0.0 && !nor.IsNullVector())
            {
                return Make3D.Extrude(e, th * nor, null);
            }
            return e;
        }
        private IGeoObject CreateEllipse(ACadSharp.Entities.Ellipse ellipse)
        {
            GeoObject.Ellipse e = GeoObject.Ellipse.Construct();
            GeoPoint cnt = new GeoPoint(ellipse.Center.X, ellipse.Center.Y, ellipse.Center.Z);
            GeoVector nor = new GeoVector(ellipse.Normal.X, ellipse.Normal.Y, ellipse.Normal.Z);
            Plane plane = new Plane(cnt, nor);
            ModOp2D rot = ModOp2D.Rotate(ellipse.Rotation);
            GeoVector2D majorAxis = 0.5 * ellipse.MajorAxis * (rot * GeoVector2D.XAxis);
            GeoVector2D minorAxis = 0.5 * ellipse.MinorAxis * (rot * GeoVector2D.YAxis);
            e.SetEllipseCenterAxis(cnt, plane.ToGlobal(majorAxis), plane.ToGlobal(minorAxis));

            XY startPoint = ellipse.PolarCoordinateRelativeToCenter(ellipse.StartParameter);
            double sp = CalcStartEndParameter(startPoint, ellipse.MajorAxis, ellipse.MinorAxis);

            XY endPoint = ellipse.PolarCoordinateRelativeToCenter(ellipse.EndParameter);
            double ep = CalcStartEndParameter(endPoint, ellipse.MajorAxis, ellipse.MinorAxis);

            e.StartParameter = sp;
            e.SweepParameter = ep - sp;
            if (e.SweepParameter == 0.0) e.SweepParameter = Math.PI * 2.0;
            if (e.SweepParameter < 0.0) e.SweepParameter += Math.PI * 2.0; // seems it is always counterclockwise
            // it looks like clockwise 2d ellipses are defined with normal vector (0, 0, -1)
            return e;
        }

        private double CalcStartEndParameter(XY startEndPoint, double majorAxis, double minorAxis)
        {
            double a = 1 / (0.5 * majorAxis);
            double b = 1 / (0.5 * minorAxis);
            double parameter = Math.Atan2(startEndPoint.Y * b, startEndPoint.X * a);
            return parameter;
        }

        private IGeoObject CreateSpline(ACadSharp.Entities.Spline spline)
        {
            int degree = spline.Degree;
            bool isClosed = spline.Flags.HasFlag(SplineFlags.Closed);
            bool isPeriodic = spline.Flags.HasFlag(SplineFlags.Periodic);
            if (spline.ControlPoints.Count == 0 && spline.FitPoints.Count > 0)
            {
                BSpline bsp = BSpline.Construct();
                GeoPoint[] fp = new GeoPoint[spline.FitPoints.Count];
                for (int i = 0; i < fp.Length; i++)
                {
                    fp[i] = new GeoPoint(spline.FitPoints[i].X, spline.FitPoints[i].Y, spline.FitPoints[i].Z);
                }
                
                bsp.ThroughPoints(fp, spline.Degree, isClosed);
                return bsp;
            }
            else
            {
                bool forcePolyline2D = false;
                GeoPoint[] poles = new GeoPoint[spline.ControlPoints.Count];
                double[] weights = new double[spline.ControlPoints.Count];
                for (int i = 0; i < poles.Length; i++)
                {
                    poles[i] = new GeoPoint(spline.ControlPoints[i].X, spline.ControlPoints[i].Y, spline.ControlPoints[i].Z);
                    weights[i] = spline.Weights[i];

                    if (i > 0 && (poles[i] | poles[i - 1]) < Precision.eps)
                    {
                        forcePolyline2D = true;
                    }
                }
                double[] kn = new double[spline.Knots.Count];
                for (int i = 0; i < kn.Length; ++i)
                {
                    kn[i] = spline.Knots[i];
                }
                if (poles.Length == 2 && degree > 1)
                {   // damit geht kein vernünftiger Spline, höchstens mit degree=1
                    GeoObject.Line l = GeoObject.Line.Construct();
                    l.StartPoint = poles[0];
                    l.EndPoint = poles[1];
                    return l;
                }
                BSpline bsp = BSpline.Construct();
                //TODO: Can Periodic spline be not closed?
                if (bsp.SetData(degree, poles, weights, kn, null, isClosed && isPeriodic))
                {
                    // BSplines with inner knots of multiplicity degree+1 make problems, because the spline have no derivative at these points
                    // so we split these splines
                    List<int> splitKnots = new List<int>();
                    for (int i = degree + 1; i < kn.Length - degree - 1; i++)
                    {
                        if (kn[i] == kn[i - 1])
                        {
                            bool sameKnot = true;
                            for (int j = 0; j < degree; j++)
                            {
                                if (kn[i - 1] != kn[i + j]) sameKnot = false;
                            }
                            if (sameKnot) splitKnots.Add(i - 1);
                        }
                    }
                    if (splitKnots.Count > 0)
                    {
                        List<ICurve> parts = new List<ICurve>();
                        BSpline part = bsp.TrimParam(kn[0], kn[splitKnots[0]]);
                        if (CADability.GeoPoint.Distance(part.Poles) > Precision.eps && (part as ICurve).Length > Precision.eps) parts.Add(part);
                        for (int i = 1; i < splitKnots.Count; i++)
                        {
                            part = bsp.TrimParam(kn[splitKnots[i - 1]], kn[splitKnots[i]]);
                            if (CADability.GeoPoint.Distance(part.Poles) > Precision.eps && (part as ICurve).Length > Precision.eps) parts.Add(part);
                        }
                        part = bsp.TrimParam(kn[splitKnots[splitKnots.Count - 1]], kn[kn.Length - 1]);

                        if (CADability.GeoPoint.Distance(part.Poles) > Precision.eps && (part as ICurve).Length > Precision.eps) parts.Add(part);
                        GeoObject.Path path = GeoObject.Path.Construct();
                        path.Set(parts.ToArray());
                        return path;
                    }
                    // if (spline.IsPeriodic) bsp.IsClosed = true; // to remove strange behavior in hünfeld.dxf

                    if (forcePolyline2D)
                    {
                        //Look at https://github.com/SOFAgh/CADability/issues/173 to see why this is done.

                        ICurve curve = (ICurve)bsp;
                        //Use approximate to get the count of lines that will be needed to convert the spline into a Polyline2D
                        double maxError = Settings.GlobalSettings.GetDoubleValue("Approximate.Precision", 0.01);
                        ICurve approxCurve = curve.Approximate(true, maxError);

                        int usedCurves = 0;
                        if (approxCurve is GeoObject.Line)
                            usedCurves = 2;
                        else
                            usedCurves = approxCurve.SubCurves.Length;
                        
                        // TODO: Find a way to implement this in ACadSharp
                        netDxf.Entities.Polyline2D p2d = spline.ToPolyline2D(usedCurves);
                        var res = CreatePolyline2D(p2d);
                        
                        return res;
                    }

                    return bsp;
                }
                // strange spline in "bspline-closed-periodic.dxf"
            }
            return null;
        }

        private IGeoObject CreateFace(ACadSharp.Entities.Face3D face)
        {
            List<GeoPoint> points = new List<GeoPoint>();
            GeoPoint p = new GeoPoint(face.FirstCorner.X, face.FirstCorner.Y, face.FirstCorner.Z);
            points.Add(p);
            p = new GeoPoint(face.SecondCorner.X, face.SecondCorner.Y, face.SecondCorner.Z);
            if (points[points.Count - 1] != p) points.Add(p);
            p = new GeoPoint(face.ThirdCorner.X, face.ThirdCorner.Y, face.ThirdCorner.Z);
            if (points[points.Count - 1] != p) points.Add(p);
            p = new GeoPoint(face.FourthCorner.X, face.FourthCorner.Y, face.FourthCorner.Z);
            if (points[points.Count - 1] != p) points.Add(p);
            if (points.Count == 3)
            {
                Plane pln = new Plane(points[0], points[1], points[2]);
                PlaneSurface surf = new PlaneSurface(pln);
                Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(points[1]), pln.Project(points[2]) });
                SimpleShape ss = new SimpleShape(bdr);
                Face fc = Face.MakeFace(surf, ss);
                return fc;
            }
            else if (points.Count == 4)
            {
                Plane pln = CADability.Plane.FromPoints(points.ToArray(), out double maxDist, out bool isLinear);
                if (!isLinear)
                {
                    if (maxDist > Precision.eps)
                    {
                        Face fc1 = Face.MakeFace(points[0], points[1], points[2]);
                        Face fc2 = Face.MakeFace(points[0], points[2], points[3]);
                        GeoObject.Block blk = GeoObject.Block.Construct();
                        blk.Set(new GeoObjectList(fc1, fc2));
                        return blk;
                    }
                    else
                    {
                        PlaneSurface surf = new PlaneSurface(pln);
                        Border bdr = new Border(new GeoPoint2D[] { pln.Project(points[0]), pln.Project(points[1]), pln.Project(points[2]), pln.Project(points[3]) });
                        double[] sis = bdr.GetSelfIntersection(Precision.eps);
                        if (sis.Length > 0)
                        {
                            // multiple of three values: parameter1, parameter2, crossproduct of intersection direction
                            // there can only be one intersection
                            Border[] splitted = bdr.Split(new double[] { sis[0], sis[1] });
                            for (int i = 0; i < splitted.Length; i++)
                            {
                                if (splitted[i].IsClosed) bdr = splitted[i];
                            }
                        }
                        SimpleShape ss = new SimpleShape(bdr);
                        Face fc = Face.MakeFace(surf, ss);
                        return fc;
                    }
                }
            }
            return null;

        }
        private IGeoObject CreatePolyfaceMesh(ACadSharp.Entities.PolyfaceMesh polyfacemesh)
        {
            Entity[] exploded = polyfacemesh.Explode().ToArray();

            GeoPoint[] vertices = new GeoPoint[exploded.Length];
            for (int i = 0; i < vertices.Length; i++)
            {
                ACadSharp.Entities.Vertex vert = (ACadSharp.Entities.Vertex)exploded[i];
                vertices[i] = new GeoPoint(vert.Location.X, vert.Location.Y, vert.Location.Z); // there is more information, I would need a good example
            }

            List<Face> faces = new List<Face>();
            for (int i = 0; i < polyfacemesh.Faces.Count; i++)
            {
                
                short[] indices = {polyfacemesh.Faces[i].Index1, polyfacemesh.Faces[i].Index2, polyfacemesh.Faces[i].Index3, polyfacemesh.Faces[i].Index4};
                for (int j = 0; j < indices.Length; j++)
                {
                    indices[j] = (short)(Math.Abs(indices[j]) - 1); // why? what does it mean?
                }
                if (indices.Length <= 3 || indices[3] == indices[2])
                {
                    if (indices[0] != indices[1] && indices[1] != indices[2])
                    {
                        Plane pln = new Plane(vertices[indices[0]], vertices[indices[1]], vertices[indices[2]]);
                        PlaneSurface surf = new PlaneSurface(pln);
                        Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[1]]), pln.Project(vertices[indices[2]]) });
                        SimpleShape ss = new SimpleShape(bdr);
                        Face fc = Face.MakeFace(surf, ss);
                        faces.Add(fc);
                    }
                }
                else
                {
                    if (indices[0] != indices[1] && indices[1] != indices[2])
                    {
                        Plane pln = new Plane(vertices[indices[0]], vertices[indices[1]], vertices[indices[2]]);
                        PlaneSurface surf = new PlaneSurface(pln);
                        Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[1]]), pln.Project(vertices[indices[2]]) });
                        SimpleShape ss = new SimpleShape(bdr);
                        Face fc = Face.MakeFace(surf, ss);
                        faces.Add(fc);
                    }
                    if (indices[2] != indices[3] && indices[3] != indices[0])
                    {
                        Plane pln = new Plane(vertices[indices[2]], vertices[indices[3]], vertices[indices[0]]);
                        PlaneSurface surf = new PlaneSurface(pln);
                        Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[3]]), pln.Project(vertices[indices[0]]) });
                        SimpleShape ss = new SimpleShape(bdr);
                        Face fc = Face.MakeFace(surf, ss);
                        faces.Add(fc);
                    }
                }
            }
            if (faces.Count > 1)
            {
                GeoObjectList sewed = Make3D.SewFacesAndShells(new GeoObjectList(faces.ToArray() as IGeoObject[]));
                return sewed[0];
            }
            else if (faces.Count == 1)
            {
                return faces[0];
            }
            else return null;
        }
        private IGeoObject CreateHatch(ACadSharp.Entities.Hatch hatch)
        {
            CompoundShape cs = null;
            bool ok = true;
            List<ICurve2D> allCurves = new List<ICurve2D>();
            Plane pln = CADability.Plane.XYPlane;
            for (int i = 0; i < hatch.Paths.Count; i++)
            {

                // System.Diagnostics.Trace.WriteLine("Loop: " + i.ToString());
                //OdDbHatch.HatchLoopType.kExternal
                // hatch.BoundaryPaths[i].PathType
                List<ICurve> boundaryEntities = new List<ICurve>();
                for (int j = 0; j < hatch.Paths[i].Edges.Count; j++)
                {
                    IGeoObject ent = GeoObjectFromEntity(hatch.Paths[i].Entities[j]);
                    if (ent is ICurve crv) boundaryEntities.Add(crv);
                }
                //for (int j = 0; j < hatch.BoundaryPaths[i].Entities.Count; j++)
                //{
                //    IGeoObject ent = GeoObjectFromEntity(hatch.BoundaryPaths[i].Entities[j]);
                //    if (ent is ICurve crv) boundaryEntities.Add(crv);
                //}
                if (i == 0)
                {
                    if (!Curves.GetCommonPlane(boundaryEntities, out pln)) return null; // there must be a common plane
                }
                ICurve2D[] bdr = new ICurve2D[boundaryEntities.Count];
                for (int j = 0; j < bdr.Length; j++)
                {
                    bdr[j] = boundaryEntities[j].GetProjectedCurve(pln);
                }
                try
                {
                    Border border = Border.FromUnorientedList(bdr, true);
                    allCurves.AddRange(bdr);
                    if (border != null)
                    {
                        SimpleShape ss = new SimpleShape(border);
                        if (cs == null)
                        {
                            cs = new CompoundShape(ss);
                        }
                        else
                        {
                            CompoundShape cs1 = new CompoundShape(ss);
                            double a = cs.Area;
                            cs = cs - new CompoundShape(ss); // assuming the first border is the outer bound followed by holes
                            if (cs.Area >= a) ok = false; // don't know how to descriminate between outer bounds and holes
                        }
                    }
                }
                catch (BorderException)
                {
                }
            }
            if (cs != null)
            {
                if (cs.Area == 0.0 || !ok)
                {   // try to make something usefull from the curves
                    cs = CompoundShape.CreateFromList(allCurves.ToArray(), Precision.eps);
                    if (cs == null || cs.Area == 0.0) return null;
                }
                GeoObject.Hatch res = GeoObject.Hatch.Construct();
                res.CompoundShape = cs;
                res.Plane = pln;
                if (hatch.PatternType == HatchPatternType.SolidFill)
                {
                    HatchStyleSolid hst = FindOrCreateSolidHatchStyle(hatch.Color);
                    res.HatchStyle = hst;
                    return res;
                }
                else
                {
                    GeoObjectList list = new GeoObjectList();
                    for (int i = 0; i < hatch.Pattern.Lines.Count; i++)
                    {
                        if (i > 0) res = res.Clone() as GeoObject.Hatch;
                        double lineAngle = Angle.Deg(hatch.Pattern.Lines[i].Angle);
                        double baseX = hatch.Pattern.Lines[i].BasePoint.X;
                        double baseY = hatch.Pattern.Lines[i].BasePoint.Y;
                        double offsetX = hatch.Pattern.Lines[i].BasePoint.X;
                        double offsetY = hatch.Pattern.Lines[i].BasePoint.Y;
                        double[] dashes = hatch.Pattern.Lines[i].DashLengths.ToArray();
                        HatchStyleLines hsl = FindOrCreateHatchStyleLines(hatch, lineAngle, Math.Sqrt(offsetX * offsetX + offsetY * offsetY), dashes);
                        res.HatchStyle = hsl;
                        list.Add(res);
                    }
                    if (list.Count > 1)
                    {
                        GeoObject.Block block = GeoObject.Block.Construct();
                        block.Set(new GeoObjectList(list));
                        return block;
                    }
                    else return res;
                }
            }
            else
            {
                return null;
            }
        }
        private IGeoObject CreateSolid(ACadSharp.Entities.Solid solid)
        {
            // TODO: Look into a return type other than hatch based on DXF spec
            GeoPoint firstPoint = new GeoPoint(solid.FirstCorner.X, solid.FirstCorner.Y, solid.FirstCorner.Z);
            GeoPoint secondPoint = new GeoPoint(solid.SecondCorner.X, solid.SecondCorner.Y, solid.SecondCorner.Z);
            GeoPoint thirdPoint = new GeoPoint(solid.ThirdCorner.X, solid.ThirdCorner.Y, solid.ThirdCorner.Z);
            GeoPoint fourthPoint = new GeoPoint(solid.FourthCorner.X, solid.FourthCorner.Y, solid.FourthCorner.Z);
            Plane ocs = CADability.Plane.FromPoints(new GeoPoint[] { firstPoint, secondPoint, thirdPoint, fourthPoint }, out _, out _);
            
            // not sure, whether the ocs is correct, maybe the position is (0,0,solid.Elevation)

            HatchStyleSolid hst = FindOrCreateSolidHatchStyle(solid.Color);
            List<GeoPoint> points = new List<GeoPoint>();
            points.Add(ocs.ToGlobal(firstPoint.To2D()));
            points.Add(ocs.ToGlobal(secondPoint.To2D()));
            points.Add(ocs.ToGlobal(thirdPoint.To2D()));
            points.Add(ocs.ToGlobal(fourthPoint.To2D()));
            for (int i = 3; i > 0; --i)
            {   // gleiche Punkte wegmachen
                for (int j = 0; j < i; ++j)
                {
                    if (Precision.IsEqual(points[j], points[i]))
                    {
                        points.RemoveAt(i);
                        break;
                    }
                }
            }
            if (points.Count < 3) return null;

            Plane pln;
            try
            {
                pln = new Plane(points[0], points[1], points[2]);
            }
            catch (PlaneException)
            {
                return null;
            }
            GeoPoint2D[] vertex = new GeoPoint2D[points.Count + 1];
            for (int i = 0; i < points.Count; ++i) vertex[i] = pln.Project(points[i]);
            vertex[points.Count] = vertex[0];
            Curve2D.Polyline2D poly2d = new Curve2D.Polyline2D(vertex);
            Border bdr = new Border(poly2d);
            CompoundShape cs = new CompoundShape(new SimpleShape(bdr));
            GeoObject.Hatch hatch = GeoObject.Hatch.Construct();
            hatch.CompoundShape = cs;
            hatch.HatchStyle = hst;
            hatch.Plane = pln;
            return hatch;
        }
        private IGeoObject CreateInsert(ACadSharp.Entities.Insert insert)
        {
            // could also use insert.Explode()
            GeoObject.Block block = FindBlock(insert.Block.BlockEntity);
            if (block != null)
            {
                IGeoObject res = block.Clone();
                ModOp tranform = ModOp.Translate(new GeoVector(insert.InsertPoint.X, insert.InsertPoint.Y, insert.InsertPoint.Z)) *
                                 //ModOp.Translate(block.RefPoint.ToVector()) *
                                 ModOp.Rotate(CADability.GeoVector.ZAxis, SweepAngle.Deg(insert.Rotation)) *
                                 ModOp.Scale(insert.XScale, insert.YScale, insert.ZScale) *
                                 ModOp.Translate(CADability.GeoPoint.Origin - block.RefPoint);
                res.Modify(tranform);
                return res;
            }
            return null;
        }
        private IGeoObject CreateExplodedPolyline(ACadSharp.Entities.IPolyline polyline)
        {
            List<Entity> exploded = polyline.Explode().ToList();
            List<IGeoObject> path = new List<IGeoObject>();
            for (int i = 0; i < exploded.Count; i++)
            {
                IGeoObject ent = GeoObjectFromEntity(exploded[i]);
                if (ent != null) path.Add(ent);
            }
            GeoObject.Path go = GeoObject.Path.Construct();
            go.Set(new GeoObjectList(path), false, Precision.eps);
            if (go.CurveCount > 0) return go;
            return null;
        }
        private IGeoObject CreateMLine(ACadSharp.Entities.MLine mLine)
        {
            List<CadObject> exploded = mLine.Reactors.Values.ToList();
            List<IGeoObject> path = new List<IGeoObject>();
            foreach (CadObject obj in exploded)
            {
                if (obj is Entity ent)
                {
                    IGeoObject go = GeoObjectFromEntity(ent);
                    if (go != null) path.Add(go);
                }
            }
            GeoObjectList list = new GeoObjectList(path);
            GeoObjectList res = new GeoObjectList();
            while (list.Count > 0)
            {
                GeoObject.Path go = GeoObject.Path.Construct();
                if (go.Set(list, true, 1e-6))
                {
                    res.Add(go);
                }
                else
                {
                    break;
                }
            }
            if (res.Count > 1)
            {
                GeoObject.Block blk = GeoObject.Block.Construct();
                blk.Name = "MLINE " + mLine.Handle;
                blk.Set(res);
                return blk;
            }
            
            if (res.Count == 1) 
                return res[0];
             
            return null;
        }
        private string processAcadString(string acstr)
        {
            StringBuilder sb = new StringBuilder(acstr);
            sb.Replace("%%153", "Ø");
            sb.Replace("%%127", "°");
            sb.Replace("%%214", "Ö");
            sb.Replace("%%220", "Ü");
            sb.Replace("%%228", "ä");
            sb.Replace("%%246", "ö");
            sb.Replace("%%223", "ß");
            sb.Replace("%%u", ""); // underline
            sb.Replace("%%U", "");
            sb.Replace("%%D", "°");
            sb.Replace("%%d", "°");
            sb.Replace("%%P", "±");
            sb.Replace("%%p", "±");
            sb.Replace("%%C", "Ø");
            sb.Replace("%%c", "Ø");
            sb.Replace("%%%", "%");
            // and maybe some more, is there a documentation?
            return sb.ToString();
        }
        private IGeoObject CreateText(ACadSharp.Entities.TextEntity txt)
        {
            Text text = GeoObject.Text.Construct();
            string txtstring = processAcadString(txt.Value);
            if (txtstring.Trim().Length == 0) return null;
            string filename;
            string name;
            string typeface;
            bool bold;
            bool italic;
            filename = txt.Style.Filename;
            name = txt.Style.Name;
            typeface = "";
            bold = txt.Style.TrueType.HasFlag(FontFlags.Bold);
            italic = txt.Style.TrueType.HasFlag(FontFlags.Italic);
            GeoPoint insertionPoint = new GeoPoint(txt.InsertPoint.X, txt.InsertPoint.Y, txt.InsertPoint.Z);
            GeoVector normal = new GeoVector(txt.Normal.X, txt.Normal.Y, txt.Normal.Z);
            Angle a = Angle.Deg(txt.Rotation);
            double h = txt.Height;
            Plane plane = new Plane(insertionPoint, normal);

            bool isShx = false;
            if (typeface.Length > 0)
            {
                text.Font = typeface;
            }
            else
            {
                if (Path.GetExtension(filename).ToLower() == ".shx")
                {
                    filename = Path.GetFileNameWithoutExtension(filename);
                    isShx = true;
                }
                if (Path.GetExtension(filename).ToLower() == ".ttf")
                {
                    if (name != null && name.Length > 1) filename = name;
                    else filename = Path.GetFileNameWithoutExtension(filename);
                }
                text.Font = filename;
            }
            text.Bold = bold;
            text.Italic = italic;
            text.TextString = txtstring;
            text.Location = CADability.GeoPoint.Origin;
            text.LineDirection = h * CADability.GeoVector.XAxis; //plane.ToGlobal(new GeoVector2D(a));
            text.GlyphDirection = h * CADability.GeoVector.YAxis; // plane.ToGlobal(new GeoVector2D(a + SweepAngle.ToLeft));
            text.TextSize = h;
            text.Alignment = Text.AlignMode.Bottom;
            text.LineAlignment = Text.LineAlignMode.Left;
            txt.HorizontalAlignment = TextHorizontalAlignment.Left;

            switch (txt.HorizontalAlignment)
            {
                case TextHorizontalAlignment.Left:
                    text.LineAlignment = Text.LineAlignMode.Left;
                    break;
                case TextHorizontalAlignment.Center:
                    text.LineAlignment = Text.LineAlignMode.Center;
                    break;
                case TextHorizontalAlignment.Right:
                    text.LineAlignment = Text.LineAlignMode.Right;
                    break;
                case TextHorizontalAlignment.Aligned:
                    text.LineAlignment = Text.LineAlignMode.Left;
                    break;
                case TextHorizontalAlignment.Middle:
                    text.LineAlignment = Text.LineAlignMode.Center;
                    break;
                case TextHorizontalAlignment.Fit:
                    text.LineAlignment = Text.LineAlignMode.Left;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            switch (txt.VerticalAlignment)
            {
                case TextVerticalAlignmentType.Baseline:
                    text.Alignment = Text.AlignMode.Baseline;
                    break;
                case TextVerticalAlignmentType.Bottom:
                    text.Alignment = Text.AlignMode.Bottom;
                    break;
                case TextVerticalAlignmentType.Middle:
                    text.Alignment = Text.AlignMode.Center;
                    break;
                case TextVerticalAlignmentType.Top:
                    text.Alignment = Text.AlignMode.Top;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            text.Location = new GeoPoint(txt.InsertPoint.X, txt.InsertPoint.Y, txt.InsertPoint.Z);
            GeoVector2D dir2d = new GeoVector2D(a);
            GeoVector linedir = plane.ToGlobal(dir2d);
            GeoVector glyphdir = plane.ToGlobal(dir2d.ToLeft());
            text.LineDirection = linedir;
            text.GlyphDirection = glyphdir;
            text.TextSize = h;
            //if (isShx) h *= AdditionalShxFactor(text.Font);
            linedir.Length = h * txt.WidthFactor;
            if (!linedir.IsNullVector()) text.LineDirection = linedir;
            if (text.TextSize < Precision.eps) return null;
            return text;
        }
        private IGeoObject CreateDimension(ACadSharp.Entities.Dimension dimension)
        {
            // we could create a CADability Dimension object usind the dimension data and setting the block with the FindBlock values.
            // but then we would need a "CustomBlock" flag in the CADability Dimension object and also save this Block
            if (dimension.Block != null)
            {
                GeoObject.Block block = FindBlock(dimension.Block.BlockEntity);
                if (block != null)
                {
                    IGeoObject res = block.Clone();
                    return res;
                }
            }
            else
            {
                // make a dimension from the dimension data
            }
            return null;
        }
        private IGeoObject CreateMText(ACadSharp.Entities.MText mText)
        {
            // TODO: Support importing real MText either through stacking single text GeoObjects or adding multi-line support to CADAbility
            ACadSharp.Entities.TextEntity txt = new ACadSharp.Entities.TextEntity()
            {
                Value = mText.Value.Replace(@"\P", " "),
                Height = mText.Height,
                WidthFactor = 1.0,
                Rotation = mText.Rotation,
                ObliqueAngle = mText.Style.ObliqueAngle,
                // IsBackward = false,
                // IsUpsideDown = false,
                Style = mText.Style,
                InsertPoint = mText.InsertPoint,
                Normal = mText.Normal,
                VerticalAlignment = TextVerticalAlignmentType.Baseline,
                HorizontalAlignment = TextHorizontalAlignment.Left
            };
            return CreateText(txt);
        }
        private IGeoObject CreateLeader(ACadSharp.Entities.Leader leader)
        {
            GeoPoint[] leaderVertices = leader.Vertices.ToArray().Select(vert => new GeoPoint(vert.X, vert.Y, vert.Z)).ToArray();
            Plane ocs = CADability.Plane.FromPoints(leaderVertices, out _, out _);
            GeoObject.Block blk = GeoObject.Block.Construct();
            blk.Name = "Leader:" + leader.Handle;
            if (leader.CreationType != LeaderCreationType.CreatedWithoutAnnotation)
            {
                IGeoObject annotation = GeoObjectFromEntity(leader.AssociatedAnnotation);
                if (annotation != null) blk.Add(annotation);
            }
            GeoPoint[] vtx = new GeoPoint[leaderVertices.Length];
            for (int i = 0; i < vtx.Length; i++)
            {
                vtx[i] = ocs.ToGlobal(new GeoPoint2D(leaderVertices[i].x, leaderVertices[i].y));
            }
            GeoObject.Polyline pln = GeoObject.Polyline.Construct();
            pln.SetPoints(vtx, false);
            blk.Add(pln);
            return blk;
        }
        // Trying to consolidate polyline creation to IPolyline
        // private IGeoObject CreatePolyline2D(ACadSharp.Entities.Polyline2D polyline2D)
        // {
        //     List<CadObject> exploded = polyline2D.Explode();
        //     List<IGeoObject> path = new List<IGeoObject>();
        //     for (int i = 0; i < exploded.Count; i++)
        //     {
        //         IGeoObject ent = GeoObjectFromEntity(exploded[i]);
        //         if (ent != null) path.Add(ent);
        //     }
        //     GeoObject.Path go = GeoObject.Path.Construct();
        //     go.Set(new GeoObjectList(path), false, 1e-6);
        //     if (go.CurveCount > 0) return go;
        //     return null;
        // }
        // private IGeoObject CreatePolyline3D(netDxf.Entities.Polyline3D polyline3D)
        // {
        //     // polyline.Explode();
        //     bool hasWidth = false, hasBulges = false;
        //     for (int i = 0; i < polyline3D.Vertexes.Count; i++)
        //     {
        //         //hasBulges |= polyline.Vertexes[i].Bulge != 0.0;
        //         //hasWidth |= (polyline.Vertexes[i].StartWidth != 0.0) || (polyline.Vertexes[i].EndWidth != 0.0);
        //     }
        //     if (hasWidth && !hasBulges)
        //     {
        //
        //     }
        //     else
        //     {
        //         if (hasBulges)
        //         {   // must be in a single plane
        //
        //         }
        //         else
        //         {
        //             GeoObject.Polyline res = GeoObject.Polyline.Construct();
        //             for (int i = 0; i < polyline3D.Vertexes.Count; ++i)
        //             {
        //                 res.AddPoint(GeoPoint(polyline3D.Vertexes[i]));
        //             }
        //             res.IsClosed = polyline3D.IsClosed;
        //             if (res.GetExtent(0.0).Size < 1e-6) return null; // only identical points
        //             return res;
        //         }
        //     }
        //     return null;
        // }
        private IGeoObject CreatePoint(ACadSharp.Entities.Point point)
        {
            CADability.GeoObject.Point p = CADability.GeoObject.Point.Construct();
            p.Location = new GeoPoint(point.Location.X, point.Location.Y, point.Location.Z);
            p.Symbol = PointSymbol.Cross;
            return p;
        }
        private IGeoObject CreateMesh(ACadSharp.Entities.Mesh mesh)
        {
            GeoPoint[] vertices = new GeoPoint[mesh.Vertices.Count];
            for (int i = 0; i < vertices.Length; i++)
            {
                vertices[i] = new GeoPoint(mesh.Vertices[i].X, mesh.Vertices[i].Y, mesh.Vertices[i].Z);
            }
            List<Face> faces = new List<Face>();
            for (int i = 0; i < mesh.Faces.Count; i++)
            {
                int[] indices = mesh.Faces[i];
                if (indices.Length <= 3 || indices[3] == indices[2])
                {
                    if (indices[0] != indices[1] && indices[1] != indices[2])
                    {
                        try
                        {
                            Plane pln = new Plane(vertices[indices[0]], vertices[indices[1]], vertices[indices[2]]);
                            PlaneSurface surf = new PlaneSurface(pln);
                            Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[1]]), pln.Project(vertices[indices[2]]) });
                            SimpleShape ss = new SimpleShape(bdr);
                            Face fc = Face.MakeFace(surf, ss);
                            faces.Add(fc);
                        }
                        catch { };
                    }
                }
                else
                {
                    if (indices[0] != indices[1] && indices[1] != indices[2])
                    {
                        try
                        {
                            Plane pln = new Plane(vertices[indices[0]], vertices[indices[1]], vertices[indices[2]]);
                            PlaneSurface surf = new PlaneSurface(pln);
                            Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[1]]), pln.Project(vertices[indices[2]]) });
                            SimpleShape ss = new SimpleShape(bdr);
                            Face fc = Face.MakeFace(surf, ss);
                            faces.Add(fc);
                        }
                        catch { };
                    }
                    if (indices[2] != indices[3] && indices[3] != indices[0])
                    {
                        try
                        {
                            Plane pln = new Plane(vertices[indices[2]], vertices[indices[3]], vertices[indices[0]]);
                            PlaneSurface surf = new PlaneSurface(pln);
                            Border bdr = new Border(new GeoPoint2D[] { new GeoPoint2D(0.0, 0.0), pln.Project(vertices[indices[3]]), pln.Project(vertices[indices[0]]) });
                            SimpleShape ss = new SimpleShape(bdr);
                            Face fc = Face.MakeFace(surf, ss);
                            faces.Add(fc);
                        }
                        catch { };
                    }
                }
            }
            if (faces.Count > 1)
            {
                GeoObjectList sewed = Make3D.SewFacesAndShells(new GeoObjectList(faces.ToArray() as IGeoObject[]));
                if (sewed.Count == 1) return sewed[0];
                else
                {
                    GeoObject.Block blk = GeoObject.Block.Construct();
                    blk.Name = "Mesh";
                    blk.Set(new GeoObjectList(faces as ICollection<IGeoObject>));
                    return blk;
                }
            }
            else if (faces.Count == 1)
            {
                return faces[0];
            }
            else return null;
        }
        private static LwPolyline SplineToPolyline2D(Spline spline, int precision)
        {
            //List<Vector3> vetexes3D = spline.
        }

        private List<Vector3> SplinePolygonalVertexes(Spline spline, int precision)
        {
            bool isClosed = (spline.Flags & SplineFlags.Closed) != 0;
            bool isClosedPeriodic = (spline.Flags & SplineFlags.Periodic) != 0;
            return NurbsEvaluator(spline.ControlPoints, spline.Weights, spline.Knots, spline.Degree, isClosed, isClosedPeriodic, precision );
        }
        
        /// <summary>
        /// Calculate points along a NURBS curve.
        /// </summary>
        /// <param name="controls">List of spline control points.</param>
        /// <param name="weights">Spline control weights. If null the weights vector will be automatically initialized with 1.0.</param>
        /// <param name="knots">List of spline knot points. If null the knot vector will be automatically generated.</param>
        /// <param name="degree">Spline degree.</param>
        /// <param name="isClosed">Specifies if the spline is closed.</param>
        /// <param name="isClosedPeriodic">Specifies if the spline is closed and periodic.</param>
        /// <param name="precision">Number of vertexes generated.</param>
        /// <returns>A list vertexes that represents the spline.</returns>
        /// <remarks>
        /// NURBS evaluator provided by mikau16 based on Michael V. implementation, roughly follows the notation of http://cs.mtu.edu/~shene/PUBLICATIONS/2004/NURBS.pdf
        /// Added a few modifications to make it work for open, closed, and periodic closed splines.
        /// </remarks>
        private static List<Vector3> NurbsEvaluator(Vector3[] controls, double[] weights, double[] knots, int degree, bool isClosed, bool isClosedPeriodic, int precision)
        {
            if (precision < 2)
            {
                throw new ArgumentOutOfRangeException(nameof(precision), precision, "The precision must be equal or greater than two.");
            }

            // control points
            if (controls == null)
            {
                throw new ArgumentNullException(nameof(controls), "A spline entity with control points is required.");
            }

            int numCtrlPoints = controls.Length;

            if (numCtrlPoints == 0)
            {
                throw new ArgumentException("A spline entity with control points is required.", nameof(controls));
            }

            // weights
            if (weights == null)
            {
                // give the default 1.0 to the control points weights
                weights = new double[numCtrlPoints];
                for (int i = 0; i < numCtrlPoints; i++)
                {
                    weights[i] = 1.0;
                }
            }
            else if (weights.Length != numCtrlPoints)
            {
                throw new ArgumentException("The number of control points must be the same as the number of weights.", nameof(weights));
            }

            // knots
            if (knots == null)
            {
                knots = CreateKnotVector(numCtrlPoints, degree, isClosedPeriodic);
            }
            else
            {
                int numKnots;
                if (isClosedPeriodic)
                {
                    numKnots = numCtrlPoints + 2 * degree + 1;
                }
                else
                {
                    numKnots = numCtrlPoints + degree + 1;
                }
                if (knots.Length != numKnots)
                {
                    throw new ArgumentException("Invalid number of knots.");
                }
            }

            Vector3[] ctrl;
            double[] w;
            if (isClosedPeriodic)
            {
                ctrl = new Vector3[numCtrlPoints + degree];
                w = new double[numCtrlPoints + degree];
                for (int i = 0; i < degree; i++)
                {
                    int index = numCtrlPoints - degree + i;
                    ctrl[i] = controls[index];
                    w[i] = weights[index];
                }

                controls.CopyTo(ctrl, degree);
                weights.CopyTo(w, degree);
            }
            else
            {
                ctrl = controls;
                w = weights;
            }

            double uStart;
            double uEnd;
            List<Vector3> vertexes = new List<Vector3>();

            if (isClosed)
            {
                uStart = knots[0];
                uEnd = knots[knots.Length - 1];
            }
            else if (isClosedPeriodic)
            {
                uStart = knots[degree];
                uEnd = knots[knots.Length - degree - 1];
            }
            else
            {
                precision -= 1;
                uStart = knots[0];
                uEnd = knots[knots.Length - 1];
            }

            double uDelta = (uEnd - uStart) / precision;

            for (int i = 0; i < precision; i++)
            {
                double u = uStart + uDelta * i;
                vertexes.Add(C(ctrl, w, knots, degree, u));
            }

            if (!(isClosed || isClosedPeriodic))
            {
                vertexes.Add(ctrl[ctrl.Length - 1]);
            }

            return vertexes;
        }
        
        private static double[] CreateKnotVector(int numControlPoints, int degree, bool isPeriodic)
        {
            // create knot vector
            int numKnots;
            double[] knots;

            if (!isPeriodic)
            {
                numKnots = numControlPoints + degree + 1;
                knots = new double[numKnots];

                int i;
                for (i = 0; i <= degree; i++)
                {
                    knots[i] = 0.0;
                }

                for (; i < numControlPoints; i++)
                {
                    knots[i] = i - degree;
                }

                for (; i < numKnots; i++)
                {
                    knots[i] = numControlPoints - degree;
                }
            }
            else
            {
                numKnots = numControlPoints + 2 * degree + 1;
                knots = new double[numKnots];

                double factor = 1.0 / (numControlPoints - degree);
                for (int i = 0; i < numKnots; i++)
                {
                    knots[i] = (i - degree) * factor;
                }
            }

            return knots;
        }
        private static Vector3 C(Vector3[] ctrlPoints, double[] weights, double[] knots, int degree, double u)
        {
            Vector3 vectorSum = Vector3.Zero;
            double denominatorSum = 0.0;

            // optimization suggested by ThVoss
            for (int i = 0; i < ctrlPoints.Length; i++)
            {
                double n = N(knots, i, degree, u);
                denominatorSum += n * weights[i];
                vectorSum += weights[i] * n * ctrlPoints[i];
            }

            // avoid possible divided by zero error, this should never happen
            if (Math.Abs(denominatorSum) < double.Epsilon)
            {
                return Vector3.Zero;
            }

            return (1.0 / denominatorSum) * vectorSum;
        }
        
        private static double N(double[] knots, int i, int p, double u)
        {
            if (p <= 0)
            {
                if (knots[i] <= u && u < knots[i + 1])
                {
                    return 1;
                }

                return 0.0;
            }

            double leftCoefficient = 0.0;
            if (!(Math.Abs(knots[i + p] - knots[i]) < double.Epsilon))
            {
                leftCoefficient = (u - knots[i]) / (knots[i + p] - knots[i]);
            }

            double rightCoefficient = 0.0; // article contains error here, denominator is Knots[i + p + 1] - Knots[i + 1]
            if (!(Math.Abs(knots[i + p + 1] - knots[i + 1]) < double.Epsilon))
            {
                rightCoefficient = (knots[i + p + 1] - u) / (knots[i + p + 1] - knots[i + 1]);
            }

            return leftCoefficient * N(knots, i, p - 1, u) + rightCoefficient * N(knots, i + 1, p - 1, u);
        }
        
    }
}