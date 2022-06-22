using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModelPlagin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            var levels = GetLevels(doc);
            Level level1 = levels.Where(x => x.Name == "Level 1").FirstOrDefault();
            Level level2 = levels.Where(x => x.Name == "Level 2").FirstOrDefault();

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);

            List<Wall> walls = new List<Wall>();
            Transaction transaction = new Transaction(doc, "Create model");
            transaction.Start();

            walls = CreateWalls(doc, width, depth, level1, level2);

            AddDoor(doc, level1, walls[0]);

            AddWindow(doc, level1, walls[1]);
            AddWindow(doc, level1, walls[2]);
            AddWindow(doc, level1, walls[3]);

            AddRoof(doc, level2, walls);

            transaction.Commit();

            return Result.Succeeded;
        }

        private void AddRoof(Document doc, Level level, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Generic - 125mm"))
                .Where(x => x.FamilyName.Equals("Basic Roof"))
                .FirstOrDefault();

            LocationCurve curve = walls[1].Location as LocationCurve;
            XYZ p1 = curve.Curve.GetEndPoint(0);
            XYZ p2 = curve.Curve.GetEndPoint(1);
            XYZ p3 = (p1 + p2) / 2;
            XYZ dz = new XYZ(0, 0, 5);
            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;

            double x1 = p1.X+dt;
            double x2 = p2.X+dt;
            double x3 = p3.X+dt;
            double y1 = p1.Y-dt;
            double y2 = p2.Y+dt;
            double y3 = p3.Y;
            double wallHeight = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();

            double zCommon = wallHeight;
            XYZ startPoint = new XYZ(x1, y1, zCommon);
            XYZ middlePoint = new XYZ(x3, y3, zCommon+dz.Z);
            XYZ endPoint = new XYZ(x2, y2, zCommon);

            CurveArray curveArray = new CurveArray();
            curveArray.Append(Line.CreateBound(startPoint, middlePoint));
            curveArray.Append(Line.CreateBound(middlePoint, endPoint));


            ReferencePlane plane = doc.Create.NewReferencePlane2(startPoint, middlePoint, endPoint, doc.ActiveView);
                /*(walls[0].Location as LocationCurve).Curve.GetEndPoint(0),
                                                                 (walls[0].Location as LocationCurve).Curve.GetEndPoint(1),
                                                                 (walls[2].Location as LocationCurve).Curve.GetEndPoint(0), 
                                                                 doc.ActiveView);*/
            ExtrusionRoof extrusionRoof = doc.Create.NewExtrusionRoof(curveArray, plane, level, roofType, 0, 
                walls[0].get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()+wallWidth);
            extrusionRoof.get_Parameter(BuiltInParameter.ROOF_EAVE_CUT_PARAM).Set(33618);
            /*RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Generic - 125mm"))
                .Where(x => x.FamilyName.Equals("Basic Roof"))
                .FirstOrDefault();

            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));
            points.Add(new XYZ(-dt, dt, 0));
            points.Add(new XYZ(-dt, -dt, 0));

            Application application = doc.Application;
            CurveArray footPrint = application.Create.NewCurveArray();
            for (int i = 0; i < 4; i++)
            {
                LocationCurve curve = walls[i].Location as LocationCurve;
                XYZ p1 = curve.Curve.GetEndPoint(0);
                XYZ p2 = curve.Curve.GetEndPoint(1);
                Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
                footPrint.Append(line);
            }

            ModelCurveArray modelCurveArray = new ModelCurveArray();
            FootPrintRoof roof = doc.Create.NewFootPrintRoof(footPrint, level, roofType, out modelCurveArray);

            ModelCurveArrayIterator iterator = modelCurveArray.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                ModelCurve modelCurve = iterator.Current as ModelCurve;
                roof.set_DefinesSlope(modelCurve, true);
                roof.set_SlopeAngle(modelCurve, 0.5);
            }
            foreach (ModelCurve m in modelCurveArray)
            {
                roof.set_DefinesSlope(m, true);
                roof.set_SlopeAngle(m, 0.5);
            }*/
        }

        private void AddWindow(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("1200 x 1500mm"))
                .Where(x => x.FamilyName.Equals("M_Window-Casement-Double"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            if (!windowType.IsActive)
                windowType.Activate();

            var window = doc.Create.NewFamilyInstance(point, windowType, wall, level1, StructuralType.NonStructural);
            window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(UnitUtils.ConvertToInternalUnits(800, UnitTypeId.Millimeters));
        }

        private void AddDoor(Document doc, Level level1, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .OfType<FamilySymbol>()
                    .Where(x => x.Name.Equals("0915 x 2134mm"))
                    .Where(x => x.FamilyName.Equals("M_Single-Flush"))
                    .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level1, StructuralType.NonStructural);
        }

        public List<Level> GetLevels(Document doc)
        {
            List<Level> levels = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .OfType<Level>()
                        .ToList();
            return levels;
        }

        public List<Wall> CreateWalls(Document doc, double width, double depth, Level baseLevel, Level topLevel)
        {
            List<XYZ> points = new List<XYZ>();
            double dx = width / 2;
            double dy = depth / 2;
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, baseLevel.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(topLevel.Id);
            }

            return walls;
        }
    }
}
