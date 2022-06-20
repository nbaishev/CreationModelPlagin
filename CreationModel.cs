using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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

            var walls = CreateWalls(doc, width, depth, level1, level2);

            return Result.Succeeded;
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

            Transaction transaction = new Transaction(doc, "Create model");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, baseLevel.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(topLevel.Id);
            }
            transaction.Commit();

            return walls;
        }
    }
}
