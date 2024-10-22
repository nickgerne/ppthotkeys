using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PPTShortcuts.Workers
{
    internal class GroupObjects
    {
        private Application _application;

        public GroupObjects(Application application)
        {
            _application = application;
        }

        internal void GroupObjectsByColumn()
        {
            var activeWindow = _application.ActiveWindow;
            Slide slide = activeWindow.View.Slide;

            //not a shape - exit
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
                return;

            //create a dictionary to store the shape groups
            Dictionary<int, List<Shape>> shapeGroups = new Dictionary<int, List<Shape>>();
            Shape slideShape;

            ShapeRange shapeRange = activeWindow.Selection.ShapeRange;

            foreach(Shape shape in shapeRange)
            {
                shape.Name = "Shape" + shape.Id;
            }

            //group shapes by vertical position
            foreach(Shape shape in shapeRange)
            {
                bool doesShapeGroupExist = false;

                foreach(var kvp in shapeGroups)
                {
                    var listShape = kvp.Value[0];

                    if ((shape.Left + shape.Width) >= listShape.Left && shape.Left <= (listShape.Left + listShape.Width))
                    {
                        kvp.Value.Add(shape);
                        doesShapeGroupExist = true;
                        break;
                    }
                }

                if (!doesShapeGroupExist)
                {
                    shapeGroups.Add(shape.Id, new List<Shape> { shape });
                }
            }

            //group shapes in PPT
            foreach (var kvp in shapeGroups)
            {
                List<Shape> groupShapes = kvp.Value;
                string[] shapeNames = new string[groupShapes.Count];

                for (int i = 0; i < groupShapes.Count; i++)
                {
                    shapeNames[i] = groupShapes[i].Name;
                }

                if (groupShapes.Count > 1)
                {
                    slide.Shapes.Range(shapeNames).Group();
                }
            }
        }


        internal void GroupShapesByRow()
        {
            var activeWindow = _application.ActiveWindow;
            Slide slide = activeWindow.View.Slide;

            //not a shape - exit
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
                return;

            //create a dictionary to store the shape groups
            Dictionary<int, List<Shape>> shapeGroups = new Dictionary<int, List<Shape>>();
            Shape slideShape;

            ShapeRange shapeRange = activeWindow.Selection.ShapeRange;

            foreach (Shape shape in shapeRange)
            {
                shape.Name = "Shape" + shape.Id;
            }

            //group shapes by vertical position
            foreach (Shape shape in shapeRange)
            {
                bool doesShapeGroupExist = false;

                foreach (var kvp in shapeGroups)
                {
                    var listShape = kvp.Value[0];

                    //if ((shape.Left + shape.Width) >= listShape.Left && shape.Left <= (listShape.Left + listShape.Width))
                    if ((shape.Top + shape.Height) >= listShape.Top && shape.Top <= (listShape.Top + listShape.Height))
                    {
                        kvp.Value.Add(shape);
                        doesShapeGroupExist = true;
                        break;
                    }
                }

                if (!doesShapeGroupExist)
                {
                    shapeGroups.Add(shape.Id, new List<Shape> { shape });
                }
            }

            //group shapes in PPT
            foreach (var kvp in shapeGroups)
            {
                List<Shape> groupShapes = kvp.Value;
                string[] shapeNames = new string[groupShapes.Count];

                for (int i = 0; i < groupShapes.Count; i++)
                {
                    shapeNames[i] = groupShapes[i].Name;
                }

                if (groupShapes.Count > 1)
                {
                    slide.Shapes.Range(shapeNames).Group();
                }
            }

        }
    }
}