using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PPTShortcuts.Workers
{
    internal class AlignObjects
    {
        private Application _application;
        public AlignObjects(Application application)
        {
            _application = application;
        }
        internal void Align(MsoAlignCmd alignment)
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - exit
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
                return;


            //have only 1 object - align to the slide
            if (activeWindow.Selection.ShapeRange.Count == 1)
            {
                activeWindow.Selection.ShapeRange.Align(alignment, MsoTriState.msoTrue);
                return;
            }

            ShapeRange shapes;
            if (activeWindow.Selection.HasChildShapeRange)
            {
                shapes = activeWindow.Selection.ChildShapeRange;
                if (shapes.Count > 1)
                {
                    var left1 = shapes[1].Left;
                    var left2 = shapes[shapes.Count].Left;
                    shapes.Align(alignment, MsoTriState.msoFalse);

                    foreach (Shape shape in shapes)
                    {
                        shape.Left = left1;
                    }
                }
            }
            else
            {
                shapes = activeWindow.Selection.ShapeRange;
                shapes.Align(alignment, MsoTriState.msoFalse);
            }
        }
    }
}
