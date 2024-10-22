using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PPTShortcuts.Workers
{
    internal class DistributeObjects
    {
        private Application _application;

        public DistributeObjects(Application application)
        {
            _application = application;
        }

        internal void DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd distributionCmd)
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - or less than 3 shapes selected
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes || activeWindow.Selection.ShapeRange.Count < 3)
                return;


            ShapeRange shapes;
            if (activeWindow.Selection.HasChildShapeRange)
            {
                shapes = activeWindow.Selection.ChildShapeRange;
                if (shapes.Count > 2)
                    activeWindow.Selection.ChildShapeRange.Distribute(distributionCmd, MsoTriState.msoFalse);
            }
            else
            {
                activeWindow.Selection.ShapeRange.Distribute(distributionCmd, MsoTriState.msoFalse);
            }
        }
    }
}