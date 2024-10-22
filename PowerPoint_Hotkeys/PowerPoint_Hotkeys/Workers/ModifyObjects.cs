using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using TextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;

namespace PPTShortcuts.Workers
{
    internal class ModifyObjects
    {
        private Application _application;

        public ModifyObjects(Application application)
        {
            _application = application;
        }

        internal void SwapObject()
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - or less than 2 shapes selected
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes || activeWindow.Selection.ShapeRange.Count < 2)
                return;


            ShapeRange shapes;
            float left1, left2, top1, top2;

            if (activeWindow.Selection.ShapeRange.Count == 2)
            {
                shapes = activeWindow.Selection.ShapeRange;

                left1 = shapes[1].Left;
                left2 = shapes[2].Left;
                top1 = shapes[1].Top;
                top2 = shapes[2].Top;

                shapes[1].Left = left2;
                shapes[2].Left = left1;
                shapes[1].Top = top2;
                shapes[2].Top = top1;
            }
            else if (activeWindow.Selection.HasChildShapeRange)
            {
                ShapeRange childShapes = activeWindow.Selection.ChildShapeRange;

                if (childShapes.Count == 2)
                {
                    left1 = childShapes[1].Left;
                    left2 = childShapes[2].Left;
                    top1 = childShapes[1].Top;
                    top2 = childShapes[2].Top;

                    childShapes[1].Left = left2;
                    childShapes[2].Left = left1;
                    childShapes[1].Top = top2;
                    childShapes[2].Top = top1;
                }
            }

        }


        internal void MakeObjectsSameWidth()
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - or less than 2 shapes selected
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes || activeWindow.Selection.ShapeRange.Count < 2)
                return;

            if (activeWindow.Selection.HasChildShapeRange)
            {
                activeWindow.Selection.ChildShapeRange.Width = activeWindow.Selection.ChildShapeRange[1].Width;
            }
            else
            {
                activeWindow.Selection.ShapeRange.Width = activeWindow.Selection.ShapeRange[1].Width;
            }
        }

        internal void MakeObjectsSameHeight()
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - or less than 2 shapes selected
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes || activeWindow.Selection.ShapeRange.Count < 2)
                return;

            if (activeWindow.Selection.HasChildShapeRange)
            {
                activeWindow.Selection.ChildShapeRange.Height = activeWindow.Selection.ChildShapeRange[1].Height;
            }
            else
            {
                activeWindow.Selection.ShapeRange.Height = activeWindow.Selection.ShapeRange[1].Height;
            }
        }

        internal void AutosizeShapeToFitText()
        {
            var activeWindow = _application.ActiveWindow;

            //not a shape - exit
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionText)
                return;


            TextFrame textFrame;
            TextRange textRange;
            if (activeWindow.Selection.HasChildShapeRange)
            {
                //activeWindow.Selection.ChildShapeRange.Height = activeWindow.Selection.ChildShapeRange[1].Height;
                ShapeRange shapeRange = activeWindow.Selection.ChildShapeRange;

                foreach (Shape shape in shapeRange)
                {
                    textFrame = shape.TextFrame;
                    textFrame.MarginTop = 0;
                    textFrame.MarginBottom = 0;
                    textFrame.MarginLeft = 0;
                    textFrame.MarginRight = 0;
                    textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                    textFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
            else
            {
                ShapeRange shapeRange = activeWindow.Selection.ShapeRange;
                foreach (Shape shape in shapeRange)
                {
                    textFrame = shape.TextFrame;
                    textFrame.MarginTop = 0;
                    textFrame.MarginBottom = 0;
                    textFrame.MarginLeft = 0;
                    textFrame.MarginRight = 0;
                    textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                    textFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }

        }

        internal void ChangeObjectAutosizeProperty()
        {
            var activeWindow = _application.ActiveWindow;

            bool skip = true;
            if (activeWindow.Selection.Type == PpSelectionType.ppSelectionText) { skip = false;  }

            ShapeRange shapeRange = activeWindow.Selection.ShapeRange;
            foreach (Shape shape  in shapeRange)
            {
                if (shape.Name.StartsWith("Sticky") || shape.Type == MsoShapeType.msoTextBox) { skip = false; }
                
            }

            if (skip) 
                return;


            TextFrame textFrame;
            TextRange textRange;

            if (activeWindow.Selection.HasChildShapeRange)
            {
                ShapeRange childShapeRange = activeWindow.Selection.ChildShapeRange;
                foreach (Shape shape in childShapeRange) {
                    textFrame = shape.TextFrame;

                    var originalAutosizeValue = textFrame.AutoSize;
                    textFrame.AutoSize = originalAutosizeValue == PpAutoSize.ppAutoSizeShapeToFitText
                        ? PpAutoSize.ppAutoSizeNone
                        : PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
            else 
            {
                foreach (Shape shape in shapeRange)
                {
                    textFrame = shape.TextFrame;

                    var originalAutosizeValue = textFrame.AutoSize;
                    textFrame.AutoSize = originalAutosizeValue == PpAutoSize.ppAutoSizeShapeToFitText
                        ? PpAutoSize.ppAutoSizeNone
                        : PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
        }
    }
}