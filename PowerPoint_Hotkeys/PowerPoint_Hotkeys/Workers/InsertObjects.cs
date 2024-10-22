using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using TextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;

namespace PPTShortcuts.Workers
{
    internal class InsertObjects
    {
        private Application _application;

        public InsertObjects(Application application)
        {
            _application = application;
        }

        //for rectangle sticky margin
        private float defaultTopMargin = 5;
        private float defaultRightMargin = 5;
        private float defaultBottomMargin = 5;
        private float defaultLeftMargin = 5;


        internal void InsertLine(bool isArrow)
        {
            //https://learn.microsoft.com/en-us/office/vba/api/powerpoint.lineformat
            var activeWindow = _application.ActiveWindow;

            Shape newShape = activeWindow.Selection.SlideRange.Shapes.AddLine(BeginX: 50, BeginY: 50, EndX: 200, EndY: 50);
            newShape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            newShape.Line.BackColor.RGB = Convert.ToInt32("000000", 16);

            if (isArrow)
            {
                newShape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            }
        }

        internal void InsertRectangle()
        {
            var activeWindow = _application.ActiveWindow;

            //https://learn.microsoft.com/en-us/office/vba/api/powerpoint.colorformat
            Shape newShape = activeWindow.Selection.SlideRange.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, Left: 20, Top: 50, Width: 400, Height: 50);
            newShape.Fill.ForeColor.RGB = Convert.ToInt32("1E90FF", 16);        //BG color of the shape  --- this gets messed up for some reason


            newShape.Line.Visible = MsoTriState.msoFalse;   //hide the border

            TextFrame textFrame = newShape.TextFrame;
            textFrame.MarginTop = defaultTopMargin;
            textFrame.MarginBottom = defaultBottomMargin;
            textFrame.MarginLeft = defaultLeftMargin;
            textFrame.MarginRight = defaultRightMargin;
            textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;

            TextRange textRange = textFrame.TextRange;
            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            textRange.Text = "TBD";

            Font font = textRange.Font;
            font.Size = 12;
            font.Color.RGB = Convert.ToInt32("ffffff", 16);
        }

        internal void InsertTextbox()
        {
            var activeWindow = _application.ActiveWindow;

            Shape newShape = activeWindow.Selection.SlideRange.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, Left: 50, Top: 50, Width: 200, Height: 40);
            newShape.Line.Visible = MsoTriState.msoFalse;   //hide the border

            TextFrame textFrame = newShape.TextFrame;
            textFrame.MarginTop = 0;
            textFrame.MarginBottom = 0;
            textFrame.MarginLeft = 0;
            textFrame.MarginRight = 0;
            textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            textFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            TextRange textRange = textFrame.TextRange;
            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            textRange.Text = "TBD";

            Font font = textRange.Font;
            font.Size = 12;
            font.Color.RGB = Convert.ToInt32("000000", 16);
        }


        internal void InsertStickyNote()
        {
            var activeWindow = _application.ActiveWindow;
            var slideRange = activeWindow.Selection.SlideRange;

            int numOfStickies = (from Shape shape in slideRange.Shapes
                                 where shape.Name.StartsWith("Sticky")
                                 select shape).Count();

            float xPosition = _application.ActivePresentation.PageSetup.SlideWidth - (105 * (numOfStickies + 1));  //this puts it to the RHS

            Shape newShape = activeWindow.Selection.SlideRange.Shapes.AddShape(MsoAutoShapeType.msoShapeSnip2DiagRectangle, Left: xPosition, Top: 50, Width: 200, Height: 50);
            newShape.Line.Visible = MsoTriState.msoFalse;   //hide the border
            newShape.Fill.ForeColor.RGB = int.Parse("49407");
            newShape.Name = "Sticky-" + Guid.NewGuid().ToString();


            TextFrame textFrame = newShape.TextFrame;
            textFrame.MarginTop = defaultTopMargin;
            textFrame.MarginBottom = defaultBottomMargin;
            textFrame.MarginLeft = defaultLeftMargin;
            textFrame.MarginRight = defaultRightMargin;
            textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            textFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            TextRange textRange = textFrame.TextRange;
            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            textRange.Text = "Note:";

            Font font = textRange.Font;
            font.Size = 10;
            font.Color.RGB = int.Parse("000000");
        }
    }
}