using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Gma.System.MouseKeyHook;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;


using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using TextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;
using Font = Microsoft.Office.Interop.PowerPoint.Font;

namespace PowerPoint_Hotkeys
{
    public partial class ThisAddIn
    {

        /*
         *  ------------------------------------------------------------------------------------
         *  Object alignment
         *  ------------------------------------------------------------------------------------
         * 
         * Ctrl + Shift + Left Arrow            - align objects left
         * Ctrl + Shift + Right Arrow           - align objects right
         * Ctrl + Shift + Up Arrow              - align objects top
         * Ctrl + Shift + Down Arrow            - align objects bottom
         * Ctrl + Shift + Minus (by backspace)  - align objects middle  (like a shiskabob) 
         * Ctrl + Shift + C                     - align objects center (all objects would center on a flag pole)
         *  
         *  ------------------------------------------------------------------------------------
         *  distribute objects
         *  ------------------------------------------------------------------------------------
         * Ctrl + Shift + H                     - distribute objects horizontally
         * Ctrl + Shift + V                     - distribute objects vertically
         * 
         * Ctrl + Shift + |                     - Swap objects
         * Ctrl + Alt + [                       - Objects same width
         * Ctrl + Alt + ]                       - Objects same height
         * Ctrl + Alt + \                       - Objects same width and height
         * 
         * 
         * 
         *  ------------------------------------------------------------------------------------
         *  textbox adjustments
         *  ------------------------------------------------------------------------------------
         *** Ctrl + Shift + :                   - decrease textbox margins
         *** Ctrl + Shift + '                   - increase textbox margins
         * Ctrl + 8                             - Make textbox fit text
         *** Ctrl + Shift + L                   - increase line spacing of a textbox
         *** Ctrl + Shift + K                   - decrease line spacing of a textbox
         *
         *  ------------------------------------------------------------------------------------
         *  insert object
         *  ------------------------------------------------------------------------------------
         * Ctrl + Shift + T                     - insert textbox
         * Ctrl + Shift + A                     - insert arrow
         * Ctrl + Shift + L                     - insert line
         * Ctrl + Shift + R                     - insert rectangle
         * Ctrl + 0                             - insert sticky on the RHS of the slide         
         * 
         * 
         * 
         * --- Forbidden commands (PPT has them)
         * Ctrl + Shift + [                     - Move object fwd / backward
         * Ctrl + Shift + ]                     - Move object fwd / backward
         * Ctrl + Shift + M                     - Creates a new slide (e.g., Ctrl + M)
         * 
         * 
         * Feature list:
         *      - align objects left
         *      - align objects right
         *      - align objects bottom
         *      - align objects top
         *      - align objects middle
         *      - align objects center
         *      
         *      - distribute objects vertically
         *      - distribute objects horiztonally
         *      
         *      - swap objects
         *      - make objects same width
         *      - make objects same height
         *      - make objects same width and height
         *      
         *      - insert textbox
         *      - insert rectangle (can't get BG color to work correctly)
         *      - insert sticky on RHS
         *      - insert line
         *      - insert arrow
         *      
         *      
         *******************************************
         *TODO feature list
         *      - group selection by columns   (https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsRowsAndColumns.bas)
         *      - group selection by rows       (https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsRowsAndColumns.bas)
         *      - decrease textbox margins      https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas       > ObjectsMarginsDecrease()
         *      - increase textbox margins      https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas       > ObjectsMarginsIncrease()
         *      - Make textbox fit text
         *      -
         *      
         *      - increase line spacing of a textbox  https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas     > ObjectsDecreaseLineSpacing()
         *      - decrease line spacing of a textbox    https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas     > ObjectsIncreaseLineSpacing()
         *      
         *      
         *      - increase line weight
         *      - decrease line weight
         *      
         */
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region shape alignment
            var comboAlignLeft = Combination.FromString("Control+Shift+Left");
            var comboAlignRight = Combination.FromString("Control+Shift+Right");
            var comboAlignTop = Combination.FromString("Control+Shift+Up");
            var comboAlignBottom = Combination.FromString("Control+Shift+Down");
            var comboAlignMiddle = Combination.FromString("Control+Shift+OemMinus");
            var comboAlignCenter = Combination.FromString("Control+Shift+C");

            var comboDistributeHoriztonally = Combination.FromString("Control+Shift+H");
            var comboDistributeVertically = Combination.FromString("Control+Shift+V");


            Action actionAlignLeft = () => AlignObject(MsoAlignCmd.msoAlignLefts);
            Action actionAlignRight = () => AlignObject(MsoAlignCmd.msoAlignRights);
            Action actionAlignTop = () => AlignObject(MsoAlignCmd.msoAlignTops);
            Action actionAlignBottom = () => AlignObject(MsoAlignCmd.msoAlignBottoms);
            Action actionAlignMiddle = () => AlignObject(MsoAlignCmd.msoAlignMiddles);
            Action actionAlignCenter = () => AlignObject(MsoAlignCmd.msoAlignCenters);

            Action actionDistributeHorizontally = () => DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeHorizontally);
            Action actionDistributeVertically = () => DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeVertically);
            #endregion



            #region modify objects

            var comboSwapObjects = Combination.FromString("Control+Shift+OemPipe");
            var comboMakeObjectsSameWidth = Combination.FromString("Control+Alt+OemOpenBrackets");
            var comboMakeObjectsSameHeight = Combination.FromString("Control+Alt+OemCloseBrackets");
            var comboMakeObjectsSameWidthAndHeight = Combination.FromString("Control+Alt+OemPipe");

            var comboAutosizeShapeToFitText = Combination.FromString("Control+D8");



            Action actionSwapObjects = () => SwapObject();
            Action actionMakeObjectsSameWidth = () => MakeObjectsSameWidth();
            Action actionMakeObjectsSameHeight = () => MakeObjectsSameHeight();
            Action actionMakeObjectsSameWidthAndHeight = () => { MakeObjectsSameWidth(); MakeObjectsSameHeight(); };
            Action actionAutosizeShapeToFitText = () => { AutosizeShapeToFitText(); };

            #endregion


            #region insert objects
            var comboInsertRectangle = Combination.FromString("Control+Shift+R");
            var comboInsertTextbox = Combination.FromString("Control+Shift+T");
            var comboInsertArrow = Combination.FromString("Control+Shift+A");
            var comboInsertLine = Combination.FromString("Control+Shift+L");
            var comboInsertSticky = Combination.FromString("Control+D0");       // the '0' on the main keyboard (not the num pad)


            Action actionInsertRectrangle = () => InsertRegtangle();
            Action actionInsertLine = () => InsertLine(false);
            Action actionInsertArrow = () => InsertLine(true);
            Action actionInsertTextbox = () => InsertTextbox();
            Action actionInsertStickyNote = () => InsertStickyNote();
            #endregion


            var comboAssignments = new Dictionary<Combination, Action>
            {
                {comboAlignLeft, actionAlignLeft },
                {comboAlignRight, actionAlignRight },
                {comboAlignTop, actionAlignTop },
                {comboAlignBottom, actionAlignBottom},
                {comboAlignMiddle, actionAlignMiddle},
                {comboAlignCenter, actionAlignCenter},

                {comboDistributeHoriztonally, actionDistributeHorizontally},
                {comboDistributeVertically, actionDistributeVertically},

                {comboSwapObjects, actionSwapObjects},
                {comboMakeObjectsSameWidth, actionMakeObjectsSameWidth},
                {comboMakeObjectsSameHeight, actionMakeObjectsSameHeight},
                {comboMakeObjectsSameWidthAndHeight, actionMakeObjectsSameWidthAndHeight},


                {comboInsertRectangle, actionInsertRectrangle},
                {comboInsertLine, actionInsertLine},
                {comboInsertArrow, actionInsertArrow},
                {comboInsertTextbox, actionInsertTextbox},
                {comboInsertSticky, actionInsertStickyNote},
                {comboAutosizeShapeToFitText, actionAutosizeShapeToFitText }
            };

            Hook.AppEvents().OnCombination(comboAssignments);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            Hook.AppEvents().Dispose();
        }

        private void AlignObject(MsoAlignCmd alignment)
        {
            var activeWindow = Application.ActiveWindow;

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

        private void DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd distributionCmd)
        {
            var activeWindow = Application.ActiveWindow;

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

        private void SwapObject()
        {
            var activeWindow = Application.ActiveWindow;

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

        private void MakeObjectsSameWidth()
        {
            var activeWindow = Application.ActiveWindow;

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

        private void MakeObjectsSameHeight()
        {
            var activeWindow = Application.ActiveWindow;

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

        private void InsertRegtangle()
        {
            var activeWindow = Application.ActiveWindow;




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

        private void InsertLine(bool isArrow)
        {
            //https://learn.microsoft.com/en-us/office/vba/api/powerpoint.lineformat
            var activeWindow = Application.ActiveWindow;

            Shape newShape = activeWindow.Selection.SlideRange.Shapes.AddLine(BeginX: 50, BeginY: 50, EndX: 200, EndY: 50);
            newShape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            newShape.Line.BackColor.RGB = Convert.ToInt32("000000", 16);

            if (isArrow)
            {
                newShape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            }
        }

        private void InsertTextbox()
        {
            var activeWindow = Application.ActiveWindow;

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

        private float defaultTopMargin = 5;
        private float defaultBottomMargin = 5;
        private float defaultLeftMargin = 5;
        private float defaultRightMargin = 5;


        private void InsertStickyNote()
        {
            var activeWindow = Application.ActiveWindow;
            var slideRange = activeWindow.Selection.SlideRange;

            int numOfStickies = (from Shape shape in slideRange.Shapes
                                 where shape.Name.StartsWith("Sticky")
                                 select shape).Count();

            float xPosition = Application.ActivePresentation.PageSetup.SlideWidth - (105 * (numOfStickies + 1));  //this puts it to the RHS

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

        private void AutosizeShapeToFitText()
        {
            var activeWindow = Application.ActiveWindow;

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



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
