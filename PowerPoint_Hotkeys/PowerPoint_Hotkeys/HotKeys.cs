using System;
using System.Collections.Generic;
using Gma.System.MouseKeyHook;
using Microsoft.Office.Core;
using PPTShortcuts.Workers;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint_Hotkeys;

/*
        *  ------------------------------------------------------------------------------------
        *  textbox adjustments
        *  ------------------------------------------------------------------------------------
        *** Ctrl + Shift + :                   - decrease textbox margins
        *** Ctrl + Shift + '                   - increase textbox margins
        *** Ctrl + Shift + L                   - increase line spacing of a textbox
        *** Ctrl + Shift + K                   - decrease line spacing of a textbox
        *
        * --- Forbidden commands (PPT has them)
        * Ctrl + Shift + [                     - Move object fwd / backward
        * Ctrl + Shift + ]                     - Move object fwd / backward
        * Ctrl + Shift + M                     - Creates a new slide (e.g., Ctrl + M)
        * 
        * 
        *******************************************
        *TODO feature list
        *      - decrease textbox margins      https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas       > ObjectsMarginsDecrease()
        *      - increase textbox margins      https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas       > ObjectsMarginsIncrease()
        *      - Make textbox fit text
        *      
        *      - increase line spacing of a textbox  https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas     > ObjectsDecreaseLineSpacing()
        *      - decrease line spacing of a textbox    https://github.com/iappyx/Instrumenta/blob/main/src/Modules/ModuleObjectsText.bas     > ObjectsIncreaseLineSpacing()
        *      
        *      - increase line weight
        *      - decrease line weight
        */

namespace PPTShortcuts
{
    public partial class HotKeys 
    {


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //icons:  https://bert-toolkit.com/imagemso-list.html
            //icons2: https://codekabinett.com/download/Microsoft-Office-2016_365-imageMso-Gallery.pdf
            return new RibbonController();
        }


        private void InitAlignObjects()
        {
            Combination comboAlignLeft = Combination.FromString("Control+Alt+Left");
            Combination comboAlignRight = Combination.FromString("Control+Alt+Right");
            Combination comboAlignTop = Combination.FromString("Control+Alt+Up");
            Combination comboAlignBottom = Combination.FromString("Control+Alt+Down");
            Combination comboAlignMiddle = Combination.FromString("Control+Alt+OemMinus");
            Combination comboAlignCenter = Combination.FromString("Control+Alt+C");

            AlignObjects alignObjects = new AlignObjects(_application);

            Action actionAlignLeft = () => alignObjects.Align(MsoAlignCmd.msoAlignLefts);
            Action actionAlignRight = () => alignObjects.Align(MsoAlignCmd.msoAlignRights);
            Action actionAlignTop = () => alignObjects.Align(MsoAlignCmd.msoAlignTops);
            Action actionAlignBottom = () => alignObjects.Align(MsoAlignCmd.msoAlignBottoms);
            Action actionAlignMiddle = () => alignObjects.Align(MsoAlignCmd.msoAlignMiddles);
            Action actionAlignCenter = () => alignObjects.Align(MsoAlignCmd.msoAlignCenters);

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboAlignLeft, actionAlignLeft);
            _comboAssignments.Add(comboAlignRight, actionAlignRight);
            _comboAssignments.Add(comboAlignTop, actionAlignTop);
            _comboAssignments.Add(comboAlignBottom, actionAlignBottom);
            _comboAssignments.Add(comboAlignMiddle, actionAlignMiddle);
            _comboAssignments.Add(comboAlignCenter, actionAlignCenter);
        }

        private void InitDistributeObjects()
        {
            Combination comboDistributeVertically = Combination.FromString("Shift+Alt+V");
            Combination comboDistributeHorizontally = Combination.FromString("Shift+Alt+H");

            DistributeObjects distributeObjects = new DistributeObjects(_application);

            Action actionDistributeVertically = () => distributeObjects.DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeVertically);
            Action actionDistributeHorizontally = () => distributeObjects.DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeHorizontally);

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboDistributeVertically, actionDistributeVertically);
            _comboAssignments.Add(comboDistributeHorizontally, actionDistributeHorizontally);
        }

        private void InitModifyObjects()
        {
            Combination comboSwapObjects = Combination.FromString("Control+Shift+OemPipe");
            Combination comboMakeObjectsSameWidth = Combination.FromString("Control+Alt+OemOpenBrackets");
            Combination comboMakeObjectsSameHeight = Combination.FromString("Control+Alt+OemCloseBrackets");
            Combination comboMakeObjectsSameWidthAndHeight = Combination.FromString("Control+Alt+OemPipe");
            Combination comboAutosizeShapeToFitText = Combination.FromString("Control+D8");
            Combination comboChangeObjectAutosizeProperty = Combination.FromString("Control+D7");

            ModifyObjects modifyObjects = new ModifyObjects(_application);

            Action actionSwapObjects = () => modifyObjects.SwapObject();
            Action actionMakeObjectsSameWidth = () => modifyObjects.MakeObjectsSameWidth();
            Action actionMakeObjectsSameHeight = () => modifyObjects.MakeObjectsSameHeight();
            Action actionMakeObjectsSameWidthAndHeight = () => { modifyObjects.MakeObjectsSameWidth(); modifyObjects.MakeObjectsSameHeight(); };
            Action actionAutosizeShapeToFitText = () => modifyObjects.AutosizeShapeToFitText();
            Action actionChangeObjectAutosizeProperty = () => modifyObjects.ChangeObjectAutosizeProperty();

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboSwapObjects, actionSwapObjects);
            _comboAssignments.Add(comboMakeObjectsSameWidth, actionMakeObjectsSameWidth);
            _comboAssignments.Add(comboMakeObjectsSameHeight, actionMakeObjectsSameHeight);
            _comboAssignments.Add(comboMakeObjectsSameWidthAndHeight, actionMakeObjectsSameWidthAndHeight);
            _comboAssignments.Add(comboAutosizeShapeToFitText, actionAutosizeShapeToFitText);
            _comboAssignments.Add(comboChangeObjectAutosizeProperty, actionChangeObjectAutosizeProperty);
        }

        private void InitGroupObjects()
        {
            Combination comboGroupObjectsByRow = Combination.FromString("Shift+Alt+R");
            Combination comboGroupObjectsByColumn = Combination.FromString("Shift+Alt+C");

            GroupObjects groupObjects = new GroupObjects(_application);

            Action actionGroupObjectsByRow = () => groupObjects.GroupShapesByRow();
            Action actionGroupObjectsByColumn = () => groupObjects.GroupObjectsByColumn();

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboGroupObjectsByRow, actionGroupObjectsByRow);
            _comboAssignments.Add(comboGroupObjectsByColumn, actionGroupObjectsByColumn);
        }

        private void InitExportSelectedPages()
        {
            Combination comboExportSelectedPages = Combination.FromString("Control+D1");

            ExportPages exportPages = new ExportPages(_application);

            Action actionExportSelectedPages = () => exportPages.EmailSelectedPages();

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboExportSelectedPages, actionExportSelectedPages);
        }

        private void InitInsertObjects()
        {
            Combination comboInsertRectangle = Combination.FromString("Control+Shift+R");
            Combination comboInsertTextbox = Combination.FromString("Control+Shift+T");
            Combination comboInsertArrow = Combination.FromString("Control+Shift+A");
            Combination comboInsertLine = Combination.FromString("Control+Shift+L");
            Combination comboInsertSticky = Combination.FromString("Control+D0");

            InsertObjects insertObjects = new InsertObjects(_application);

            Action actionInsertRectangle = () => insertObjects.InsertRectangle();
            Action actionInsertTextbox = () => insertObjects.InsertTextbox();
            Action actionInsertArrow = () => insertObjects.InsertLine(true);
            Action actionInsertLine = () => insertObjects.InsertLine(false);
            Action actionInsertSticky = () => insertObjects.InsertStickyNote();

            //add to the dictionary to wire it up
            _comboAssignments.Add(comboInsertRectangle, actionInsertRectangle);
            _comboAssignments.Add(comboInsertTextbox, actionInsertTextbox);
            _comboAssignments.Add(comboInsertArrow, actionInsertArrow);
            _comboAssignments.Add(comboInsertLine, actionInsertLine);
            _comboAssignments.Add(comboInsertSticky, actionInsertSticky);
        }


        private Application _application;
        private Dictionary<Combination, Action> _comboAssignments;
        private void HotKeys_Startup(object sender, System.EventArgs e)
        {
            _application = Application;
            _comboAssignments = new Dictionary<Combination, Action>();

            InitAlignObjects();
            InitDistributeObjects();
            InitInsertObjects();
            InitModifyObjects();
            InitGroupObjects();
            InitExportSelectedPages();

            Hook.AppEvents().OnCombination(_comboAssignments);
        }

        private void HotKeys_Shutdown(object sender, System.EventArgs e)
        {
            Hook.AppEvents().Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(HotKeys_Startup);
            this.Shutdown += new System.EventHandler(HotKeys_Shutdown);
        }

        #endregion
    }
}
