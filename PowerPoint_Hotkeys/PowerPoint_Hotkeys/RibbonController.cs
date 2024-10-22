using Microsoft.Office.Core;
using PPTShortcuts.Workers;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RibbonController();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PPTShortcuts
{
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private AlignObjects _alignObjects;
        private DistributeObjects _distributeObjects;
        private InsertObjects _insertObjects;
        private ModifyObjects _modifyObjects;
        private GroupObjects _groupObjects;
        private ExportPages _exportPages;

        public RibbonController()
        {
        }

        private void InitAlignObjects()
        {
            if (_alignObjects == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _alignObjects = new AlignObjects(applicaiton);
                }
            }
        }

        private void InitDistributeObjects()
        {
            if (_distributeObjects == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _distributeObjects = new DistributeObjects(applicaiton);
                }
            }
        }

        private void InitInsertObjects()
        {
            if (_insertObjects == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _insertObjects = new InsertObjects(applicaiton);
                }
            }
        }

        private void InitModifyObjects()
        {
            if (_modifyObjects == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _modifyObjects = new ModifyObjects(applicaiton);
                }
            }
        }

        private void InitGroupObjects()
        {
            if (_groupObjects == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _groupObjects = new GroupObjects(applicaiton);
                }
            }
        }

        private void InitExportPages()
        {
            if (_exportPages == null)
            {
                {
                    var applicaiton = Globals.HotKeys.Application;
                    _exportPages = new ExportPages(applicaiton);
                }
            }
        }

        public void AlignLeft(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignLefts);
        }
        public void AlignRight(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignRights);
        }
        public void AlignTop(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignTops);
        }
        public void AlignBottom(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignBottoms);
        }
        public void AlignCenter(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignCenters);
        }
        public void AlignMiddle(IRibbonControl control)
        {
            InitAlignObjects();
            _alignObjects.Align(MsoAlignCmd.msoAlignMiddles);
        }

        public void DistributeVertically(IRibbonControl control)
        {
            InitDistributeObjects();
            _distributeObjects.DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeVertically);
        }
        public void DistributeHorizontally(IRibbonControl control)
        {
            InitDistributeObjects();
            _distributeObjects.DistributeObjectsHorizontallyOrVertically(MsoDistributeCmd.msoDistributeHorizontally);
        }

        public void GroupByColumns()
        {
            InitGroupObjects();
            _groupObjects.GroupObjectsByColumn();
        }
        public void GroupByRows()
        {
            InitGroupObjects();
            _groupObjects.GroupShapesByRow();
        }

        public void EmailSelectedPages()
        {
            InitExportPages();
            _exportPages.EmailSelectedPages();
        }

        public void InsertTextbox(IRibbonControl control)
        {
            InitInsertObjects();
            _insertObjects.InsertTextbox();
        }

        public void InsertLine(IRibbonControl control)
        {
            InitInsertObjects();
            _insertObjects.InsertLine(false);
        }

        public void InsertArrow(IRibbonControl control)
        {
            InitInsertObjects();
            _insertObjects.InsertLine(true);
        }

        public void InsertStickNote(IRibbonControl control)
        {
            InitInsertObjects();
            _insertObjects.InsertStickyNote();
        }

        public void InsertRectangle(IRibbonControl control)
        {
            InitInsertObjects();
            _insertObjects.InsertRectangle();
        }

        public void MakeObjectSameWidth(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.MakeObjectsSameWidth();
        }

        public void MakeObjectSameHeight(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.MakeObjectsSameHeight();
        }

        public void MakeObjectSameWidthAndHeight(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.MakeObjectsSameWidth();
            _modifyObjects.MakeObjectsSameHeight();
        }

        public void SwapObjects(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.SwapObject();
        }

        public void AutosizeShapeToFitText(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.AutosizeShapeToFitText();
        }

        public void ChangeAutosizeProperty(IRibbonControl control)
        {
            InitModifyObjects();
            _modifyObjects.ChangeObjectAutosizeProperty();
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPoint_Hotkeys.RibbonController.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
