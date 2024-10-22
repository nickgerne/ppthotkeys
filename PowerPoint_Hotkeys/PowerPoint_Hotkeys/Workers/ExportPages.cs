using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PPTShortcuts.Workers
{
    internal class ExportPages
    {
        private Application _application;

        public ExportPages(Application application)
        {
            _application = application;
        }

        internal void EmailSelectedPages()
        {
            var activePresentation = _application.ActivePresentation;
            var activeWindow = _application.ActiveWindow;

            //active slides not selected
            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionSlides)
                return;

            //delete previous export tags
            foreach (Slide slide in activePresentation.Slides)
            {
                slide.Tags.Delete("ToExport");
            }

            string currentPresentationName = Path.GetFileNameWithoutExtension(activePresentation.Name);
            string emailSubject = currentPresentationName;
            string tmpPresentationName = currentPresentationName + "-";

            var selectedSlieCount = activeWindow.Selection.SlideRange.Count;

            int tagSlideCount = 0;
            foreach (Slide slide in activeWindow.Selection.SlideRange)
            {
                slide.Tags.Add("ToExport", "1");
                tagSlideCount += 1;

                if (slide != activeWindow.Selection.SlideRange[activeWindow.Selection.SlideRange.Count])
                {
                    tmpPresentationName += slide.SlideIndex + ",";
                }
                else
                {
                    tmpPresentationName += slide.SlideIndex;
                }
            }

            //save a new copy of the presentation with only the slides you want to export
            string tmpPath = @"C:\";
            string tmpPresentationPath = Path.Combine(tmpPath, tmpPresentationName + ".pptx");
            var tmpActivePresentationPath = activePresentation.Path;

            activePresentation.SaveCopyAs(tmpPresentationPath);

            Presentation tmpPresentation = _application.Presentations.Open(tmpPresentationPath);

            int deleteCounter = 0;

            //delete non-selected slides
            for (int i = 1; i <= tmpPresentation.Slides.Count; i++)
            {
                Slide slide = tmpPresentation.Slides[i];
                if (slide.Tags["ToExport"] != "1")
                {
                    slide.Delete();
                    deleteCounter += 1;
                    i -= 1;
                }
            }

            tmpPresentation.Save();
            tmpPresentation.Close();

            //now open new email from Outlook
            Microsoft.Office.Interop.Outlook.Application outlookApp = null;
            try
            {
                outlookApp = (Microsoft.Office.Interop.Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
            }
            catch
            {
                outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            }

            MailItem mailMessage = outlookApp.CreateItem(OlItemType.olMailItem);
            mailMessage.Subject = tmpPresentationName;

            mailMessage.Attachments.Add(tmpPresentationPath);
            mailMessage.Display();

            if (outlookApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                outlookApp = null;
            }

        }
    }
}