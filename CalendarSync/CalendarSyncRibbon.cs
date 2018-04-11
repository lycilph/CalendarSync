using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace CalendarSync
{
    [ComVisible(true)]
    public class CalendarSyncRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbon_id)
        {
            return GetResourceText("CalendarSync.CalendarSyncRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbon_ui)
        {
            ribbon = ribbon_ui;
        }

        public void ShowCalendarSyncWindow(Office.IRibbonControl control)
        {
            var win = new CalendarSyncWindow();
            win.Show();
        }

        #region Helpers

        private static string GetResourceText(string resource_name)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resource_names = asm.GetManifestResourceNames();
            foreach (var t in resource_names)
            {
                if (string.Compare(resource_name, t, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resource_reader = new StreamReader(asm.GetManifestResourceStream(t)))
                    {
                        if (resource_reader != null)
                        {
                            return resource_reader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}