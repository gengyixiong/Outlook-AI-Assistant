using System;
using System.Runtime.InteropServices;

namespace OutlookAiAssistant.OutlookIntegration
{
    internal static class ComRelease
    {
        public static void Release(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Outlook owns these objects; cleanup errors are non-fatal.
            }
        }
    }
}

