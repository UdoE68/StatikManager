using System;
using System.Runtime.InteropServices;

namespace StatikManager.Modules.Dokumente
{
    /// <summary>
    /// COM-sichtbares Scripting-Objekt für den IE-WebBrowser.
    /// Wird als window.external gesetzt – JavaScript ruft LoescheAusschnitt() auf.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AusschnittScriptingObjekt
    {
        internal Action<string>? LoeschCallback;

        /// <summary>Wird aus JavaScript aufgerufen: window.external.LoescheAusschnitt(dateiname)</summary>
        public void LoescheAusschnitt(string pngDateiname)
        {
            LoeschCallback?.Invoke(pngDateiname);
        }
    }
}
