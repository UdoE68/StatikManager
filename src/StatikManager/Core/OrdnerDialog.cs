using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace StatikManager.Core
{
    /// <summary>
    /// Öffnet den modernen Windows-Ordner-Auswahl-Dialog (Vista+ IFileOpenDialog via COM).
    /// Fällt automatisch auf den klassischen FolderBrowserDialog zurück.
    /// </summary>
    public static class OrdnerDialog
    {
        private const uint FOS_PICKFOLDERS     = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint SIGDN_FILESYSPATH   = 0x80058000;

        /// <summary>
        /// Zeigt den Ordner-Auswahl-Dialog und gibt den gewählten Pfad zurück,
        /// oder null wenn der Benutzer abbricht.
        /// </summary>
        /// <param name="startPfad">Ordner, in dem der Dialog startet.</param>
        /// <param name="titel">Titel des Dialogs.</param>
        /// <param name="besitzer">WPF-Elternfenster (für modale Darstellung).</param>
        public static string? Zeigen(string startPfad = "", string titel = "Ordner wählen",
                                     Window? besitzer = null)
        {
            var hwnd = besitzer != null ? new WindowInteropHelper(besitzer).Handle : IntPtr.Zero;
            try   { return ZeigenModern(hwnd, startPfad, titel); }
            catch { return ZeigenKlassisch(startPfad, titel); }
        }

        // ── Moderner Dialog (IFileOpenDialog) ─────────────────────────────────

        private static string? ZeigenModern(IntPtr hwnd, string startPfad, string titel)
        {
            var dialog = (IFileDialog)new FileOpenDialogImpl();
            try
            {
                dialog.GetOptions(out uint opt);
                dialog.SetOptions(opt | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                dialog.SetTitle(titel);

                // Startpfad setzen
                if (Directory.Exists(startPfad))
                {
                    var iid = typeof(IShellItem).GUID;
                    SHCreateItemFromParsingName(startPfad, IntPtr.Zero, ref iid, out IntPtr pItem);
                    if (pItem != IntPtr.Zero)
                    {
                        var item = (IShellItem)Marshal.GetObjectForIUnknown(pItem);
                        dialog.SetDefaultFolder(item);
                        dialog.SetFolder(item);
                        Marshal.Release(pItem);
                    }
                }

                int hr = dialog.Show(hwnd);
                if (hr != 0) return null; // Abgebrochen oder Fehler

                dialog.GetResult(out IShellItem result);
                result.GetDisplayName(SIGDN_FILESYSPATH, out string path);
                return path;
            }
            finally
            {
                Marshal.FinalReleaseComObject(dialog);
            }
        }

        // ── Fallback (alter FolderBrowserDialog) ──────────────────────────────

        private static string? ZeigenKlassisch(string startPfad, string titel)
        {
            var d = new System.Windows.Forms.FolderBrowserDialog
            {
                Description         = titel,
                SelectedPath        = Directory.Exists(startPfad) ? startPfad : "",
                ShowNewFolderButton = false
            };
            return d.ShowDialog() == System.Windows.Forms.DialogResult.OK ? d.SelectedPath : null;
        }

        // ── P/Invoke ──────────────────────────────────────────────────────────

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            out IntPtr ppv);

        // ── COM: CLSID_FileOpenDialog ─────────────────────────────────────────

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7"),
         ClassInterface(ClassInterfaceType.None)]
        private class FileOpenDialogImpl { }

        // ── COM: IFileDialog (IID 42F85136-DB7E-439C-85F1-E4075D135FC8) ────────
        // Vtable-Reihenfolge: IUnknown(0-2), IModalWindow::Show(3), IFileDialog(4-26)

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialog
        {
            [PreserveSig] int Show(IntPtr hwndOwner);
            void SetFileTypes(uint c, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder([MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            void SetFolder([MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            void GetFolder([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            void AddPlace([MarshalAs(UnmanagedType.Interface)] IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close([MarshalAs(UnmanagedType.Error)] int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
        }

        // ── COM: IShellItem (IID 43826D1E-E718-42EE-BC55-A1E261C37BFE) ─────────

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare([MarshalAs(UnmanagedType.Interface)] IShellItem psi, uint hint, out int piOrder);
        }
    }
}
