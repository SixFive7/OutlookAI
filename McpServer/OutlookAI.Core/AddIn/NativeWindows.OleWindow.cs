using System;
using System.Runtime.InteropServices;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// The part of <see cref="NativeWindows"/> only Core's code compiles: from an Outlook Explorer or Inspector to
    /// its window, through the OLE window interface. Not C# 5, and not needed by the testbed's scripts, which
    /// compile NativeWindows.cs alone.
    /// </summary>
    public static partial class NativeWindows
    {
        /// <summary>
        /// The window of an Outlook Explorer or Inspector, through <see cref="IOleWindow"/>; zero when Outlook will
        /// not say. It reads the window only - nothing is shown, activated or created.
        /// </summary>
        public static IntPtr WindowOf(object? outlookWindowObject)
        {
            if (outlookWindowObject == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                if (outlookWindowObject is IOleWindow window && window.GetWindow(out IntPtr handle) == 0)
                {
                    return handle;
                }
            }
            catch (Exception ex) when (ex is InvalidCastException || ex is COMException || ex is InvalidComObjectException)
            {
            }

            return IntPtr.Zero;
        }

        /// <summary>The OLE window interface every Outlook Explorer and Inspector implements: the one way from one to its window.</summary>
        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleWindow
        {
            /// <summary>The window.</summary>
            [PreserveSig]
            int GetWindow(out IntPtr phwnd);

            /// <summary>Context-sensitive help, never used here.</summary>
            [PreserveSig]
            int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        }
    }
}
