using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// The one way OutlookAI.Core and the add-in let go of a COM reference and tell a failed COM call
    /// from a bug (the code-sharing refactor's stage S4, survey group A05). Until then the add-in kept
    /// its own release, which skipped the <see cref="Marshal.IsComObject"/> test, and OutlookComSession
    /// kept the rest; OutlookComSession's members of the same names forward here.
    /// </summary>
    public static class ComObjects
    {
        /// <summary>Releases a COM reference; null and objects that are not COM objects are left alone.</summary>
        public static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        /// <summary>
        /// Late-bound COM failures do not always surface as COMException: the dynamic
        /// binder maps E_INVALIDARG to ArgumentException, E_POINTER to
        /// ArgumentNullException, and binding problems to RuntimeBinderException. Optional
        /// COM paths must treat all of these as "that call did not work here". Every
        /// caller with an optional COM path (Phase 2 fact 2) uses this same test.
        /// </summary>
        public static bool IsCallFailure(Exception ex)
        {
            // MissingMemberException: the dynamic COM binder's shape for a failed
            // GetIDsOfNames ("Could not get dispatch ID for X") - live-observed with
            // 0x800706BA (RPC_S_SERVER_UNAVAILABLE) when Outlook exited mid-call
            // (soak, 2026-07-24). InvalidComObjectException: a detached RCW - what a
            // released-by-the-exit-watcher reference throws when used mid-flight
            // (ComGateway already treats it as a disconnect shape).
            return ex is COMException
                || ex is ArgumentException
                || ex is InvalidCastException
                || ex is MissingMemberException
                || ex is InvalidComObjectException
                || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException;
        }

        /// <summary>
        /// The pointer that identifies a COM object: its IUnknown. Two references to one object - two
        /// runtime callable wrappers, or one object reached through two interfaces - give the same pointer,
        /// which makes it the only sound COM identity test. The pointer is released before it is returned:
        /// it is a value to compare, never one to call through, and it identifies the object only while a
        /// reference keeps the object alive. Zero for null and when the object cannot give one.
        /// </summary>
        public static IntPtr IdentityOf(object? comObject)
        {
            if (comObject == null)
            {
                return IntPtr.Zero;
            }

            IntPtr p = IntPtr.Zero;
            try
            {
                p = Marshal.GetIUnknownForObject(comObject);
                return p;
            }
            catch (Exception)
            {
                return IntPtr.Zero;
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    _ = Marshal.Release(p);
                }
            }
        }

        /// <summary>Whether two references are to one COM object, by <see cref="IdentityOf"/>.</summary>
        public static bool SameObject(object? a, object? b)
        {
            if (ReferenceEquals(a, b))
            {
                return true;
            }

            if (a == null || b == null)
            {
                return false;
            }

            IntPtr identity = IdentityOf(a);
            return identity != IntPtr.Zero && identity == IdentityOf(b);
        }

        /// <summary>A failed call in a few words: the HRESULT of a COMException, the exception's type otherwise.</summary>
        public static string DescribeFailure(Exception ex)
        {
            return ex is COMException com
                ? string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", com.HResult)
                : ex.GetType().Name;
        }
    }
}
