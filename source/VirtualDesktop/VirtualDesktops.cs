using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using WindowsDesktop.Interop;

namespace WindowsDesktop
{
	public partial class VirtualDesktops : IDisposable
	{
        private readonly ComObjects _comObjects;
        private readonly IDisposable _listener;

        private bool IsCurrentVirtualDesktop(IntPtr handle)
        {
            return _comObjects.VirtualDesktopManager.IsWindowOnCurrentVirtualDesktop(handle);
        }

        private void MoveToDesktop(IntPtr hWnd, VirtualDesktop virtualDesktop)
        {
            int processId;
            NativeMethods.GetWindowThreadProcessId(hWnd, out processId);

            if (Process.GetCurrentProcess().Id == processId)
            {
                var guid = virtualDesktop.Id;
                _comObjects.VirtualDesktopManager.MoveWindowToDesktop(hWnd, ref guid);
            }
            else
            {
                IApplicationView view;
                _comObjects.ApplicationViewCollection.GetViewForHwnd(hWnd, out view);
                _comObjects.VirtualDesktopManagerInternal.MoveViewToDesktop(view, virtualDesktop.ComObject);
            }
        }

        private VirtualDesktop wrap(IVirtualDesktop x) => new VirtualDesktop(this, x);

        /// <summary>
        /// Gets the virtual desktop that is currently displayed.
        /// </summary>
        public VirtualDesktop Current => wrap(_comObjects.VirtualDesktopManagerInternal.GetCurrentDesktop());

		public VirtualDesktops()
		{
            _comObjects = new ComObjects();
            _listener = RegisterListener();
		}

        public void Dispose()
        {
            _listener.Dispose();
            _comObjects.Dispose();
        }

		/// <summary>
		/// Creates a virtual desktop.
		/// </summary>
		public VirtualDesktop Create() => wrap(_comObjects.VirtualDesktopManagerInternal.CreateDesktopW());

        [EditorBrowsable(EditorBrowsableState.Never)]
        public VirtualDesktop FromComObject(IVirtualDesktop desktop) => wrap(desktop);

		/// <summary>
		/// Returns the virtual desktop of the specified identifier.
		/// </summary>
		public VirtualDesktop FromId(Guid desktopId)
		{
			try
			{
                return wrap(_comObjects.VirtualDesktopManagerInternal.FindDesktop(ref desktopId));
			}
			catch (COMException ex) when (ex.Match(HResult.TYPE_E_ELEMENTNOTFOUND))
			{
				return null;
			}
		}

		/// <summary>
		/// Returns the virtual desktop that the specified window is located.
		/// </summary>
		public VirtualDesktop FromHwnd(IntPtr hwnd)
		{
			if (hwnd == IntPtr.Zero) return null;

			try
			{
				var desktopId = _comObjects.VirtualDesktopManager.GetWindowDesktopId(hwnd);
				return wrap(_comObjects.VirtualDesktopManagerInternal.FindDesktop(ref desktopId));
			}
			catch (COMException ex) when (ex.Match(HResult.REGDB_E_CLASSNOTREG, HResult.TYPE_E_ELEMENTNOTFOUND))
			{
				return null;
			}
		}

        /// <summary>
        /// Display the virtual desktop.
        /// </summary>
        public void Switch(VirtualDesktop desktop) => _comObjects.VirtualDesktopManagerInternal.SwitchDesktop(desktop.ComObject);

        /// <summary>
        /// Remove the virtual desktop, specifying a virtual desktop that display after destroyed.
        /// </summary>
        public void Remove(VirtualDesktop desktop, VirtualDesktop fallbackDesktop = null)
        {
            if (fallbackDesktop == null)
            {
                fallbackDesktop = this.FirstOrDefault(x => x.Id != desktop.Id) ?? Create();
            }

            _comObjects.VirtualDesktopManagerInternal.RemoveDesktop(desktop.ComObject, fallbackDesktop.ComObject);
        }

        /// <summary>
        /// Returns a virtual desktop on the left.
        /// </summary>
        public VirtualDesktop GetLeft(VirtualDesktop desktop)
        {
            try
            {
                return wrap(_comObjects.VirtualDesktopManagerInternal.GetAdjacentDesktop(desktop.ComObject, AdjacentDesktop.LeftDirection));
            }
            catch (COMException ex) when (ex.Match(HResult.TYPE_E_OUTOFBOUNDS))
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a virtual desktop on the right.
        /// </summary>
        public VirtualDesktop GetRight(VirtualDesktop desktop)
        {
            try
            {
                return wrap(_comObjects.VirtualDesktopManagerInternal.GetAdjacentDesktop(desktop.ComObject, AdjacentDesktop.RightDirection));
            }
            catch (COMException ex) when (ex.Match(HResult.TYPE_E_OUTOFBOUNDS))
            {
                return null;
            }
        }

        private IApplicationView GetApplicationView(IntPtr hWnd)
        {
            IApplicationView view;
            _comObjects.ApplicationViewCollection.GetViewForHwnd(hWnd, out view);

            return view;
        }

        private string GetAppId(IntPtr hWnd)
        {
            string appId;
            GetApplicationView(hWnd).GetAppUserModelId(out appId);

            return appId;
        }
    }
}
