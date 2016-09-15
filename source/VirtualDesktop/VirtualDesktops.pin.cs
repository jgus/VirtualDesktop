using System;
using System.Collections.Generic;
using System.Linq;
using WindowsDesktop.Interop;

namespace WindowsDesktop
{
	partial class VirtualDesktops
	{
		public bool IsPinnedWindow(IntPtr hWnd)
		{
			return _comObjects.VirtualDesktopPinnedApps.IsViewPinned(GetApplicationView(hWnd));
		}

		public void PinWindow(IntPtr hWnd)
		{
			var view = GetApplicationView(hWnd);

			if (!_comObjects.VirtualDesktopPinnedApps.IsViewPinned(view))
			{
                _comObjects.VirtualDesktopPinnedApps.PinView(view);
			}
		}

		public void UnpinWindow(IntPtr hWnd)
		{
			var view = GetApplicationView(hWnd);

			if (_comObjects.VirtualDesktopPinnedApps.IsViewPinned(view))
			{
                _comObjects.VirtualDesktopPinnedApps.UnpinView(view);
			}
		}

		public bool IsPinnedApplication(string appId)
		{
			return _comObjects.VirtualDesktopPinnedApps.IsAppIdPinned(appId);
		}

		public void PinApplication(string appId)
		{
			if (!_comObjects.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{
                _comObjects.VirtualDesktopPinnedApps.PinAppID(appId);
			}
		}

		public void UnpinApplication(string appId)
		{
			if (_comObjects.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{
                _comObjects.VirtualDesktopPinnedApps.UnpinAppID(appId);
			}
		}
	}
}
