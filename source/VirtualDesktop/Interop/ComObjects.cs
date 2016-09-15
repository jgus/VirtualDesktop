using System;
using WindowsDesktop.Internal;

namespace WindowsDesktop.Interop
{
    public class ComObjects : IDisposable
	{
		private ExplorerRestartListenerWindow _listenerWindow;

		internal IVirtualDesktopManager VirtualDesktopManager { get; private set; }
		internal VirtualDesktopManagerInternal VirtualDesktopManagerInternal { get; private set; }
		internal IVirtualDesktopNotificationService VirtualDesktopNotificationService { get; private set; }
		internal IVirtualDesktopPinnedApps VirtualDesktopPinnedApps { get; private set; }
		internal IApplicationViewCollection ApplicationViewCollection { get; private set; }

		internal ComObjects()
		{
            _listenerWindow = new ExplorerRestartListenerWindow(() => GetObjects());
            _listenerWindow.Show();
            GetObjects();
		}

        private void GetObjects()
        {
            VirtualDesktopManager = GetVirtualDesktopManager();
            VirtualDesktopManagerInternal = VirtualDesktopManagerInternal.GetInstance();
            VirtualDesktopNotificationService = GetVirtualDesktopNotificationService();
            VirtualDesktopPinnedApps = GetVirtualDesktopPinnedApps();
            ApplicationViewCollection = GetApplicationViewCollection();
        }

        public void Dispose()
        {
            _listenerWindow?.Close();
        }

		private class ExplorerRestartListenerWindow : TransparentWindow
		{
			private uint _explorerRestertedMessage;
			private readonly Action _action;

			public ExplorerRestartListenerWindow(Action action)
			{
				this.Name = nameof(ExplorerRestartListenerWindow);
				this._action = action;
			}

			public override void Show()
			{
				base.Show();
				this._explorerRestertedMessage = NativeMethods.RegisterWindowMessage("TaskbarCreated");
			}

			protected override IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
			{
				if (msg == this._explorerRestertedMessage)
				{
					this._action();
					return IntPtr.Zero;
				}

				return base.WndProc(hwnd, msg, wParam, lParam, ref handled);
			}
		}

        private IVirtualDesktopManager GetVirtualDesktopManager()
		{
			var vdmType = Type.GetTypeFromCLSID(CLSID.VirtualDesktopManager);
			var instance = Activator.CreateInstance(vdmType);

			return (IVirtualDesktopManager)instance;
		}

        private IVirtualDesktopNotificationService GetVirtualDesktopNotificationService()
		{
			var shellType = Type.GetTypeFromCLSID(CLSID.ImmersiveShell);
			var shell = (IServiceProvider)Activator.CreateInstance(shellType);

			object ppvObject;
			shell.QueryService(CLSID.VirtualDesktopNotificationService, typeof(IVirtualDesktopNotificationService).GUID, out ppvObject);

			return (IVirtualDesktopNotificationService)ppvObject;
		}

        private IVirtualDesktopPinnedApps GetVirtualDesktopPinnedApps()
		{
			var shellType = Type.GetTypeFromCLSID(CLSID.ImmersiveShell);
			var shell = (IServiceProvider)Activator.CreateInstance(shellType);

			object ppvObject;
			shell.QueryService(CLSID.VirtualDesktopPinnedApps, typeof(IVirtualDesktopPinnedApps).GUID, out ppvObject);

			return (IVirtualDesktopPinnedApps)ppvObject;
		}

        private IApplicationViewCollection GetApplicationViewCollection()
		{
			var shellType = Type.GetTypeFromCLSID(CLSID.ImmersiveShell);
			var shell = (IServiceProvider)Activator.CreateInstance(shellType);

			object ppvObject;
			shell.QueryService(typeof(IApplicationViewCollection).GUID, typeof(IApplicationViewCollection).GUID, out ppvObject);

			return (IApplicationViewCollection)ppvObject;
		}
    }
}
