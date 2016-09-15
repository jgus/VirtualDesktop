using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using WindowsDesktop.Internal;
using WindowsDesktop.Interop;

namespace WindowsDesktop
{
	partial class VirtualDesktops
	{
		private uint? dwCookie;
		private VirtualDesktopNotificationListener listener;

		/// <summary>
		/// Occurs when a virtual desktop is created.
		/// </summary>
		public event EventHandler<VirtualDesktop> Created;
		public event EventHandler<VirtualDesktopDestroyEventArgs> DestroyBegin;
		public event EventHandler<VirtualDesktopDestroyEventArgs> DestroyFailed;

		/// <summary>
		/// Occurs when a virtual desktop is destroyed.
		/// </summary>
		public event EventHandler<VirtualDesktopDestroyEventArgs> Destroyed;

		[EditorBrowsable(EditorBrowsableState.Never)]
		public event EventHandler ApplicationViewChanged;

		/// <summary>
		/// Occurs when a current virtual desktop is changed.
		/// </summary>
		public event EventHandler<VirtualDesktopChangedEventArgs> CurrentChanged;


		internal IDisposable RegisterListener()
		{
			var service = _comObjects.VirtualDesktopNotificationService;
			listener = new VirtualDesktopNotificationListener(this);
			dwCookie = service.Register(listener);
			return Disposable.Create(() => service.Unregister(dwCookie.Value));
		}

		private class VirtualDesktopNotificationListener : IVirtualDesktopNotification
		{
            private VirtualDesktops virtualDesktops;

            public VirtualDesktopNotificationListener(VirtualDesktops virtualDesktops)
            {
                this.virtualDesktops = virtualDesktops;
            }

            void IVirtualDesktopNotification.VirtualDesktopCreated(IVirtualDesktop pDesktop)
			{
                virtualDesktops.Created?.Invoke(this, virtualDesktops.FromComObject(pDesktop));
			}

			void IVirtualDesktopNotification.VirtualDesktopDestroyBegin(IVirtualDesktop pDesktopDestroyed, IVirtualDesktop pDesktopFallback)
			{
				var args = new VirtualDesktopDestroyEventArgs(virtualDesktops.FromComObject(pDesktopDestroyed), virtualDesktops.FromComObject(pDesktopFallback));
                virtualDesktops.DestroyBegin?.Invoke(this, args);
			}

			void IVirtualDesktopNotification.VirtualDesktopDestroyFailed(IVirtualDesktop pDesktopDestroyed, IVirtualDesktop pDesktopFallback)
			{
				var args = new VirtualDesktopDestroyEventArgs(virtualDesktops.FromComObject(pDesktopDestroyed), virtualDesktops.FromComObject(pDesktopFallback));
                virtualDesktops.DestroyFailed?.Invoke(this, args);
			}

			void IVirtualDesktopNotification.VirtualDesktopDestroyed(IVirtualDesktop pDesktopDestroyed, IVirtualDesktop pDesktopFallback)
			{
				var args = new VirtualDesktopDestroyEventArgs(virtualDesktops.FromComObject(pDesktopDestroyed), virtualDesktops.FromComObject(pDesktopFallback));
                virtualDesktops.Destroyed?.Invoke(this, args);
			}

			void IVirtualDesktopNotification.ViewVirtualDesktopChanged(IntPtr pView)
			{
                virtualDesktops.ApplicationViewChanged?.Invoke(this, EventArgs.Empty);
			}

			void IVirtualDesktopNotification.CurrentVirtualDesktopChanged(IVirtualDesktop pDesktopOld, IVirtualDesktop pDesktopNew)
			{
				var args = new VirtualDesktopChangedEventArgs(virtualDesktops.FromComObject(pDesktopOld), virtualDesktops.FromComObject(pDesktopNew));
                virtualDesktops.CurrentChanged?.Invoke(this, args);
			}
		}
	}
}
