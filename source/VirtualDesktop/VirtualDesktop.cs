using System;
using System.ComponentModel;
using System.Diagnostics;
using WindowsDesktop.Interop;

namespace WindowsDesktop
{
    /// <summary>
    /// Encapsulates a virtual desktop on Windows 10.
    /// </summary>
    [DebuggerDisplay("{Id}")]
	public partial class VirtualDesktop
	{
		/// <summary>
		/// Gets the unique identifier for the virtual desktop.
		/// </summary>
		public Guid Id { get { return ComObject.GetID(); } }

		[EditorBrowsable(EditorBrowsableState.Never)]
		public IVirtualDesktop ComObject { get; }

        private readonly VirtualDesktops _desktops;

		internal VirtualDesktop(VirtualDesktops desktops, IVirtualDesktop comObject)
		{
            _desktops = desktops;
            this.ComObject = comObject;
		}


        /// <summary>
        /// Display the virtual desktop.
        /// </summary>
        public void Switch() => _desktops.Switch(this);

        /// <summary>
        /// Remove the virtual desktop, specifying a virtual desktop that display after destroyed.
        /// </summary>
        public void Remove(VirtualDesktop fallbackDesktop = null) => _desktops.Remove(this, fallbackDesktop);

        /// <summary>
        /// Returns a virtual desktop on the left.
        /// </summary>
        public VirtualDesktop Left => _desktops.GetLeft(this);

		/// <summary>
		/// Returns a virtual desktop on the right.
		/// </summary>
		public VirtualDesktop Right => _desktops.GetRight(this);
    }
}
