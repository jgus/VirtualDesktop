using System;
using System.Collections;
using System.Collections.Generic;
using WindowsDesktop.Interop;

namespace WindowsDesktop
{
    public partial class VirtualDesktops : IReadOnlyList<VirtualDesktop>
    {
        private IObjectArray Objects => _comObjects.VirtualDesktopManagerInternal.GetDesktops();

        public VirtualDesktop this[int index]
        {
            get
            {
                if (index < 0 || Count <= index)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }
                object ppvObject;
                Objects.GetAt((uint) index, typeof(IVirtualDesktop).GUID, out ppvObject);
                return wrap((IVirtualDesktop)ppvObject);
            }
        }

        public int Count => (int)Objects.GetCount();

        public IEnumerator<VirtualDesktop> GetEnumerator()
        {
            var count = Count;
            for (uint i = 0; i < count; i++)
            {
                object ppvObject;
                Objects.GetAt(i, typeof(IVirtualDesktop).GUID, out ppvObject);
                yield return wrap((IVirtualDesktop)ppvObject);
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
