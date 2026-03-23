using System.Runtime.InteropServices;
using System.Windows;

namespace Ion.Tool.Rename;

[StructLayout(LayoutKind.Sequential)]
internal struct IMAGEINFO
{
    public IntPtr hbmImage;
    public IntPtr hbmMask;
    public int Unused1;
    public int Unused2;
    public Rect rcImage;
}