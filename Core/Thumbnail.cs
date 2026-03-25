using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Ion.Tool.Rename;

public static class Thumbnail
{
    /// <see cref="private"/>
    #region

    private const int FILE_ATTRIBUTE_NORMAL = 0x80;
    private const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    private const uint SHGFI_ICON = 0x000000100;
    private const uint SHGFI_SYSICONINDEX = 0x000004000;
    private const uint SHGFI_LINKOVERLAY = 0x000008000;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [ComImport, Guid("46EB5926-582E-4017-9FDF-E8998DAA0950"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IImageList
    {
        [PreserveSig] int Add(IntPtr hbmImage, IntPtr hbmMask, ref int pi);
        [PreserveSig] int ReplaceIcon(int i, IntPtr hicon, ref int pi);
        [PreserveSig] int SetOverlayImage(int iImage, int iOverlay);
        [PreserveSig] int Replace(int i, IntPtr hbmImage, IntPtr hbmMask);
        [PreserveSig] int AddMasked(IntPtr hbmImage, int crMask, ref int pi);
        [PreserveSig] int Draw(IntPtr pimldp); /// Simplified
        [PreserveSig] int Remove(int i);
        [PreserveSig] int GetIcon(int i, int flags, ref IntPtr picon);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(string pszPath, int dwFileAttributes, ref SHFILEINFO psfi, int cbFileInfo, uint uFlags);

    [DllImport("shell32.dll", EntryPoint = "#727")]
    private static extern int SHGetImageList(int iImageList, ref Guid riid, out IImageList ppv);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static BitmapSource Clean(IntPtr hIcon)
    {
        try
        {
            var bs = Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bs.Freeze(); /// Crucial for multi-threading and preventing leaks
            return bs;
        }
        finally
        {
            DestroyIcon(hIcon);
        }
    }

    private static BitmapSource? GetShell(string filePath, int sizeIndex, bool checkDisk)
    {
        try
        {
            SHFILEINFO shinfo = new();
            uint flags = SHGFI_SYSICONINDEX;
            if (!checkDisk) flags |= SHGFI_USEFILEATTRIBUTES;

            int attribute = checkDisk ? 0 : FILE_ATTRIBUTE_NORMAL;
            SHGetFileInfo(filePath, attribute, ref shinfo, Marshal.SizeOf(shinfo), flags);

            Guid iidImageList = new("46EB5926-582E-4017-9FDF-E8998DAA0950");
            if (SHGetImageList(sizeIndex, ref iidImageList, out IImageList iml) == 0)
            {
                IntPtr hIcon = IntPtr.Zero;
                if (iml.GetIcon(shinfo.iIcon, 1, ref hIcon) == 0) /// 1 = ILD_TRANSPARENT
                {
                    return Clean(hIcon);
                }
            }
        }
        catch { /* Log error if necessary */ }
        return null;
    }

    private static BitmapSource? GetStandard(string filePath, bool small, bool checkDisk, bool addOverlay)
    {
        SHFILEINFO shinfo = new();
        uint flags = SHGFI_ICON | (small ? (uint)0x1 : (uint)0x0);

        if (!checkDisk) flags |= SHGFI_USEFILEATTRIBUTES;
        if (addOverlay) flags |= SHGFI_LINKOVERLAY;

        int attribute = checkDisk ? 0 : FILE_ATTRIBUTE_NORMAL;

        var res = SHGetFileInfo(filePath, attribute, ref shinfo, Marshal.SizeOf(shinfo), flags);
        if (res == 0 || shinfo.hIcon == IntPtr.Zero) return null;

        return Clean(shinfo.hIcon);
    }

    #endregion

    /// <see cref="public"/>
    #region

    public static ImageSource GetIcon(string filePath, ThumbnailSize size, bool checkDisk = true, bool addOverlay = false)
    {
        /// For Small/Large, we use SHGetFileInfo. For XL/Jumbo, we use SHGetImageList.
        if (size == ThumbnailSize.Small || size == ThumbnailSize.Large)
        {
            return GetStandard(filePath, size == ThumbnailSize.Small, checkDisk, addOverlay);
        }
        return GetShell(filePath, (int)size, checkDisk);
    }

    public static ImageSource GetImage(string path, int decodeHeight = 256)
    {
        try
        {
            /// Use DelayCreation and OnLoad to avoid locking the file or reading the whole thing into memory
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var frame = BitmapFrame.Create(stream, BitmapCreateOptions.DelayCreation, BitmapCacheOption.OnLoad);

            BitmapSource result = frame.Thumbnail;

            if (result == null)
            {
                /// Fallback: Create a scaled version of the main image
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.StreamSource = stream;
                bi.DecodePixelHeight = decodeHeight;
                bi.EndInit();
                result = bi;
            }

            /// Handle Orientation
            if (frame.Metadata is BitmapMetadata meta)
            {
                var query = meta.GetQuery("/app1/ifd/{ushort=274}");
                if (query != null)
                {
                    double angle = (ushort)query switch
                    {
                        3 => 180,
                        6 => 90,
                        8 => -90,
                        _ => 0
                    };
                    if (angle != 0)
                    {
                        result = new TransformedBitmap(result, new RotateTransform(angle));
                    }
                }
            }

            if (result.CanFreeze) result.Freeze();
            return result;
        }
        catch { return null; }
    }

    #endregion
}

public enum ThumbnailSize
{
    Small = 0,      /// 16x16
    Large = 1,      /// 32x32
    ExtraLarge = 2, /// 48x48
    Jumbo = 4       /// 256x256
}