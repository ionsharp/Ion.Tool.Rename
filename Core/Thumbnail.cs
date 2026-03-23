using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Ion.Tool.Rename;

/// <inheritdoc/>
public static class Thumbnail
{
    /// <inheritdoc/>
    private static ImageSource Get(string filePath, bool Small, bool checkDisk, bool addOverlay)
    {
        SHFILEINFO shinfo = new();

        uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
        uint SHGFI_LINKOVERLAY = 0x000008000;

        uint flags;
        if (Small)
        {
            flags = SHGFI_ICON | SHGFI_SMALLICON;
        }
        else
        {
            flags = SHGFI_ICON | SHGFI_LARGEICON;
        }
        if (!checkDisk)
            flags |= SHGFI_USEFILEATTRIBUTES;
        if (addOverlay)
            flags |= SHGFI_LINKOVERLAY;

        var res = SHGetFileInfo(filePath, 0, ref shinfo, Marshal.SizeOf(shinfo), flags);
        if (res == 0)
        {
            throw (new System.IO.FileNotFoundException());
        }

        var myIcon = System.Drawing.Icon.FromHandle(shinfo.hIcon);

        var bs = GetFrom(myIcon);
        myIcon.Dispose();
        bs.Freeze(); // importantissimo se no fa memory leak
        DestroyIcon(shinfo.hIcon);
        SendMessage(shinfo.hIcon, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        return bs;
    }

    /// <inheritdoc/>
    public static ImageSource GetFrom(System.Drawing.Icon i)
    {
        var bitmap = i.ToBitmap();
        var hBitmap = bitmap.GetHbitmap();

        var result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
        if (!DeleteObject(hBitmap))
            throw new Win32Exception();

        return result;
    }

    /// <inheritdoc/>
    public static ImageSource GetThumb(string path)
    {
        BitmapMetadata meta = null;

        double angle = 0;
        var orientation = ExifOrientations.None;
        BitmapSource result;
        try
        {
            //Attempt of creation of the thumbnail via Bitmap frame: very fast and very inexpensive memory!
            var frame = BitmapFrame.Create(new Uri(path), BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
            if (frame.Thumbnail is null) // failure, attempts with BitmapImage (slower and expensive in memory)
            {
                MemoryStream ms = new();
                byte[] bytes = File.ReadAllBytes(path);
                ms.Write(bytes, 0, bytes.Length);

                var image = new BitmapImage()
                {
                    CacheOption = BitmapCacheOption.None,
                    CreateOptions = BitmapCreateOptions.DelayCreation,
                    DecodePixelHeight = 256
                };

                image.BeginInit();
                image.StreamSource = ms;
                image.EndInit();

                if (image.CanFreeze)
                    image.Freeze(); //To avoid memory leak 

                result = image;
            }
            else
            {
                //Get the image meta
                meta = frame.Metadata as BitmapMetadata;
                result = frame.Thumbnail;
            }

            if ((meta != null) && (result != null)) //si on a des meta, tentative de récupération de l'orientation
            {
                if (meta.GetQuery("/app1/ifd/{ushort=274}") != null) orientation = (ExifOrientations)Enum.Parse(typeof(ExifOrientations), meta.GetQuery("/app1/ifd/{ushort=274}").ToString());
                switch (orientation)
                {
                    case ExifOrientations.Rotate90:
                        angle = -90;
                        break;
                    case ExifOrientations.Rotate180:
                        angle = 180;
                        break;
                    case ExifOrientations.Rotate270:
                        angle = 90;
                        break;
                }
                if (angle != 0) // we have to rotate the image
                {
                    result = new TransformedBitmap(result.Clone(), new RotateTransform(angle));
                    result.Freeze();
                }
            }
        }
        catch
        {
            return null;
        }
        return result;
    }

    /// <inheritdoc/>
    private static ImageSource GetLarge(string filePath, bool jumbo, bool checkDisk)
    {
        try
        {
            SHFILEINFO shinfo = new();

            uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
            uint SHGFI_SYSICONINDEX = 0x4000;

            int FILE_ATTRIBUTE_NORMAL = 0x80;

            uint flags;
            flags = SHGFI_SYSICONINDEX;

            if (!checkDisk)  // This does not seem to work. If I try it, a folder icon is always returned.
                flags |= SHGFI_USEFILEATTRIBUTES;

            var res = SHGetFileInfo(filePath, FILE_ATTRIBUTE_NORMAL, ref shinfo, Marshal.SizeOf(shinfo), flags);
            if (res == 0)
            {
                throw (new System.IO.FileNotFoundException());
            }
            var iconIndex = shinfo.iIcon;

            // Get the System IImageList object from the Shell:
            Guid iidImageList = new("46EB5926-582E-4017-9FDF-E8998DAA0950");

            int size = jumbo ? SHIL_JUMBO : SHIL_EXTRALARGE;
            var hres = SHGetImageList(size, ref iidImageList, out IImageList iml);
            // writes iml
            //if (hres == 0)
            //{
            //    throw (new System.Exception("Error SHGetImageList"));
            //}

            IntPtr hIcon = IntPtr.Zero;
            int ILD_TRANSPARENT = 1;
            hres = iml.GetIcon(iconIndex, ILD_TRANSPARENT, ref hIcon);
            //if (hres == 0)
            //{
            //    throw (new System.Exception("Error iml.GetIcon"));
            //}

            var myIcon = System.Drawing.Icon.FromHandle(hIcon);
            var bs = GetFrom(myIcon);
            myIcon.Dispose();
            bs.Freeze(); // very important to avoid memory leak
            DestroyIcon(hIcon);
            SendMessage(hIcon, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return bs;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public static ImageSource GetLarge(string filePath) => GetLarge(filePath, true, true);


    /// <see cref="DllImportAttribute"/>
    #region

    /// <inheritdoc/>
    private const int WM_CLOSE = 0x0010;

    /// <inheritdoc/>
    private const int SHGFI_ICON = 0x100;

    /// <inheritdoc/>
    private const int SHGFI_SMALLICON = 0x1;

    /// <inheritdoc/>
    private const int SHGFI_LARGEICON = 0x0;

    /// <inheritdoc/>
    private const int SHIL_JUMBO = 0x4;

    /// <inheritdoc/>
    private const int SHIL_EXTRALARGE = 0x2;


    /// <inheritdoc/>
    private struct Pair
    {
        public System.Drawing.Icon Icon { get; set; }

        public IntPtr HandleToDestroy { set; get; }
    }


    /// <inheritdoc/>
    [DllImport("user32")]
    internal static extern IntPtr SendMessage(IntPtr handle, int Msg, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// SHGetImageList is not exported correctly in XP.  See KB316931
    /// http://support.microsoft.com/default.aspx?scid=kb;EN-US;Q316931
    /// Apparently (and hopefully) ordinal 727 isn't going to change.
    /// </summary> 
    [DllImport("shell32.dll", EntryPoint = "#727")]
    internal static extern int SHGetImageList(int iImageList, ref Guid riid, out IImageList ppv);

    /// <summary>
    /// The signature of SHGetFileInfo (located in Shell32.dll)
    /// </summary>
    [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
    internal static extern int SHGetFileInfo(string pszPath, int dwFileAttributes, ref SHFILEINFO psfi, int cbFileInfo, uint uFlags);

    /// <inheritdoc/>
    [DllImport("Shell32.dll")]
    internal static extern int SHGetFileInfo(IntPtr pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, int cbFileInfo, uint uFlags);

    /// <inheritdoc/>
    [DllImport("shell32.dll", SetLastError = true)]
    internal static extern int SHGetSpecialFolderLocation(IntPtr hwndOwner, int nFolder, ref IntPtr ppidl);

    /// <inheritdoc/>
    [DllImport("user32")]
    internal static extern int DestroyIcon(IntPtr hIcon);

    /// <inheritdoc/>
    [DllImport("gdi32.dll")]
    internal static extern bool DeleteObject(IntPtr hObject);

    #endregion
}
