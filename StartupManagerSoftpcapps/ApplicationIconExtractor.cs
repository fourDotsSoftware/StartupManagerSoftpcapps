using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;

namespace StartupManagerSoftpcapps
{
    public class ApplicationIconExtractor
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public IntPtr iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };

        public const uint SHGFI_ICON = 0x100;
        public const uint SHGFI_LARGEICON = 0x0;
        public const uint SHGFI_SMALLICON = 0x1;

        [DllImport("shell32.dll")]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        IntPtr hImgSmall;
        SHFILEINFO shinfo = new SHFILEINFO();

        private string _Filepath = "";
        public ApplicationIconExtractor(string filepath)
        {
            _Filepath = filepath;

            if (!System.IO.File.Exists(filepath))
            {
                throw new Exception("File does not exist !");
            }
        }

        public Bitmap Icon
        {
            get
            {
                try
                {
                    hImgSmall = SHGetFileInfo(_Filepath, 0, ref shinfo, (uint)Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_SMALLICON);

                    Icon ico = System.Drawing.Icon.FromHandle(shinfo.hIcon);

                    return ico.ToBitmap();
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
