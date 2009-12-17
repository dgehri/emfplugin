using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace EMFPlugIn
{

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    [ProgId("EmfPlugIn.ThumbnailProvider"), Guid("97E9072F-057F-4678-BBFC-B01249130A50")]
    public class ThumbnailProvider : IThumbnailProvider, IInitializeWithStream
    {
        #region IInitializeWithStream

        private IStream BaseStream { get; set; }

        public void Initialize(IStream stream, int grfMode)
        {
            this.BaseStream = stream;
        }

        #endregion

        #region IThumbnailProvider

        public void GetThumbnail(int cx, out IntPtr hBitmap, out WTS_ALPHATYPE bitmapType)
        {

            hBitmap = IntPtr.Zero;
            bitmapType = WTS_ALPHATYPE.WTSAT_UNKNOWN;

            try
            {
                using (System.Drawing.Image image = System.Drawing.Image.FromStream(GetStreamContents()))
                {
                    int width, height;
                    if (image.Width > image.Height)
                    {
                        width = cx;
                        height = image.Height * cx / image.Width;
                    }
                    else
                    {
                        height = cx;
                        width = image.Width * cx / image.Height;
                    }

                    using (Image thumb = image.GetThumbnailImage(width, height, abortCallback, IntPtr.Zero))
                    {
                        using (Bitmap bmp = new Bitmap(thumb))
                        {
                            hBitmap = bmp.GetHbitmap(Color.White);
                        }
                    }
                }

            }
            catch { } // A dirty cop-out.

        }

        System.Drawing.Image.GetThumbnailImageAbort abortCallback = null;

        #endregion

        private Stream GetStreamContents()
        {

            if (this.BaseStream == null) return null;

            System.Runtime.InteropServices.ComTypes.STATSTG statData;
            this.BaseStream.Stat(out statData, 1);

            byte[] buf = new byte[statData.cbSize];

            IntPtr P = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(UInt64)));
            try
            {
                this.BaseStream.Read(buf, buf.Length, P);
            }
            finally
            {
                Marshal.FreeCoTaskMem(P);
            }

            return new MemoryStream(buf);
        }
    }
}