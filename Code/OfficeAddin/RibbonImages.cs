//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace SystemBannerOffice
{
    //button images drawn by the add-in, for the ribbon's getImage callback: the same on every Office version,
    //where an imageMso name may not exist (Office 2007 has fewer). Returned as an OLE icon picture: its mask keeps
    //the background transparent, which Office 2007 loses with a 32 bit bitmap.
    public static class RibbonImages
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct PictDescIcon
        {
            public int Size;
            public int Type;
            public IntPtr Icon;
            public IntPtr Unused;
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void OleCreatePictureIndirect(ref PictDescIcon description, ref Guid iid, bool own, [MarshalAs(UnmanagedType.IDispatch)] out object picture);

        private const int PictypeIcon = 3;

        //a document with a red arrow pointing down: lower the classification
        public static Bitmap Downgrade(int size)
        {
            Bitmap image = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(image))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float s = size / 32f;
                //the page, its corner folded
                PointF[] page = new PointF[] { P(5, 2, s), P(19, 2, s), P(25, 8, s), P(25, 30, s), P(5, 30, s) };
                using (Brush paper = new SolidBrush(Color.White))
                using (Pen edge = new Pen(Color.FromArgb(90, 90, 90), Math.Max(1f, s)))
                {
                    g.FillPolygon(paper, page);
                    g.DrawPolygon(edge, page);
                    g.DrawLines(edge, new PointF[] { P(19, 2, s), P(19, 8, s), P(25, 8, s) });
                }
                //the classification band at the top of the page
                using (Brush band = new SolidBrush(Color.FromArgb(200, 16, 46)))
                {
                    g.FillRectangle(band, 7 * s, 10 * s, 11 * s, 3 * s);
                }
                //the arrow down
                PointF[] arrow = new PointF[] { P(19, 13, s), P(25, 13, s), P(25, 21, s), P(30, 21, s), P(22, 30, s), P(14, 21, s), P(19, 21, s) };
                using (Brush fill = new SolidBrush(Color.FromArgb(200, 16, 46)))
                using (Pen outline = new Pen(Color.White, Math.Max(1f, s)))
                {
                    g.FillPolygon(fill, arrow);
                    g.DrawPolygon(outline, arrow);
                }
            }
            return image;
        }

        private static PointF P(float x, float y, float scale)
        {
            return new PointF(x * scale, y * scale);
        }

        //an IPictureDisp Office can show; the picture owns the icon handle
        public static object Picture(Bitmap image)
        {
            PictDescIcon description = new PictDescIcon();
            description.Size = Marshal.SizeOf(typeof(PictDescIcon));
            description.Type = PictypeIcon;
            description.Icon = image.GetHicon();
            Guid iPictureDisp = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
            object picture;
            OleCreatePictureIndirect(ref description, ref iPictureDisp, true, out picture);
            return picture;
        }
    }
}
