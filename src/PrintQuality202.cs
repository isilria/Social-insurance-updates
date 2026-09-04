using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;

namespace InsurancePayrollValidator
{
    // Preserve the existing A4 layout. Four times the original pixel density is ~600 dpi.
    static class PrintQuality202
    {
        public static Bitmap CreateBitmap(int width,int height)
        {
            var bitmap=new Bitmap(width*4,height*4,PixelFormat.Format24bppRgb);
            // Fonts use points; keep the original drawing DPI and scale the whole drawing once.
            bitmap.SetResolution(96,96);
            return bitmap;
        }
        public static void WriteLossless(Bitmap bitmap,Stream output)
        {
            output.WriteByte(0x78);output.WriteByte(0x9c);
            uint a=1,b=0;
            var data=bitmap.LockBits(new Rectangle(0,0,bitmap.Width,bitmap.Height),ImageLockMode.ReadOnly,PixelFormat.Format24bppRgb);
            try{
                byte[] row=new byte[bitmap.Width*3];
                using(var deflate=new DeflateStream(output,CompressionLevel.Optimal,true)){
                    for(int y=0;y<bitmap.Height;y++){
                        Marshal.Copy(IntPtr.Add(data.Scan0,y*data.Stride),row,0,row.Length);
                        for(int x=0;x<row.Length;x+=3){byte red=row[x+2];row[x+2]=row[x];row[x]=red;}
                        for(int i=0;i<row.Length;i++){a=(a+row[i])%65521;b=(b+a)%65521;}
                        deflate.Write(row,0,row.Length);
                    }
                }
            }finally{bitmap.UnlockBits(data);}
            uint checksum=(b<<16)|a;
            output.WriteByte((byte)(checksum>>24));output.WriteByte((byte)(checksum>>16));output.WriteByte((byte)(checksum>>8));output.WriteByte((byte)checksum);
        }
    }
}
