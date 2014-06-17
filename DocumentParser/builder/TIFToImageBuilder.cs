using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace DocumentParser.builder
{
    public class TIFToImageBuilder
    {

        public void TIFToImage(string source, string destPath)
        {
            FileInfo fi = new FileInfo(source);
            string fileName = fi.Name.Replace(fi.Extension, "");
            Image img = Image.FromFile(source);
            Guid guid = (Guid)img.FrameDimensionsList.GetValue(0);
            FrameDimension dimension = new FrameDimension(guid);
            int totalPage = img.GetFrameCount(dimension);

            for (int i = 0; i < totalPage; i++)
            {
                img.SelectActiveFrame(dimension, i);
                img.Save(destPath + "\\" + fileName + i + ".gif", System.Drawing.Imaging.ImageFormat.Gif);
            }
            img.Dispose();
        }


        public void TIFToImage(string source)
        {
            string destPath = source.Substring(0, source.LastIndexOf("."));
            if (!Directory.Exists(destPath))
            {
                Directory.CreateDirectory(destPath);
            }
            TIFToImage(source, destPath);
        }

        public void ImageToGif(string source)
        {
            Image img = Image.FromFile(source);
            img.Save(source.Substring(0, source.LastIndexOf(".")) + ".gif", System.Drawing.Imaging.ImageFormat.Gif);
            img.Dispose();
        }
    }
}
