/***************************************************************************
 * 说明：文档转换类
 * 作者：陈海峰
 * 日期：2013-04-26
 ****************************************************************************/
using DocumentParser.builder;
using DocumentParser.helper;
using System;
using System.IO;

namespace DocumentParser
{
    public class DocumentConvertor
    {
        #region 文档转换
        public void Convert(string zipPath, string outPath)
        {
            DocumentBuilder office = new DocumentBuilder();

            string tempDir = Environment.GetEnvironmentVariable("TEMP");
            string path = tempDir + "\\" + Guid.NewGuid().ToString();
            ZipHelper.UnZip(zipPath, path);

            // string[] filesArray = Directory.GetFileSystemEntries(path);
            StreamReader sr = File.OpenText(path + "\\list.txt");
            string s = null;
            while ((s = sr.ReadLine()) != null)
            {
                office.InsertObject(path + "\\" + s);
            }
            sr.Close();
            office.SaveAs(outPath);
            Directory.Delete(path, true);
        }
        #endregion
    }
}
