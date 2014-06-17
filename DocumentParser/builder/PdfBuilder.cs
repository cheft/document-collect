/***************************************************************************
 * 说明：PDF文档构建类
 * 作者：陈海峰
 * 日期：2013-04-26
 *  * 注意事项：
 * 1、依赖于Adobe Acrobat Pro
 ****************************************************************************/
using log4net;
using System;
using System.IO;
using System.Reflection;

namespace DocumentParser.builder
{
    public class PdfBuilder
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(PdfBuilder));

        #region 多个PDF文档合并，需引用iTextSharp.dll,弃用
        /*
        public void MergePdf(string[] list, string dest)
        {
            PdfReader reader;
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(dest, FileMode.OpenOrCreate));
            document.Open();
            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage newPage;
            for (int i = 0; i < list.Length; i++)
            {
                reader = new PdfReader(list[i]);
                int iPageNum = reader.NumberOfPages;
                for (int j = 1; j <= iPageNum; j++)
                {
                    document.NewPage();
                    newPage = writer.GetImportedPage(reader, j);
                    cb.AddTemplate(newPage, 0, 0);
                }
            }
            document.Close(); 
               
        }
         * */
        #endregion

        #region PDF转换其它常用格式，需引用 Acrobat，因为组件收费弃用
        /// <summary>
        /// 支持格式有 doc, docx, xls, xlsx, ppt, pptx, rtf, png
        /// html(导出png 图片颜色有问题)
        /// </summary>
        /// <param name="inPath">PDF输入文件</param>
        /// <param name="outPath">输出文件，格式以后缀区分</param>
        /*
        public void PdfConvertor(string inPath, string outPath)
        {
            AcroPDDoc pdDocument = new AcroPDDocClass();
            if (pdDocument != null)
            {
                try
                {
                    bool result = pdDocument.Open(inPath);
                    if (result)
                    {
                        string ext = Path.GetExtension(outPath);
                        string acrobatType = "com.adobe.acrobat" + ext;
                        object jsObject = pdDocument.GetJSObject();
                        Type type = jsObject.GetType();
                        type.InvokeMember("saveAs", BindingFlags.InvokeMethod, null, jsObject,
                            new object[] { outPath, acrobatType });
                    }
                }
                catch (Exception e)
                {
                    log.ErrorFormat("PDF {0} 转换出错，异常信息: {1}", inPath, e.Message);
                }
                finally
                {
                    pdDocument.Close();
                    pdDocument = null;
                }
            }
        }
        */
        #endregion

    }
}
