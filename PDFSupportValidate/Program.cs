using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentParser.builder;
using DocumentParser.helper;
using System.Configuration;

namespace PDFSupportValidate
{
    public class Program
    {

        static void Main(string[] args)
        {
            try
            {
                string pdf = ConfigurationManager.AppSettings["pdf"];
                PdfToImageBuilder pdfBuilder = new PdfToImageBuilder();
                pdfBuilder.PDFToImage(pdf);
                Console.WriteLine("您的环境支持PDF转换");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("您的环境不支持PDF转换");
            }
        }
    }
}
