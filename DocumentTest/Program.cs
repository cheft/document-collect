using DocumentParser.builder;
using DocumentParser.helper;
using System;
using System.Threading;
using log4net;
using System.Reflection;
using System.IO;
using System.Text;


namespace DocumentTest
{
    public class Program
    {

        private string k = "1";

        static void Main(string[] args)
        {
            /*
            DocumentConvertor convertor = new DocumentConvertor();
            string inPath = @"D:\office-test\test2.zip";
            string outPath = @"D:\office-test\test.docx";
            convertor.Convert(inPath, outPath);
            byte[] b = IOHelper.ReadFile(outPath);
            Console.WriteLine(b.Length);
            Console.ReadKey();
            */

            // insertWord();

            OfficeBuilder ob = new OfficeBuilder();
            // ob.Excel2PDF(@"D:\test.doc", @"D:\test.pdf");
            // ob.Word2PDF(@"D:\office-test\doc\test.doc", @"D:\office-test\doc\test");

            // ob.PPT2Image(@"D:\office-test\ppt\test.ppt", @"D:\office-test\ppt\test2");

            // ob.Excel2PDF(@"D:\office-test\xls\all.xls", @"D:\office-test\xls\all.pdf");
            // ob.Word2PDF(@"D:\office-test\doc\2.docx", @"D:\office-test\doc\2.pdf");

            PdfToImageBuilder pdfBuilder = new PdfToImageBuilder();
            pdfBuilder.PDFToImage(@"C:\Resources\Develop\collect\1602.pdf");
            TIFToImageBuilder tifBuilder = new TIFToImageBuilder();
            tifBuilder.TIFToImage(@"C:\Resources\Develop\collect\TIF\取证表.tif");

            FileInfo fi = new FileInfo(@"C:\Resources\Develop\collect\TIF\取证表.tif");

            ImageBuilder iBuilder = new ImageBuilder();
            iBuilder.Compress(@"C:\Resources\Develop\collect\test.png");
            iBuilder.Compress(@"C:\Resources\Develop\collect\test.gif");
            // OfficeBuilder builder = new OfficeBuilder();
            // builder.SplitExcel(@"D:\office-test\xls\all.xls");

            /*
            RequestList list = new RequestList();

            RequestData d = new RequestData();
            d.DocName = "111";
            d.StringParam = "222";
            d.ServiceName = "333";
            d.SplitParam = "%test%2012-03-13,%test2%陈之冲";

            list.RequestData.Add(d);

            d = new RequestData();
            d.DocName = "222";
            d.StringParam = "333";
            d.ServiceName = "4444";
            d.SplitParam = "%test%2012-03-13,%test2%风贤宁";

            list.RequestData.Add(d);
            */
            // SerializableHelper.SerilizeXml(list, @"D:\request_net.xml");

            // RequestList data = SerializableHelper.DeserilizeXml(@"D:\abc.xml");
            // Console.WriteLine(data.RequestData.Count);
            /*
            DocumentBuilder builder = null;
            try
            {
                string filePath = "C:\\Resource\\Develop\\document-collect\\test\\temp3.doc";
                builder = new DocumentBuilder("C:\\Resource\\Develop\\document-collect\\test\\test3.doc");
                DocumentBuilder insertWord = new DocumentBuilder(filePath);
                insertWord.InsertFrontCover("%projectName%项目名称,%auditPeriod%组织名称,%hr%人事,%merber%审计组员,"
                    + "%count%总 卷 数,%storageLife%保存年限,%creatName%录 入 人,%creatDate%录入时间");
                insertWord.Save();
                insertWord.Quit();
                // insertWord.KillWordProcess();
                builder.InsertObject(filePath);
                builder.Save();
                builder.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                builder.KillWordProcess();
            }
            */
           
            Console.ReadKey();
        }

        private static String InsertTxtFile(String path)
        {
            String txt = "";
            using (StreamReader sr = new StreamReader(path, Encoding.Default))
            {
                string s = "";
                while ((s = sr.ReadLine()) != null)
                {
                    txt += s + "\n";
                }
            }
            return txt;
        }

        public void TestConvert(string path)
        {
            string tempDir = Environment.GetEnvironmentVariable("TEMP");
            PdfToImageBuilder convertor = new PdfToImageBuilder();
            string imageDir = tempDir + "\\" + Guid.NewGuid() + "\\";
            Console.WriteLine(imageDir);

            Directory.CreateDirectory(imageDir);
            convertor.Convert(path, imageDir, "image", PdfToImageBuilder.Definition.One);
            string[] images = Directory.GetFiles(imageDir);
            foreach (string img in images)
            {
                Console.WriteLine(img);
                File.Delete(img);
            }
            Directory.Delete(imageDir);
        }

        public void TestSortFiles()
        {
            string[] files = Directory.GetFiles(@"D:\office-test\queue\request\12345678");
            foreach (string f in files)
            {
                Console.WriteLine(f);
            }
            Console.ReadKey();
        }


        private void testThread()
        {
          /*
            Program test2 = new Program();
            test2.k = "8";
            Thread d = new Thread(new ThreadStart(test2.test));
            d.Start();
            */
            AutoResetEvent autoEvent = new AutoResetEvent(true);
            Program test3 = new Program();

            // 为定时器创建一个委托方法
            TimerCallback timerDelegate = new TimerCallback(test3.test2);
            Timer stateTimer = new Timer(timerDelegate, autoEvent, 1000, 1000);

            // autoEvent.WaitOne(5000, false);
            // stateTimer.Change(0, 500);

            Console.Write("start");


            /*        
            Program test1 = new Program();
            test1.k = "7";
            test1.test2(2);
            */
        }

        private void log()
        {
            //Application.Run(new MainForm());
            //创建日志记录组件实例
            ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            //记录错误日志
            log.Error("error", new Exception("发生了一个异常"));
            //记录严重错误
            log.Fatal("fatal", new Exception("发生了一个致命错误"));
            //记录一般信息
            log.Info("info");
            //记录调试信息
            log.Debug("debug");
            //记录警告信息
            log.Warn("warn");
            Console.WriteLine("日志记录完毕。");
            Console.Read();
        }


        private void test2(object state)
        {

            Console.Write("z");

        }



        public void test()
        {
            while (true)
            {
                Console.Write(k);
            }

        }


        public static void insertWord() {
            /*
           DocumentBuilder doc = new DocumentBuilder();
            
           doc.InsertTitle("ttttttttttt");
           doc.InsertHeading1("111111111");
           doc.InsertHeading2("222222222");
           doc.InsertContent("333333333333");
           doc.InsertHeading6("66666666666");


           doc.InsertObject(@"D:\office-test\doc\3.doc");

           doc.SaveAs(@"D:\insert.doc");

           
           DocumentBuilder doc2 = new DocumentBuilder(@"D:\insert.docx");
            
           doc2.InsertPageBreak();
           doc2.InsertHeading7("77777777777");
           doc2.BuildContentsIndex();

           doc2.Save();*/

           DocumentBuilder insertWord = new DocumentBuilder(@"D:\template.doc");
           insertWord.IndentHeading(2);
           insertWord.Save();

        }
    }
}
