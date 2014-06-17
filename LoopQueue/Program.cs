using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;
using log4net;
using System.Configuration;
using DocumentParser.helper;
using System.Threading;
using DocumentParser.builder;
using Topway.Audit;



namespace LoopQueue
{
    class Program
    {
        static void Main(string[] args)
        {
            new AppExceptionHandler();
            Console.WriteLine("归档服务已经开启......" );
            QueueHelper helper =new QueueHelper();
            System.Threading.Thread d = new System.Threading.Thread(new System.Threading.ThreadStart(helper.LoopQueue));
            d.Start();
            Console.Read();
       
        }
        
    }
    public class QueueHelper
    {

        private static string queueDir;

        private static int queueInterval;

        private static int javaServerOn;

        private static string javaServerUrl;

        private static readonly object padlock = new object();

        private static readonly ILog log = LogManager.GetLogger(typeof(QueueHelper));

        public QueueHelper()
        {
            queueDir = ConfigurationManager.AppSettings["QueueDir"];
            queueInterval = int.Parse(ConfigurationManager.AppSettings["QueueInterval"]);
            javaServerOn = int.Parse(ConfigurationManager.AppSettings["javaServerOn"]);
            javaServerUrl = ConfigurationManager.AppSettings["javaServerUrl"];
        }


        public void LoopQueue()
        {
            while (true)
            {
                ExecuteQueue();
                Thread.Sleep(queueInterval);
            }
        }


        public DocumentBuilder InvokeFrontCover(RequestList list, string docPath)
        {
            RequestData d = null;
            foreach (RequestData data in list.RequestData)
            {
                if (data.ServiceName == "InsertFrontCover")
                {
                    string tempDir = queueDir + "\\temp\\";
                    string filePath = tempDir + data.DocName + "\\cover.doc";
                    File.Move(filePath, docPath);
                    d = data;
                    break;
                }
            }
            DocumentBuilder builder = new DocumentBuilder(docPath);
            if (d != null)
            {
                builder.InsertFrontCover(d.SplitParam);
            }
            return builder;
        }

        // public void ExecuteQueue(object sender, ElapsedEventArgs e)
        public void ExecuteQueue()
        {
            lock (padlock)
            {
                string docDir = queueDir + "\\document";
                string requestDir = queueDir + "\\request";
                string tempDir = queueDir + "\\temp";
                string[] files = Directory.GetFiles(requestDir);
             
                if (files.Length <= 0)
                {
                    return;
                }
                foreach (string f in files)
                {
                    Console.WriteLine("正在归档：" + f);
                    DocumentBuilder builder = null;
                    int tempIndex = f.LastIndexOf("\\") + 1;
                    string tag = f.Substring(tempIndex, f.LastIndexOf(".") - tempIndex);
                    try
                    {
                        ZipHelper.UnZip(f, tempDir + "\\" + tag);
                        FileFilter(tempDir + "\\" + tag);      
                        string docPath = docDir + "\\" + tag + ".doc";
                        string xmlPath = tempDir + "\\" + tag + "\\request.xml";
                        RequestList list = SerializableHelper.DeserilizeXml(xmlPath);
                        File.Delete(xmlPath);
                        builder = InvokeFrontCover(list, docPath);
                        // builder = new DocumentBuilder(docPath);

                        foreach (RequestData data in list.RequestData)
                        {
                            InvokeMethod(data, builder);
                        }
                        builder.Save();
                        builder.Quit();
                        Directory.Delete(tempDir + "\\" + tag, true);
                        SendConvertLog(tag, true, "");
                        
                        File.Delete(f);
                        
                        Console.WriteLine(f + " 归档成功");
                    }
                    catch (Exception ex)
                    {
                        SendConvertLog(tag, false, ex.Message);
                        File.Delete(f);
                        log.ErrorFormat("执行队列出错，异常信息： {0}", ex.Message);
                        Console.WriteLine(f + " 归档失败");
                        Console.WriteLine("执行队列出错，异常信息： {0}", ex.Message);
                        if (builder != null)
                        {
                            builder.Quit();
                            builder.KillWordProcess();
                        }
                    }
                }
            }
        }

        private void FileFilter(string dirPath)
        {
            PdfToImageBuilder pdfBuilder = new PdfToImageBuilder();
            TIFToImageBuilder tifBuilder = new TIFToImageBuilder();
            ImageBuilder iBuilder = new ImageBuilder();

            string[] files = Directory.GetFiles(dirPath);
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.Length == 0)
                {
                    continue;
                }
                switch (fi.Extension.ToLower()) {
                    case ".jpg":
                        iBuilder.Compress(file);
                        break;
                    case ".jepg":
                        iBuilder.Compress(file);
                        break;
                    case ".png":
                        iBuilder.Compress(file);
                        break;
                    case ".bmp":
                        iBuilder.Compress(file);
                        break;
                    case ".tif":
                        tifBuilder.TIFToImage(file);
                        break;
                        
                    case ".pdf":
                        pdfBuilder.PDFToImage(file);
                        break;
                         
                }
            }

        }


        public void SendConvertLog(string tag, bool isConverted, string msg)
        {
            if (javaServerOn == 0)
            {
                return;
            }
            string postData = "tag=" + tag;
            postData += "&isConverted=" + isConverted;
            postData += "&msg=" + msg;

            byte[] data = Encoding.UTF8.GetBytes(postData);

            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(javaServerUrl);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/x-www-form-urlencoded";

            myRequest.ContentLength = data.Length;
            Stream newStream = myRequest.GetRequestStream();
            // Send the data.
            newStream.Write(data, 0, data.Length);
            newStream.Close();
            // Get response
            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.Default);
            string content = reader.ReadToEnd();
            log.Debug(content);
        }


        /*
        public void ExecuteQueueOld(object sender, ElapsedEventArgs e)
        {
            lock (padlock)
            {
                string docDir = queueDir + "\\document\\";
                string requestDir = queueDir + "\\request";
                string[] dirs = Directory.GetDirectories(requestDir);
                SerializableHelper helper = new SerializableHelper();
                try
                {
                    foreach (string s in dirs)
                    {
                        string[] files = Directory.GetFiles(s);
                        string tag = s.Substring(s.LastIndexOf("\\") + 1);
                        DocumentBuilder builder = new DocumentBuilder(docDir + tag + ".doc");
                        foreach (string f in files)
                        {
                            RequestData data = helper.DeSerialize(f);
                            InvokeMethod(data, builder);
                            File.Delete(f);
                        }
                        builder.Save();
                        builder.Quit();
                        if (Directory.GetFiles(s).Length == 0)
                        {
                            Directory.Delete(s);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.ErrorFormat("执行队列出错，异常信息： {0}", ex.Message);
                }
            }
        }
        */
        public void InvokeMethod(RequestData data, DocumentBuilder builder)
        {
            try
            {
                if (data.ServiceName == "InsertTitle")
                {
                    builder.InsertTitle(data.StringParam);
                }
                else if (data.ServiceName == "InsertContent")
                {
                    builder.InsertContent(data.StringParam);
                }
                else if (data.ServiceName == "InsertHeading")
                {
                    builder.InsertHeading(data.IntParam, data.StringParam);
                }
                else if (data.ServiceName == "InsertObject")
                {
                    string tempDir = queueDir + "\\temp\\";
                    string filePath = tempDir + data.DocName + "\\" + data.Filename;
                    builder.InsertObject(filePath);
                    // File.Delete(filePath);
                }
                else if (data.ServiceName == "InsertWord")
                {
                    string tempDir = queueDir + "\\temp\\";
                    string filePath = tempDir + data.DocName + "\\" + data.Filename;
                    DocumentBuilder insertWord = new DocumentBuilder(filePath);
                    insertWord.IndentHeading(data.Indent);
                    insertWord.Save();
                    // insertWord.getWord().NormalTemplate.Saved = true;
                    insertWord.Quit();
                    // insertWord.KillWordProcess();
                    //builder.getWord().Activate();
                    builder.InsertObject(filePath);
                    //File.Delete(filePath);
                }
                else if (data.ServiceName == "InsertTemplate")
                {
                    string tempDir = queueDir + "\\temp\\";
                    string filePath = tempDir + data.DocName + "\\" + data.Filename;
                    builder.InsertObject(filePath);
                    builder.InsertFrontCover(data.SplitParam);
                    // File.Delete(filePath);
                }
                else if (data.ServiceName == "InsertPageBreak")
                {
                    builder.InsertPageBreak();
                }
                else if (data.ServiceName == "InsertLineBreak")
                {
                    builder.InsertLineBreak(data.IntParam);
                }
                else if (data.ServiceName == "InsertIndex")
                {
                    builder.BuildContentsIndex();
                }
                else if (data.ServiceName == "InsertBookmark")
                {
                    builder.InsertBookmark(data.StringParam);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(data.ServiceName + " - " + data.Sequence + "未归入，请检查", e.Message);
            }

        }

    }
}
