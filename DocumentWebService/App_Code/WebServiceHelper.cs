using DocumentParser.builder;
using DocumentParser.helper;
using log4net;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Timers;

/// <summary>
/// QueueHelper 的摘要说明
/// </summary>
/// 
namespace DocumentWeb
{

    public class WebServiceHelper
    {

        private static string queueDir;

        private static readonly ILog log = LogManager.GetLogger(typeof(WebServiceHelper));

        public WebServiceHelper()
        {
            queueDir = ConfigurationManager.AppSettings["QueueDir"];
  
        }

        public void GoIntoQueue(string filename, byte[] s)
        {
            string filePath = queueDir + "\\temp" + "\\" + filename;
            IOHelper.WriteFile(s, filePath);
            File.Move(filePath, queueDir + "\\request" + "\\" + filename);
        }

        public void GoIntoQueueOld(RequestData data)
        {
            string requestDir = queueDir + "\\request" + "\\" + data.DocName;

            if (!Directory.Exists(requestDir))
            {
                Directory.CreateDirectory(requestDir);
            }
            string requestFile = requestDir + "\\" + data.Sequence;

            SerializableHelper helper = new SerializableHelper();
            helper.Serialize(data, requestFile);
        }

        public byte[] GetDocument(string tag, bool force)
        {
            string requestFile = queueDir + "\\request\\"+ tag + ".zip";
            string documentDir = queueDir + "\\document\\";
            if (!force && File.Exists(requestFile))
            {
                return new byte[0];  
            }
            string docPath = documentDir + tag + ".doc";
            byte[] b = IOHelper.ReadFile(docPath);
            return b;
        }

    }

}