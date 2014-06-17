using DocumentParser;
using DocumentParser.builder;
using DocumentParser.helper;
using DocumentWeb;
using log4net;
using log4net.Config;
using System;
using System.Configuration;
using System.IO;
using System.Text;
using System.Web.Services;

/// <summary>
/// DocumentWebServcie 的摘要说明
/// </summary>
[WebService(Namespace = "http://localhost:16659/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
// [System.Web.Script.Services.ScriptService]
public class DocumentWebServcie : System.Web.Services.WebService {

    private static readonly ILog log = LogManager.GetLogger(typeof(DocumentWebServcie));

    public DocumentWebServcie () {

    }

    [WebMethod]
    public void Collect(string filename, string s)
    {
        byte[] b = Convert.FromBase64String(s); 
        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueue(filename, b);
    }

    [WebMethod]
    public string GetDocument(string docName)
    {
        log.DebugFormat("Invoke GetDocument API by {0}", docName);
        WebServiceHelper queueHelper = new WebServiceHelper();
        byte[] b = queueHelper.GetDocument(docName, false);
        return Convert.ToBase64String(b);
    }


    public void InsertTitle(int queueId, string tag, string text) 
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertTitle";
        data.Sequence = queueId;
        data.DocName = tag;
        data.StringParam = text;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertTitle API by {0} - {1}", queueId, tag);
    }

    public void InsertContent(int queueId, string tag, string text)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertContent";
        data.Sequence = queueId;
        data.DocName = tag;
        data.StringParam = text;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertContent API by {0} - {1}", queueId, tag);

    }


    public void InsertHeading(int queueId, string tag, string text, int type)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertHeading";
        data.DocName = tag;
        data.Sequence = queueId;
        data.StringParam = text;
        data.IntParam = type;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertHeading API by {0} - {1}", queueId, tag);

    }

    public void InsertObject(int queueId, string tag, string filename, byte[] s)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertObject";
        data.DocName = tag;
        data.Sequence = queueId;
        data.StringParam = filename;
        // data.ObjectParam = s;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertObject API by {0} - {1}", queueId, tag);

    }

    public void InsertWord(int queueId, string tag, string filename, byte[] s, int indent)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertWord";
        data.DocName = tag;
        data.Sequence = queueId;
        data.StringParam = filename;
        // data.ObjectParam = s;
        data.Indent = indent;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertWord API by {0} - {1}", queueId, tag);

    }

    public void InsertPagebreak(int queueId, string tag)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertPagebreak";
        data.DocName = tag;
        data.Sequence = queueId;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertPagebreak API by {0} - {1}", queueId, tag);

      }

    public void InsertLineBreak(int queueId, string tag, int line)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertLineBreak";
        data.DocName = tag;
        data.Sequence = queueId;
        data.IntParam = line;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertLineBreak API by {0} - {1}", queueId, tag);
    }

    public void InsertIndex(int queueId, string tag)
    {
        RequestData data = new RequestData();
        data.ServiceName = "InsertIndex";
        data.DocName = tag;
        data.Sequence = queueId;

        WebServiceHelper queueHelper = new WebServiceHelper();
        queueHelper.GoIntoQueueOld(data);

        log.DebugFormat("Invoke InsertIndex API by {0} - {1}", queueId, tag);
    }


}
