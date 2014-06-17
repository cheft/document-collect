using log4net;

namespace TestNet4J
{
    public class Bar
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Bar));

        public void DoIt(string str)
        {
            log.Debug("Did it again!  %str");
        }
    }
}
