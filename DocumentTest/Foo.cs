using log4net;
using log4net.Config;
using System;

namespace TestNet4J
{
    class Foo
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Foo));

        static void test(string[] args)
        {
            // XmlConfigurator.Configure();

            XmlConfigurator.Configure(new System.IO.FileInfo(@"D:\topway\Projects\DocumentConvertor\TestNet4J\App.config"));

            log.Info("Entering application.");

            Bar bar = new Bar();
            bar.DoIt("head");

            log.Info("Exiting application.");

            string test = "Hello";
            Console.WriteLine(String.Format("{0}, World", test));
            Console.ReadKey();

        }
    }
}
