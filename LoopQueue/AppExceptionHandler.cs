using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Threading;


namespace Topway.Audit
{
    public class AppExceptionHandler
    {
        public AppExceptionHandler()
        {

            //Application.ThreadException += new ThreadExceptionEventHandler(this.OnThreadException);
            //Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        
            // Add the event handler for handling non-UI thread exceptions to the event:
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(UnhandledExceptionFunction);
        }
        /// <summary>
        /// 线程式异常处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnThreadException(object sender, ThreadExceptionEventArgs args)
        {
            try
            {
                string errorMsg =args.Exception.Message;
                string detailMsg = args.Exception.Message + "\r\n" + args.Exception.GetType() + "\r\n\r\n";

                detailMsg += args.Exception.StackTrace;


                Console.WriteLine(detailMsg);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        /// <summary>
        /// 未经处理的异常方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void UnhandledExceptionFunction(Object sender, UnhandledExceptionEventArgs args)
        {
            try
            {
                Exception exce = (Exception)args.ExceptionObject;

                if (exce == null)
                    exce = new Exception("未知错误。。");

                string errorMsg = exce.Message;
                string detailMsg = exce.Message + "\r\n\r\n";
                detailMsg += exce.ToString();

                Console.WriteLine(detailMsg);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
