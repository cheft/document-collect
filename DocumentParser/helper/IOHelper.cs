using log4net;
using System;
using System.IO;

namespace DocumentParser.helper
{
    public class IOHelper
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(IOHelper));

        #region File To byte
        public static byte[] ReadFile(string fileName)
        {
            FileStream pFileStream = null;
            byte[] pReadByte = new byte[0];
            try
            {
                pFileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                BinaryReader r = new BinaryReader(pFileStream);
                r.BaseStream.Seek(0, SeekOrigin.Begin);                 //将文件指针设置到文件开
                pReadByte = r.ReadBytes((int)r.BaseStream.Length);
            }catch(Exception ex) 
            {
                log.ErrorFormat("读取文件 {0} 出错, 异常信息：{1}", fileName, ex.Message);
            }
            finally
            {
                if (pFileStream != null)
                {
                    pFileStream.Close();
                }
                    
            }
            return pReadByte;
        }
        #endregion

        #region byte To File
        public static bool WriteFile(byte[] pReadByte, string fileName)
        {
            FileStream pFileStream = null;
            try
            {
                pFileStream = new FileStream(fileName, FileMode.OpenOrCreate);
                pFileStream.Write(pReadByte, 0, pReadByte.Length);
            }catch(Exception ex)
            {
                log.ErrorFormat("写入文件 {0} 出错, 异常信息：{1}", fileName, ex.Message);
            }
            finally
            {
                if (pFileStream != null)
                {
                    pFileStream.Close();
                }
            }
            return true;
        }
        #endregion
    }
}
