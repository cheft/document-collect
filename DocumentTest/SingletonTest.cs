using DocumentParser.builder;

namespace DocumentTest
{
   
    public class SingletonTest
    {
        private static readonly object padlock = new object();

        public string content = "1111111111111111";

        public void ThreadInsert()
        {
            lock (padlock)
            {
                DocumentBuilder builder = new DocumentBuilder();
                builder.Open(@"D:\office-test\doc\1.docx");
                for (int i = 0; i < 10; i++)
                {
                    builder.InsertContent(content);
                }
                builder.Save();
                builder.Quit();
            }
        }
    }
}
