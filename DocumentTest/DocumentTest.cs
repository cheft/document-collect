using DocumentParser.builder;

namespace DocumentTest
{
    public class DocumentTest
    {
        public string content = "1111111111111111";

        public void open()
        {
            DocumentBuilder builder = new DocumentBuilder();
            builder.Open(@"D:\office-test\doc\2.docx");
            builder.InsertContent("22222222222222222");
            builder.Save();
            builder.Quit();

            DocumentBuilder builder2 = new DocumentBuilder();
            builder2.Open(@"D:\office-test\doc\2.docx");
            builder2.InsertContent("33333333333333333");
            builder2.Save();
            builder2.Quit();
        }


        public void ThreadInsert()
        {
            object lockThis = new object();
            lock (lockThis)
            {
                DocumentBuilder builder = new DocumentBuilder();
                builder.Open(@"D:\office-test\doc\2.docx");
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
