## 归档服务组件 JAVA Client端使用

### Get Start

#### 将给定的依赖jar引入项目中，在src目录下创建  _setting.properties_ 配置文件,修改 _serviceUrl_ 为ASP.NET Server端归档服务地址

    serviceUrl=http://172.18.80.79:12333/DocumentWebServcie.asmx

#### 在客户端创建 Servlet类,接受服务端转换完成后的请求，服务端会传回有三个参数    tag， isConverted，msg，修改ASP.NET Server端归档配置文件web.config 的JavaServerOn , JavaServerUrl 值

```xml
<appSettings>
    <!-- 省略 -->
	<add key="JavaServerOn" value="1"/>     <!--是否向客户端传回转换日志-->
    <add key="JavaServerUrl" value="http://172.18.97.60:8080/Project/ServletUrl"/>  <!--客户端传回转换日志请求地址-->
</appSettings>

</log4net>
```

```java
public class ServletTest extends HttpServlet {       
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String tag = request.getParameter("tag");    // 文档Id
		boolean isConverted = new Boolean(request.getParameter("isConverted")).booleanValue(); // 是否转换成功
		String msg = request.getParameter("msg");    // 转换错误信息
		//TODO 转换响应之后客户端的逻辑处理
	}
}
```

### Demo
	
```java
public static void test2003(String name) {
	DocumentClient client = new DocumentClient(name);
	client.insertPageBreak();
	client.insertFrontCover("D:\\office-test\\doc\\temp.doc",
			"%author%风,%date%2012-09-08");
	client.insertPageBreak();
	client.insertTitle("测试");
	client.insertLineBreak(2);
	client.insertHeading("第一章", 1);
	client.insertHeading("背景", 2);
	client.insertContent("正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正正"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文"
			+ "正文文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文正文");
	client.insertHeading("第二章", 1);
	client.insertHeading("PPT", 2);
	client.insertObject("D:\\office-test\\ppt\\表单设计器 - V1.0.ppt");
	client.insertHeading("第三章", 1);
	client.insertHeading("Excel", 2);
	client.insertObject("D:\\office-test\\xls\\all.xls");
	client.insertHeading("第四章", 1);
	client.insertHeading("Word", 2);
	client.insertWord("D:\\office-test\\doc\\2.doc", 2);
	client.insertHeading("第五章", 1);
	client.insertHeading("PDF", 2);
	client.insertObject("D:\\office-test\\pdf\\all.pdf");
	client.insertHeading("第六章", 1);
	client.insertHeading("RAR", 2);
	client.insertObject("D:\\office-test\\zip\\zip.rar");
	client.insertHeading("第七章", 1);
	client.insertHeading("ZIP", 2);
	client.insertObject("D:\\office-test\\zip\\zip.zip");
	client.insertIndex();
	client.collect();
}
```

******

```java
public static void saveDocument(String name) throws IOException {
	DocumentClient client = new DocumentClient(name);
	byte[]  b = client.getDocument();
	if(b.length > 0) {
		OutputStream stream = new FileOutputStream(new File("D:\\" + name + ".doc"));
		IOUtils.write(b, stream);
		stream.close();
	}else{
		System.out.println("正在转换，请稍后");
	}
}
```


### API
 
#### 构造方法详细信息

	DocumentClient
	public DocumentClient(java.lang.String docName)
	参数：docName - 惟一标识  

#### 方法详细信息

##### collect
	public void collect()
	归档 

##### getDocument
	public byte[] getDocument()
	获取归档文件 

	返回：文件流

##### insertFrontCover
	public void insertFrontCover(java.lang.String filePath,
	                             java.lang.String splitParam)
	插入封面 

	参数：filePath - 封面模板路径splitParam - 替换内容表达式


##### insertTitle
	public void insertTitle(java.lang.String text)
	插入标题 

	参数：text - 文本


##### insertContent
	public void insertContent(java.lang.String text)
	插入内容 

	参数：text - 文本

##### insertHeading
	public void insertHeading(java.lang.String text,
	                          int type)
	插入Heading 

	参数：text - 文本type - Heading级别，分别是1-9

##### insertObject
	public void insertObject(java.lang.String filePath)
	插入文件对象 

	参数：filePath - 对象文件路径


##### insertWord
	public void insertWord(java.lang.String filePath,
	                       int indent)
	插入需要缩进Heading的word对象 

	参数：filePath - word文件路径indent - 缩进级别

##### insertPageBreak
	public void insertPageBreak()
	插入分页 

##### insertLineBreak
	public void insertLineBreak(int line)
	插入空白行 

	参数：line - 行数

##### insertIndex
	public void insertIndex()
	插入目录 

