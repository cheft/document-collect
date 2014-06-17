package com.topway.doc.test;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.nutz.lang.Streams;

import com.topway.doc.api.DocumentClient;

public class ClientTest {

	public static void main(String[] args) throws IOException {
		/*
		test2003("test20");
		test2003("test21");
		test2003("test22");
		test2003("test23");
		test2003("test24");
		*/
		test2003("test64");

		// saveDocument("test17");
	}


	public static void saveDocument(String name) throws IOException {
		DocumentClient client = new DocumentClient(name);
		byte[]  b = client.getDocument();
		if(b.length > 0) {
			OutputStream stream = new FileOutputStream(new File("D:\\" + name + ".doc"));
			Streams.writeAndClose(stream, b);
		}else {
			System.out.println("正在转换，请稍后");
		}
	}

	public static void test2003(String name) {
		System.out.println("send request...");
		//DocumentClient client = new DocumentClient(name,
		//		"C:\\Resource\\Develop\\office-test\\doc\\cover.doc",
		//		"%projectName%文档管理系统,%creatName%成风,%creatDate%2012-09-08");
		DocumentClient client = new DocumentClient(name, "C:/cover.doc", "%projectName%文档管理系统,%creatName%成风,%creatDate%2012-09-08");
		client.insertPageBreak();
		// client.insertTemplate("C:\\Resource\\Develop\\office-test\\doc\\template.doc",
		//		"%projectName%文档管理系统,%auditPeriod%测试,%hr%组    长,%merber%审计组员,%count%7,%storageLife%1,%creatName%录 入 人,%creatDate%录入时间");
		client.insertPageBreak();
		client.insertTitle("目 录");
		client.insertIndexBookmark();
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
		// client.insertObject("C:\\Resource\\Develop\\office-test\\ppt\\表单设计器-V1.0.ppt");


		client.insertHeading("第三章", 1);
		client.insertHeading("Excel", 2);
		//client.insertObject("D:\\office-test\\xls\\all.xls");
		client.insertHeading("第四章", 1);
		client.insertHeading("Word", 2);
		//client.insertWord("D:\\office-test\\doc\\2.doc", 2);
		client.insertHeading("第五章", 1);
		client.insertHeading("PDF", 2);
		//client.insertObject("D:\\office-test\\pdf\\all.pdf");
		client.insertHeading("第六章", 1);
		client.insertHeading("RAR", 2);
		// client.insertsObject("D:\\office-test\\zip\\zip.rar");
		client.insertHeading("第七章", 1);
		client.insertHeading("ZIP", 2);
		// client.insertObject("D:\\office-test\\zip\\zip.zip");


		client.insertIndex();

		client.collect();
	}

	public static void testCollect() {
		DocumentClient client = new DocumentClient("中文文件名44", "D:\\office-test\\doc\\temp.docx", "%author%风,%date%2012-09-08");
		client.insertPageBreak();
		client.insertPageBreak();
		client.insertTitle("测试");
		client.insertLineBreak(3);
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
		client.insertObject("D:\\office-test\\ppt\\表单设计器 - V1.0.pptx");
		client.insertHeading("第三章", 1);
		client.insertHeading("Excel", 2);
		client.insertObject("D:\\office-test\\xls\\all.xls");
		client.insertHeading("第四章", 1);
		client.insertHeading("Word", 2);
		client.insertWord("D:\\office-test\\doc\\2.docx", 2);
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
}
