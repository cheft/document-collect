package com.topway.doc.api;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.CRC32;
import java.util.zip.CheckedOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipOutputStream;
import org.nutz.http.Request;
import org.nutz.http.Request.METHOD;
import org.nutz.http.Response;
import org.nutz.http.sender.PostSender;
import org.nutz.lang.Files;
import org.nutz.lang.Streams;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import com.topway.doc.entity.RequestData;
import com.topway.doc.entity.RequestList;
import com.topway.doc.util.PropertyUtil;
import com.topway.doc.util.XmlUtil;

/**
 *
 * 归档客户端API
 * @author ChenHaifeng
 *
 */
public class DocumentClient {

	private final List<RequestData> list = new ArrayList<RequestData>();

	private static int sequence = 100000000;

	private static final int BUFFER = 1024;

	private String docName = null;

	/**
	 * @param docName 惟一标识
	 */
	public DocumentClient(String docName) {
		this.docName = docName;
	}

	public RequestData getData(String serviceName) {
	    RequestData data = new RequestData();
	    data.setSequence(++sequence);
	    data.setDocName(docName);
        data.setServiceName(serviceName);
	    return data;
	}

	/**
	 * @param docName 惟一标识
	 * @param coverPath 封面路径
	 * @param converParam 封面参数
	 */
	public DocumentClient(String docName, String coverPath, String converParam) {
		this.docName = docName;

        RequestData data = this.getData("InsertFrontCover");
        data.setStringParam(coverPath);
        data.setSplitParam(converParam);
        list.add(data);
	}

	/**
	 * 插入封面
	 * @param filePath    封面模板路径
	 * @param splitParam  替换内容表达式
	 * 表达式示例：封面模板中有若有两处需要替换，分别由  %author% 、%date% 占位,
	 * splitParam 应为  "%author%张三,%date%2013-06-06",张三、2013-06-06为要替换的内容
	 */
    public void insertTemplate(String filePath, String splitParam)
    {
        RequestData data = this.getData("InsertTemplate");
        data.setStringParam(filePath);
        data.setFilename(filePath.substring(filePath.lastIndexOf("\\") + 1));
        data.setSplitParam(splitParam);
        list.add(data);
    }

	/**
	 * 插入标题
	 * @param text 文本
	 */
    public void insertTitle(String text) {
        RequestData data = this.getData("InsertTitle");
        data.setStringParam(text);
        list.add(data);
    }

    /**
     * 插入内容
     * @param text 文本
     */
    public void insertContent(String text) {
        RequestData data = this.getData("InsertContent");
        data.setStringParam(text);
        list.add(data);
    }

    /**
     * 插入Heading
     * @param text 文本
     * @param type Heading级别，分别是1-9
     */
    public void insertHeading(String text, int type) {
        RequestData data = this.getData("InsertHeading");
        data.setStringParam(text);
        data.setIntParam(type);
        list.add(data);
    }

    /**
     * 插入文件对象
     * @param filePath 对象文件路径
     */
    public void insertObject(String filePath) {
        RequestData data = this.getData("InsertObject");
        data.setStringParam(filePath);
        data.setFilename(filePath.substring(filePath.lastIndexOf("\\") + 1));
        list.add(data);
    }

    /**
     * 插入需要缩进Heading的word对象
     * @param filePath word文件路径
     * @param indent 缩进级别
     */
    public void insertWord(String filePath, int indent)
    {
        RequestData data = this.getData("InsertWord");
        data.setStringParam(filePath);
        data.setFilename(filePath.substring(filePath.lastIndexOf("\\") + 1));
        data.setIndent(indent);
        list.add(data);
    }

    /**
     * 插入分页
     */
    public void insertPageBreak() {
        RequestData data = this.getData("InsertPageBreak");
        list.add(data);
    }

    /**
     * 插入空白行
     * @param line 行数
     */
    public void insertLineBreak(int line) {
        RequestData data = this.getData("InsertLineBreak");
        data.setIntParam(line);
        list.add(data);
    }

    /**
     * 插入目录
     */
    public void insertIndex() {
        RequestData data = this.getData("InsertIndex");
        list.add(data);
    }

    /**
     * 插入书签
     */
    public void insertBookmark(String key) {
        RequestData data = this.getData("InsertBookmark");
        data.setStringParam(key);
        list.add(data);
    }

    /**
     * 插入目录书签
     */
    public void insertIndexBookmark() {
        RequestData data = this.getData("InsertBookmark");
        data.setStringParam("IndexBookmark");
        list.add(data);
    }

    /**
     * 归档
     */
    public void collect() {
    	String tempPath = System.getProperty("java.io.tmpdir");
    	if(!"\\".equals(tempPath.charAt(tempPath.length() - 1))) {
    		tempPath += "\\";
    	}
    	String xmlPath = tempPath + "request.xml";
    	String zipPath = tempPath + docName + ".zip";
    	RequestList requestList = new RequestList();
    	requestList.setRequestList(this.list);
    	XmlUtil.serializer(requestList, xmlPath);
    	// 对输出文件做CRC32校验
    	OutputStream outStream = null;
		CheckedOutputStream cos = null;
		ZipOutputStream zos = null;
    	try{
    		outStream = new FileOutputStream(zipPath);
    		cos = new CheckedOutputStream(outStream, new CRC32());
    		zos = new ZipOutputStream(cos);
    		zos.setEncoding("UTF-8");
	    	for (int i = 0; i < list.size(); i++) {
				RequestData data = list.get(i);
				if("InsertObject".equals(data.getServiceName()) || "InsertWord".equals(data.getServiceName())
						|| "InsertTemplate".equals(data.getServiceName())){
					File file = new File(data.getStringParam());
					putEntry(zos, file);
				}else if("InsertFrontCover".equals(data.getServiceName())) {
					File file = new File(data.getStringParam());
					File target = new File(tempPath + "cover.doc");
					Files.copy(file, target);
					putEntry(zos, target);
				}
			}
			File xmlFile = new File(xmlPath);
	    	putEntry(zos, xmlFile);
			xmlFile.delete();
    	}catch(IOException e) {
    		e.printStackTrace();
    	}finally {
    		try{
    			if(zos != null ) {
    				zos.flush();
        			zos.close();
    				zos = null;
    			}
    			if(cos != null ) {
    				cos.close();
    				cos = null;
    			}
    			if(outStream != null ) {
    				outStream.close();
    				outStream = null;
    			}
    		}catch(Exception e) {
    			e.printStackTrace();
    		}
    	}
    	invokeCollect(zipPath);
    }

    /**
     * 获取归档文件
     * @param docName 文件名
     * @return 文件流
     */
    public byte[] getDocument() {
    	Map<String, Object> params = new HashMap<String, Object>();
        params.put("docName", docName);
        String url = PropertyUtil.getServiceUrl();
        Request request =  Request.create(url + "/GetDocument",
        		METHOD.POST, params);
        PostSender sender = new PostSender(request);
        Response response = sender.send();
    	String xml = response.getContent();
    	String s = getBase64Binary(xml);
    	byte[] b = null;
    	BASE64Decoder dec = new BASE64Decoder();
		try {
			b = dec.decodeBuffer(s);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
    	return b;
	}


    private void putEntry(ZipOutputStream zos, File file) {
    	BufferedInputStream bis = null;
    	InputStream inStream = null;
    	try {
	    	ZipEntry entry = new ZipEntry(file.getName());
			zos.putNextEntry(entry);
			inStream = new FileInputStream(file);
			bis = new BufferedInputStream(inStream);
			int count;
			byte buf[] = new byte[BUFFER];
			while ((count = bis.read(buf, 0, BUFFER)) != -1) {
				zos.write(buf, 0, count);
			}
	    }catch(IOException e) {
			e.printStackTrace();
		}finally {
			try{
				if(bis != null ) {
					bis.close();
					bis = null;
					zos.closeEntry();
				}
				if(inStream != null ) {
					inStream.close();
					inStream = null;
				}
			}catch(IOException e) {
				e.printStackTrace();
			}
		}
    }

    public void invokeCollect(String zipPath) {
    	System.out.println("-------- invokeCollect ----------");
        String isLocal = PropertyUtil.getProperty("isLocal");
        String requestPath = PropertyUtil.getProperty("collectPath");
        if (("1".equals(isLocal)) || ("yes".equals(isLocal)) || ("true".equals(isLocal))) {
            System.out.println("------- zip文件本地移动 --------");
            File src = new File(zipPath);
            File target = new File(requestPath + "\\request\\" + src.getName());
            Files.copy(src, target);
            return;
        }
    	File f = new File(zipPath);
    	FileInputStream stream = null;
		byte[] data = null;
    	try {
			stream = new FileInputStream(f);
			data = Streams.readBytes(stream);
			BASE64Encoder enc = new BASE64Encoder();
			String str = enc.encode(data);
			Map<String, Object> params = new HashMap<String, Object>();
	        params.put("filename", f.getName());
	        params.put("s", str);
	        String url = PropertyUtil.getServiceUrl();
	        Request request =  Request.create(url + "/Collect",
	        		METHOD.POST, params);
	        PostSender sender = new PostSender(request);
	        sender.send();
    	} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} finally {
		    try {
		        if(stream != null) {
                    stream.close();
                    stream = null;
		        }
		    } catch (IOException e) {
                e.printStackTrace();
            }
		}
	}

    public String getBase64Binary(String xml) {
    	String result = "";
    	InputStream ins = null;
		try {
			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
	        DocumentBuilder builder = dbf.newDocumentBuilder();
	        ins = new ByteArrayInputStream(xml.getBytes("UTF-8"));
	        Document doc = builder.parse(ins);
	        XPathFactory factory = XPathFactory.newInstance();
	        XPath xpath = factory.newXPath();
	        XPathExpression expr = xpath.compile("string");
	        NodeList nodes = (NodeList) expr.evaluate(doc, XPathConstants.NODESET);
	        result = nodes.item(0).getTextContent();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (SAXException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (XPathExpressionException e) {
			e.printStackTrace();
		} finally {
            try {
                if(ins != null) {
                    ins.close();
                    ins = null;
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

}
