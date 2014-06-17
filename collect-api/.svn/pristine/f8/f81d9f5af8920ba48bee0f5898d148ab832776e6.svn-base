package com.topway.doc.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.nutz.lang.Streams;

import com.thoughtworks.xstream.XStream;
import com.topway.doc.entity.RequestData;
import com.topway.doc.entity.RequestList;

public class XmlUtil {
	
	public static void serializer(RequestList data, String filePath) {
		XStream xstream = new XStream();
        xstream.alias("RequestList", RequestList.class);
        xstream.alias("RequestData", RequestData.class);
        
        xstream.aliasField("RequestData", RequestList.class, "requestList");
        xstream.aliasField("Indent", RequestData.class, "indent");
        xstream.aliasField("IntParam", RequestData.class, "intParam");
        xstream.aliasField("ServiceName", RequestData.class, "serviceName");
        xstream.aliasField("Sequence", RequestData.class, "sequence");
        xstream.aliasField("StringParam", RequestData.class, "stringParam");
        xstream.aliasField("Filename", RequestData.class, "filename");
        xstream.aliasField("DocName", RequestData.class, "docName");
        xstream.aliasField("SplitParam", RequestData.class, "splitParam");
        FileOutputStream stream = null;
		try {
			stream = new FileOutputStream(filePath);
	        String xml = xstream.toXML(data);
	        Streams.writeAndClose(stream, xml.getBytes("UTF-8"));
		}catch(FileNotFoundException e) {
			System.err.println(e.getMessage());
		} catch (IOException e) {
			System.err.println(e.getMessage());
		}
	}
	
	/*
	public static void serializer(RequestList data, String filePath) {
		FileOutputStream stream = null;
		try {
			stream = new FileOutputStream(filePath);
			JAXBContext context = JAXBContext.newInstance(RequestList.class);
			Marshaller m = context.createMarshaller();
			m.marshal(data, stream);
		}catch(JAXBException e) {
			System.err.println(e.getMessage());
		}catch(FileNotFoundException e) {
			System.err.println(e.getMessage());
		}finally {
			try {
				if(stream != null) {
					stream.close();
					stream = null;
				}
			}catch(IOException e) {
				System.err.println(e.getMessage());
			}
		}
	}

	public static byte[] serializer(RequestList data) {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		try {
			JAXBContext context = JAXBContext.newInstance(RequestList.class);
			Marshaller m = context.createMarshaller();
			// m.marshal(data, stream);
		}catch(JAXBException e) {
			System.err.println(e.getMessage());
		}
		return stream.toByteArray();
	}
	
	public static RequestList deSerializer(byte[] buf) {
		RequestList rd = null;
		try {
			JAXBContext context = JAXBContext.newInstance(RequestList.class);
			ByteArrayInputStream stream = new ByteArrayInputStream(buf);
			Unmarshaller um = context.createUnmarshaller();
			// rd = (RequestList) um.unmarshal(stream);
		}catch(JAXBException e) {
			System.err.println(e.getMessage());
		}
		return rd;
	}
	*/
}
