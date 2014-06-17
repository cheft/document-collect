package com.topway.doc.entity;

import java.util.ArrayList;
import java.util.List;

public class RequestList {

	private List<RequestData> requestList = new ArrayList<RequestData>();

	//@XmlElementWrapper(name="RequestData")
	public List<RequestData> getRequestList() {
		return requestList;
	}

	public void setRequestList(List<RequestData> requestList) {
		this.requestList = requestList;
	}
	
}
