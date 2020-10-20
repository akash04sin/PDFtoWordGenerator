package com.app.pojo;

import java.util.List;

public class PDFFileDetails {

	
	private String eaicode,projectName, issuedDate, projectCode, urlTested;
	private List<FlawDetail> flawDetailArray;
	
	
	public List<FlawDetail> getFlawDetailArray() {
		return flawDetailArray;
	}
	public void setFlawDetailArray(List<FlawDetail> flawDetailArray) {
		this.flawDetailArray = flawDetailArray;
	}
	public String getEaicode() {
		return eaicode;
	}
	public void setEaicode(String eaicode) {
		this.eaicode = eaicode;
	}
	public String getProjectName() {
		return projectName;
	}
	
	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}
	public String getIssuedDate() {
		return issuedDate;
	}
	public void setIssuedDate(String issuedDate) {
		this.issuedDate = issuedDate;
	}
	public String getProjectCode() {
		return projectCode;
	}
	public void setProjectCode(String projectCode) {
		this.projectCode = projectCode;
	}
	public String getUrlTested() {
		return urlTested;
	}
	public void setUrlTested(String urlTested) {
		this.urlTested = urlTested;
	}

	@Override
	public String toString() {
		return "FileDetails [eaicode=" + eaicode + ", projectName="
				+ projectName + ", issuedDate=" + issuedDate + ", projectCode="
				+ projectCode + ", urlTested=" + urlTested
				+ ", flawDetailArray=" + flawDetailArray + "]";
	}
			
}
