package com.app.pojo;

public class FlawDetail {
	private String flawName;
	private String flawID;
	private String severity;
	private String status;
	private String solution;
	public String getSolution() {
		return solution;
	}
	public void setSolution(String solution) {
		this.solution = solution;
	}
	public String getFlawName() {
		return flawName;
	}
	public void setFlawName(String flawName) {
		this.flawName = flawName;
	}
	public String getFlawID() {
		return flawID;
	}
	public void setFlawID(String flawID) {
		this.flawID = flawID;
	}
	public String getSeverity() {
		return severity;
	}
	public void setSeverity(String severity) {
		this.severity = severity;
	}
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	@Override
	public String toString() {
		return "FlawDetail [flawName=" + flawName + ", flawID=" + flawID
				+ ", severity=" + severity + ", status=" + status + "]";
	}
	
	
}
