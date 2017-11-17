package com.yangwan.bean;

import java.util.List;

public class PersonWorkTimeInfo {

	/**
	 * ��ʼʱ��
	 */
	private String startTime;
	
	/**
	 * ����ʱ��
	 */
	private String endTime;
	
	/**
	 * ǩ��ʱ���б�
	 */
	private List<WorkOnOffTime> workOnOffTimeList;
	
	/**
	 * ͳ��ʱ��
	 */
	private String  tabulationTime;
	
	/**
	 * ����
	 */
	private String personName;
	
	/**
	 * Ա������
	 */
	private String jobNumber;
	
	/**
	 * ����
	 */
	private String department;
	
	public List<WorkOnOffTime> getWorkOnOffTimeList() {
		return workOnOffTimeList;
	}

	public void setWorkOnOffTimeList(List<WorkOnOffTime> workOnOffTimeList) {
		this.workOnOffTimeList = workOnOffTimeList;
	}

	public String getStartTime() {
		return startTime;
	}

	public void setStartTime(String startTime) {
		this.startTime = startTime;
	}

	public String getEndTime() {
		return endTime;
	}

	public void setEndTime(String endTime) {
		this.endTime = endTime;
	}

	public String getDepartment() {
		return department;
	}

	public void setDepartment(String department) {
		this.department = department;
	}

	public String getTabulationTime() {
		return tabulationTime;
	}

	public void setTabulationTime(String tabulationTime) {
		this.tabulationTime = tabulationTime;
	}

	public String getPersonName() {
		return personName;
	}

	public void setPersonName(String personName) {
		this.personName = personName;
	}

	public String getJobNumber() {
		return jobNumber;
	}

	public void setJobNumber(String jobNumber) {
		this.jobNumber = jobNumber;
	}

}
