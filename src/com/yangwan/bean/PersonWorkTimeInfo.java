package com.yangwan.bean;

import java.util.Date;

public class PersonWorkTimeInfo {

	/**
	 * ����
	 */
	private String personName;
	
	/**
	 * Ա������
	 */
	private String jobNumber;
	
	/**
	 * ��������
	 */
	public Date workDate;
	
	/**
	 * �ϰ�ʱ��
	 */
	public String onTime;
	
	/**
	 * �°�ʱ��
	 */
	public String offTime;

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

	public Date getWorkDate() {
		return workDate;
	}

	public void setWorkDate(Date workDate) {
		this.workDate = workDate;
	}

	public String getOnTime() {
		return onTime;
	}

	public void setOnTime(String onTime) {
		this.onTime = onTime;
	}

	public String getOffTime() {
		return offTime;
	}

	public void setOffTime(String offTime) {
		this.offTime = offTime;
	}
	
}
