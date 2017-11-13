package com.yangwan.bean;

import java.util.Date;

public class PersonWorkTimeInfo {

	/**
	 * 姓名
	 */
	private String personName;
	
	/**
	 * 员工工号
	 */
	private String jobNumber;
	
	/**
	 * 工作日期
	 */
	public Date workDate;
	
	/**
	 * 上班时间
	 */
	public String onTime;
	
	/**
	 * 下班时间
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
