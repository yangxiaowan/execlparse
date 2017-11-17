package com.yangwan.bean;

import java.util.List;

public class PersonWorkTimeInfo {

	/**
	 * 开始时间
	 */
	private String startTime;
	
	/**
	 * 结束时间
	 */
	private String endTime;
	
	/**
	 * 签到时间列表
	 */
	private List<WorkOnOffTime> workOnOffTimeList;
	
	/**
	 * 统计时间
	 */
	private String  tabulationTime;
	
	/**
	 * 姓名
	 */
	private String personName;
	
	/**
	 * 员工工号
	 */
	private String jobNumber;
	
	/**
	 * 部分
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
