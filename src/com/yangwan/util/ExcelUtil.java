package com.yangwan.util;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import com.yangwan.bean.PersonWorkTimeInfo;
import com.yangwan.bean.WorkOnOffTime;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ExcelUtil {

	public static final String[] namelist = {"祖亮", "王逸林", "顾飞", "何雁稳",
			"姚俊杰", "陈洋洋", "孙敏", "邵兆辉", "孙世民", "王云田", "杨万"};
	
	public static SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	/**
	 * 人员工号
	 */
	public static final int JOB_NUMBER = 1;
	
	/**
	 * 员工姓名
	 */
	public static final int EMPLOYEE_NAME = 2;
	
	/**
	 * 部门分类
	 */
	public static final int DEPARTMENT_CLASS = 3;
	
	/**
	 * 天数
	 */
	public static final int DAY_CLASS = 4;
	
	/**
	 * 上下班时间
	 */
	public static final int WORK_TIME = 5;
	
	/**
	 * 制表时间, 格式: 2017-11-01 12:28:14
	 */
	public static final int TABULATION_TIME = 6;
	
	/**
	 * 考勤日期 格式:2017-10-01～2017-10-31
	 */
	public static final int ATTENDANCE_DATE = 7;
	
	/**
	 * 无用cell,不解析
	 */
	public static final int UNUSE_CELL = 0;
	
	public int LAST_FLAG = 0;
	
	/**
	 * 根据问价路径获得execl文件的WorkBook
	 * @param filePath
	 * @return
	 * @throws Exception
	 */
	public static Workbook getReadWorkBook(String filePath)throws Exception{
		InputStream instream = new FileInputStream(filePath);
		return Workbook.getWorkbook(instream);
	}
	
	public int parseWorkBook(String contentLine){
		if(contentLine.equals("工号：")){
			LAST_FLAG = JOB_NUMBER;
		}else if(contentLine.equals("姓名：")){
			LAST_FLAG = EMPLOYEE_NAME;
		}else if(contentLine.equals("部门：")){
			LAST_FLAG = DEPARTMENT_CLASS;
		}else if(contentLine.startsWith("考勤日期")){
			LAST_FLAG = ATTENDANCE_DATE;
		}else if(contentLine.startsWith("制表时间")){
			LAST_FLAG = TABULATION_TIME;
		}else{
			LAST_FLAG = UNUSE_CELL;
		}
		return LAST_FLAG;
	}
	
	public String getCellContent(int j, int i, Sheet readsheet){
		return readsheet.getCell(j, i).getContents();
	}
	
	/**
	 * 判断cell中字段是否为上下班时间
	 * @param workTime
	 * @return
	 */
	public static boolean isWorkTime(String workTime){
		if(workTime == null){
			workTime = "";
		}
		String splits[]	 = workTime.split("\n");
		String regex = "^\\d\\d:\\d\\d$";
		Pattern pattern = Pattern.compile(regex);
		return pattern.matcher(splits[0]).find();
	}
	
	public WorkOnOffTime getOnOffTime(int index, int day, Sheet readsheet){
		WorkOnOffTime workOnOffTime = new WorkOnOffTime();
		String dayNum = readsheet.getCell(day, index).getContents();
		if(Integer.parseInt(dayNum) == day){
			List<String> workTimeList = new ArrayList<String>();
			index += 1;	//跳转至签到时间
			String workTime = "";
			while(isWorkTime((workTime = getCellContent(day, index, readsheet)))){//收集签到时间
				String wkList[] = workTime.split("\n");
				for(String wt : wkList){
					if(isWorkTime(wt)){
						workTimeList.add(wt);
					}
				}
				index ++; 
			}
			/*上下班时间收集规则: 
			 * 签到时间列表有2个以上（包括2个）的签到时间,且第一个与最后一个时间相差不过半小时,将第一个签到时间作为上班签到时间，最后一个时间作为下班签到时间。
			 * 特殊情况:
			 * 如果签到列表为空，改天没有出勤
			 * 如果签到列表只有一个签到时间（这种为忘记签到的特殊情况），下午2点之前作为上班签到时间，2点后作为下班签到时间。
			 * 忘记签到需要找管理员修改上下班等级表格，方可将计算当日出勤的加班时间，否则无效。
			 */
			String startTime = "";
			String endTime = "";
			if(workTimeList.size() >= 2){
				try{
					/**
					 * 以08:00为准,对应为0L, 1s = 60000L, 半小时为30 * 60000 = 1800000
					 */
					startTime = workTimeList.get(0);
					endTime = workTimeList.get(workTimeList.size() - 1);
					Long startTL =  sdf.parse(startTime).getTime();
					Long endTL = sdf.parse(endTime).getTime();
					if(endTL - startTL <= 1800000L){
						if(startTL >= 21600000L){
							startTime = ""; //无上班签到时间
						}else{
							endTime = ""; //无下班签到时间
						}
					}
				}catch(Exception e){
					System.out.println("时间转换失败!!! time :" + workTimeList.toString());
					e.printStackTrace();
				}
			}else if(workTimeList.size() == 1){
				try{
					Long timeL = sdf.parse(workTimeList.get(0)).getTime();
					if(timeL >= 21600000L){
						startTime = ""; //无上班签到时间
						endTime = workTimeList.get(0);
					}else{
						endTime = ""; //无下班签到时间
						startTime = workTimeList.get(0);
					}
				}catch(Exception e){
					System.out.println("时间转换失败!!! time :" + workTimeList.toString());
					e.printStackTrace();
				}
			}else{
				startTime = "";
				endTime = "";
			}
			workOnOffTime.setOnTime(startTime);
			workOnOffTime.setOffTime(endTime);
			workOnOffTime.setDay(day);
			return workOnOffTime;
		}else{
			System.out.println("解析出勤天数异常！ dayNum :" + dayNum);
			return null;
		}
	}
	
	public List<PersonWorkTimeInfo> parseWorkBook(Sheet readsheet){
		List<PersonWorkTimeInfo> personList = new ArrayList<>();
		PersonWorkTimeInfo personWorkTimeInfo = null;
		String timeRange = ""; //统计出勤时间范围
		String startTime = "";  //统计出勤开始时间
		String endTime = ""; //统计出勤结束时间
		String tabulationTime = ""; //统计时间
		for(int i = 0; i < readsheet.getRows(); i++){
			for(int j = 0; j < readsheet.getColumns(); j++){
				Cell cell = readsheet.getCell(j, i);   
				String content = cell.getContents();
				switch(parseWorkBook(content)){ //对excel中每个单元进行检验
					case JOB_NUMBER:
						//工号标识，开始个人工时统计
						personWorkTimeInfo = new PersonWorkTimeInfo();
						personWorkTimeInfo.setStartTime(startTime);
						personWorkTimeInfo.setEndTime(endTime);
						personWorkTimeInfo.setTabulationTime(tabulationTime);
						j += 2;
						personWorkTimeInfo.setJobNumber(getCellContent(j, i, readsheet)); //获取员工工号
						break;
					case UNUSE_CELL: //无需解析cell,直接跳过
						break;
					case EMPLOYEE_NAME:
						j += 1;
						personWorkTimeInfo.setPersonName(getCellContent(j, i, readsheet)); //获取员工姓名
						break;
					case DEPARTMENT_CLASS:
						j += 1;
						personWorkTimeInfo.setDepartment(getCellContent(j, i, readsheet)); //获取员工部门
						i += 1;  //跳转至出勤天数行
						List<WorkOnOffTime> list = new ArrayList<WorkOnOffTime>();
						for(j = 1; j < readsheet.getColumns(); j++){
							WorkOnOffTime tempWk = getOnOffTime(i, j, readsheet);
							list.add(tempWk);
						}
						personWorkTimeInfo.setWorkOnOffTimeList(list);
						personList.add(personWorkTimeInfo);
						i += 1; 
						break;
					case ATTENDANCE_DATE: //出勤统计时间范围
						timeRange = content.split("：")[1];//获得统计时间范围
						startTime = timeRange.split("～")[0];
						endTime = timeRange.split("～")[1];
						break;
					case TABULATION_TIME:
						tabulationTime = content.split("：")[1]; //获得制表时间
						break;
				}
			}
		}
		return personList;
	}
	
	
}
