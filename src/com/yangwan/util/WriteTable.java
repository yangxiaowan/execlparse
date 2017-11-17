package com.yangwan.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import com.yangwan.bean.PersonWorkTimeInfo;
import com.yangwan.bean.WorkOnOffTime;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class WriteTable {

	public static SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	
	public static final String tableHeader[] = {"人员", "审计总时长(小时)", "开始时间", "结束时间",
			"时长(小时)", "总时长(小时)", "补休时长(小时)", "地点", "项目", "处理"}; 
	
	/**
	 * 存储表格文件的工作区
	 */
	public String filePath = "";  
	
	/**
	 * 创建表格存储目录
	 * @param dirPath
	 */
	public void createSaveDir(String dirPath){
		this.filePath = dirPath;
		try{
			File saveDir = new File(dirPath);
			if(saveDir.exists()){
				System.out.println("The File Directory is existed!!!");
			}else{
				saveDir.mkdirs();
			}
		}catch(Exception e){
			System.out.println("Create File Directory failed!!!");
			e.printStackTrace();
		}
	}
	
	public String getXLSName(String personName, String startTime, String endTime, String currenyTableName){
		StringBuffer xlsNameBuff = new StringBuffer();
		startTime = startTime.replaceAll("-", "");
		endTime = endTime.replaceAll("-", "");
		xlsNameBuff.append(startTime).append("-").append(endTime);
		xlsNameBuff.append("_").append(personName).append("_").append("南京全民乐彩非工作时间处理系统记录表").append(".xls");
		return xlsNameBuff.toString();
	}
	
	public List<WorkOnOffTime> getOvertimeList(String startDate, List<WorkOnOffTime> originalList){
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm");
		SimpleDateFormat standardSDF = new SimpleDateFormat("HH:mm");
		String dateList[] = startDate.split("-");
		List<WorkOnOffTime> overTimeList = new ArrayList<WorkOnOffTime>();
		String standarTime = "20:00";
		try{
			Long standarTL = standardSDF.parse(standarTime).getTime();
			for(WorkOnOffTime workTime : originalList){
				String offTime = workTime.getOffTime(); //获得下班时间
				if(offTime != null && ExcelUtil.isWorkTime(offTime)
						&& standardSDF.parse(offTime).getTime() >= standarTL){ //有下班签到时间，并且签到时间在晚上八点之后，则为加班时间
					Long overTL = standardSDF.parse(offTime).getTime() - standarTL;
					//开始加班时间
					String startTime = dateList[0] + "/" + dateList[1] + "/" + String.valueOf(workTime.getDay()) + " " + standarTime;
					//记录下班签到时间
					String endTime = dateList[0] + "/" + dateList[1] + "/" + String.valueOf(workTime.getDay()) + " " + offTime;
					WorkOnOffTime workOnOffTime = new WorkOnOffTime();
					workOnOffTime.setOnTime(startTime);
					workOnOffTime.setOffTime(endTime);
					workOnOffTime.setOvertime(overTL/60000.0/60.0);
					overTimeList.add(workOnOffTime);
				}
			}
		}catch(Exception e){
			System.out.println("时间转化错误!!!");
			e.printStackTrace();
		}
		return overTimeList;
	}
	
	public void writePersonWorktimeTable(PersonWorkTimeInfo personWorktime){
		String xlsName = getXLSName(personWorktime.getPersonName(), 
				personWorktime.getStartTime(), personWorktime.getEndTime(), "南京全民乐彩非工作时间处理系统记录表");
		String saveFile = filePath + File.separator +xlsName;
		File file = new File(saveFile);
		OutputStream os = null;
		WritableWorkbook workBook = null;
		try{
			if(!file.exists()){ //if the file is not existed, create it
				file.createNewFile();
			}
			os = new FileOutputStream(file); 
			workBook = Workbook.createWorkbook(os);
			try{
				WritableSheet sheet = workBook.createSheet("加班时长统计表", 0); 
				WritableCellFormat wc = new WritableCellFormat();   
				wc.setAlignment(Alignment.CENTRE);
				wc.setWrap(false);
				WritableCellFormat mergeWc = new WritableCellFormat();   
				mergeWc.setAlignment(Alignment.CENTRE);
				mergeWc.setVerticalAlignment(VerticalAlignment.CENTRE);
				for(int i = 0; i < tableHeader.length; i++){
					Label label = new Label(i, 0, tableHeader[i], wc);
					sheet.addCell(label);
				}
				List<WorkOnOffTime> overtimeList = getOvertimeList(personWorktime.getStartTime(),
						personWorktime.getWorkOnOffTimeList());
				sheet.mergeCells(0,1,0,overtimeList.size()); //合并人员单元格
				Label nameLabel = new Label(0, 1, personWorktime.getPersonName(), mergeWc); 
				sheet.addCell(nameLabel);
				NumberFormat numberFormat=NumberFormat.getNumberInstance() ;
				numberFormat.setMaximumFractionDigits(2);  //设置保留两位小数
				double totalOverTime = 0.0;
				int index = 0;
				int dayNum = 0;
				for(WorkOnOffTime workOnOffTime : overtimeList){ //填写加班开始时间和结束时间以及加班时长
					dayNum++;
					index = 2; //跳转加班开始时间
					Label startTimeLabel = new Label(index, dayNum, workOnOffTime.getOnTime(), wc);
					sheet.addCell(startTimeLabel);
					index = 3; //跳转至加班结束时间
					Label endTimeLabel = new Label(index, dayNum, workOnOffTime.getOffTime(), wc);
					sheet.addCell(endTimeLabel);
					index = 4; //跳转至加班时长
					totalOverTime += workOnOffTime.getOvertime();
					Label timeLabel = new Label(index, dayNum, numberFormat.format(workOnOffTime.getOvertime()), wc);
					sheet.addCell(timeLabel);
				}
				sheet.mergeCells(1, 1, 1, overtimeList.size());  //合并审计时长单元格
				Label totalOverTimeLabel = new Label(1, 1, numberFormat.format(totalOverTime), mergeWc); 
				sheet.addCell(totalOverTimeLabel);
				sheet.mergeCells(5, 1, 5, overtimeList.size());  //合并总时长单元格
				Label totalTimeLabel = new Label(5, 1, numberFormat.format(totalOverTime), mergeWc); //总时长
				sheet.addCell(totalTimeLabel);
				index = 6; //跳转至补休时长
				for(int j = 1; j <= overtimeList.size(); j++){
					Label readdTime = new Label(index, j, "0", wc);
					sheet.addCell(readdTime);
				}
				index = 7; //跳转至加班地点
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workPlaceLabel = new Label(index, j, "公司", wc);
					sheet.addCell(workPlaceLabel);
				}
				index = 8; //跳转至项目
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workProLabel = new Label(index, j, "大数据、mgr", wc);
					sheet.addCell(workProLabel);
				}
				index = 9; //跳转至处理
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workDellLabel = new Label(index, j, "需求开发", wc);
					sheet.addCell(workDellLabel);
				}
				sheet.setColumnView(0, 15);
				sheet.setColumnView(1, 20);
				sheet.setColumnView(2, 20);
				sheet.setColumnView(3, 20);
				sheet.setColumnView(4, 15);
				sheet.setColumnView(5, 15);
				sheet.setColumnView(6, 15);
				sheet.setColumnView(7, 15);
				sheet.setColumnView(8, 20);
				sheet.setColumnView(9, 20);
			}catch(Exception e){
				System.out.println("新增单元格失败!!!");
			}
			workBook.write();
			workBook.close();
		}catch(IOException ioE){
			System.out.println("Function|writePersonWorktimeTable|File Operation Exception!!!");
		}catch(Exception e){
			System.out.println("Function|writePersonWorktimeTable|Excel Operation Exception!!!");
		}
		
	}
	
}
