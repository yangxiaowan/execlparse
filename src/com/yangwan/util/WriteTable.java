package com.yangwan.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import com.yangwan.bean.PersonWorkTimeInfo;

import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class WriteTable {

	public static SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	
	public static final String tableHeader[] = {""}; 
	
	/**
	 * 存储表格文件的工作区
	 */
	public String filePath = "";  
	
	/**
	 * 创建表格存储目录
	 * @param dirPath
	 */
	public void createSaveDir(String dirPath){
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
		xlsNameBuff.append("_").append(personName).append("_").append("currenyTableName").append(".xls");
		return xlsNameBuff.toString();
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
			WritableSheet sheet = workBook.createSheet("加班时长统计表", 0); 
			
		}catch(IOException ioE){
			System.out.println("Function|writePersonWorktimeTable|File Operation Exception!!!");
		}
		
	}
	
}
