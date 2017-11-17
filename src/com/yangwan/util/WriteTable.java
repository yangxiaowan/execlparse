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
	
	public static final String tableHeader[] = {"��Ա", "�����ʱ��(Сʱ)", "��ʼʱ��", "����ʱ��",
			"ʱ��(Сʱ)", "��ʱ��(Сʱ)", "����ʱ��(Сʱ)", "�ص�", "��Ŀ", "����"}; 
	
	/**
	 * �洢����ļ��Ĺ�����
	 */
	public String filePath = "";  
	
	/**
	 * �������洢Ŀ¼
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
		xlsNameBuff.append("_").append(personName).append("_").append("�Ͼ�ȫ���ֲʷǹ���ʱ�䴦��ϵͳ��¼��").append(".xls");
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
				String offTime = workTime.getOffTime(); //����°�ʱ��
				if(offTime != null && ExcelUtil.isWorkTime(offTime)
						&& standardSDF.parse(offTime).getTime() >= standarTL){ //���°�ǩ��ʱ�䣬����ǩ��ʱ�������ϰ˵�֮����Ϊ�Ӱ�ʱ��
					Long overTL = standardSDF.parse(offTime).getTime() - standarTL;
					//��ʼ�Ӱ�ʱ��
					String startTime = dateList[0] + "/" + dateList[1] + "/" + String.valueOf(workTime.getDay()) + " " + standarTime;
					//��¼�°�ǩ��ʱ��
					String endTime = dateList[0] + "/" + dateList[1] + "/" + String.valueOf(workTime.getDay()) + " " + offTime;
					WorkOnOffTime workOnOffTime = new WorkOnOffTime();
					workOnOffTime.setOnTime(startTime);
					workOnOffTime.setOffTime(endTime);
					workOnOffTime.setOvertime(overTL/60000.0/60.0);
					overTimeList.add(workOnOffTime);
				}
			}
		}catch(Exception e){
			System.out.println("ʱ��ת������!!!");
			e.printStackTrace();
		}
		return overTimeList;
	}
	
	public void writePersonWorktimeTable(PersonWorkTimeInfo personWorktime){
		String xlsName = getXLSName(personWorktime.getPersonName(), 
				personWorktime.getStartTime(), personWorktime.getEndTime(), "�Ͼ�ȫ���ֲʷǹ���ʱ�䴦��ϵͳ��¼��");
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
				WritableSheet sheet = workBook.createSheet("�Ӱ�ʱ��ͳ�Ʊ�", 0); 
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
				sheet.mergeCells(0,1,0,overtimeList.size()); //�ϲ���Ա��Ԫ��
				Label nameLabel = new Label(0, 1, personWorktime.getPersonName(), mergeWc); 
				sheet.addCell(nameLabel);
				NumberFormat numberFormat=NumberFormat.getNumberInstance() ;
				numberFormat.setMaximumFractionDigits(2);  //���ñ�����λС��
				double totalOverTime = 0.0;
				int index = 0;
				int dayNum = 0;
				for(WorkOnOffTime workOnOffTime : overtimeList){ //��д�Ӱ࿪ʼʱ��ͽ���ʱ���Լ��Ӱ�ʱ��
					dayNum++;
					index = 2; //��ת�Ӱ࿪ʼʱ��
					Label startTimeLabel = new Label(index, dayNum, workOnOffTime.getOnTime(), wc);
					sheet.addCell(startTimeLabel);
					index = 3; //��ת���Ӱ����ʱ��
					Label endTimeLabel = new Label(index, dayNum, workOnOffTime.getOffTime(), wc);
					sheet.addCell(endTimeLabel);
					index = 4; //��ת���Ӱ�ʱ��
					totalOverTime += workOnOffTime.getOvertime();
					Label timeLabel = new Label(index, dayNum, numberFormat.format(workOnOffTime.getOvertime()), wc);
					sheet.addCell(timeLabel);
				}
				sheet.mergeCells(1, 1, 1, overtimeList.size());  //�ϲ����ʱ����Ԫ��
				Label totalOverTimeLabel = new Label(1, 1, numberFormat.format(totalOverTime), mergeWc); 
				sheet.addCell(totalOverTimeLabel);
				sheet.mergeCells(5, 1, 5, overtimeList.size());  //�ϲ���ʱ����Ԫ��
				Label totalTimeLabel = new Label(5, 1, numberFormat.format(totalOverTime), mergeWc); //��ʱ��
				sheet.addCell(totalTimeLabel);
				index = 6; //��ת������ʱ��
				for(int j = 1; j <= overtimeList.size(); j++){
					Label readdTime = new Label(index, j, "0", wc);
					sheet.addCell(readdTime);
				}
				index = 7; //��ת���Ӱ�ص�
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workPlaceLabel = new Label(index, j, "��˾", wc);
					sheet.addCell(workPlaceLabel);
				}
				index = 8; //��ת����Ŀ
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workProLabel = new Label(index, j, "�����ݡ�mgr", wc);
					sheet.addCell(workProLabel);
				}
				index = 9; //��ת������
				for(int j = 1; j <= overtimeList.size(); j++){
					Label workDellLabel = new Label(index, j, "���󿪷�", wc);
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
				System.out.println("������Ԫ��ʧ��!!!");
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
