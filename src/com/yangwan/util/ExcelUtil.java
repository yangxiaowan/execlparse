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

	public static final String[] namelist = {"����", "������", "�˷�", "������",
			"Ҧ����", "������", "����", "���׻�", "������", "������", "����"};
	
	public static SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	/**
	 * ��Ա����
	 */
	public static final int JOB_NUMBER = 1;
	
	/**
	 * Ա������
	 */
	public static final int EMPLOYEE_NAME = 2;
	
	/**
	 * ���ŷ���
	 */
	public static final int DEPARTMENT_CLASS = 3;
	
	/**
	 * ����
	 */
	public static final int DAY_CLASS = 4;
	
	/**
	 * ���°�ʱ��
	 */
	public static final int WORK_TIME = 5;
	
	/**
	 * �Ʊ�ʱ��, ��ʽ: 2017-11-01 12:28:14
	 */
	public static final int TABULATION_TIME = 6;
	
	/**
	 * �������� ��ʽ:2017-10-01��2017-10-31
	 */
	public static final int ATTENDANCE_DATE = 7;
	
	/**
	 * ����cell,������
	 */
	public static final int UNUSE_CELL = 0;
	
	public int LAST_FLAG = 0;
	
	/**
	 * �����ʼ�·�����execl�ļ���WorkBook
	 * @param filePath
	 * @return
	 * @throws Exception
	 */
	public static Workbook getReadWorkBook(String filePath)throws Exception{
		InputStream instream = new FileInputStream(filePath);
		return Workbook.getWorkbook(instream);
	}
	
	public int parseWorkBook(String contentLine){
		if(contentLine.equals("���ţ�")){
			LAST_FLAG = JOB_NUMBER;
		}else if(contentLine.equals("������")){
			LAST_FLAG = EMPLOYEE_NAME;
		}else if(contentLine.equals("���ţ�")){
			LAST_FLAG = DEPARTMENT_CLASS;
		}else if(contentLine.startsWith("��������")){
			LAST_FLAG = ATTENDANCE_DATE;
		}else if(contentLine.startsWith("�Ʊ�ʱ��")){
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
	 * �ж�cell���ֶ��Ƿ�Ϊ���°�ʱ��
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
			index += 1;	//��ת��ǩ��ʱ��
			String workTime = "";
			while(isWorkTime((workTime = getCellContent(day, index, readsheet)))){//�ռ�ǩ��ʱ��
				String wkList[] = workTime.split("\n");
				for(String wt : wkList){
					if(isWorkTime(wt)){
						workTimeList.add(wt);
					}
				}
				index ++; 
			}
			/*���°�ʱ���ռ�����: 
			 * ǩ��ʱ���б���2�����ϣ�����2������ǩ��ʱ��,�ҵ�һ�������һ��ʱ��������Сʱ,����һ��ǩ��ʱ����Ϊ�ϰ�ǩ��ʱ�䣬���һ��ʱ����Ϊ�°�ǩ��ʱ�䡣
			 * �������:
			 * ���ǩ���б�Ϊ�գ�����û�г���
			 * ���ǩ���б�ֻ��һ��ǩ��ʱ�䣨����Ϊ����ǩ�������������������2��֮ǰ��Ϊ�ϰ�ǩ��ʱ�䣬2�����Ϊ�°�ǩ��ʱ�䡣
			 * ����ǩ����Ҫ�ҹ���Ա�޸����°�ȼ���񣬷��ɽ����㵱�ճ��ڵļӰ�ʱ�䣬������Ч��
			 */
			String startTime = "";
			String endTime = "";
			if(workTimeList.size() >= 2){
				try{
					/**
					 * ��08:00Ϊ׼,��ӦΪ0L, 1s = 60000L, ��СʱΪ30 * 60000 = 1800000
					 */
					startTime = workTimeList.get(0);
					endTime = workTimeList.get(workTimeList.size() - 1);
					Long startTL =  sdf.parse(startTime).getTime();
					Long endTL = sdf.parse(endTime).getTime();
					if(endTL - startTL <= 1800000L){
						if(startTL >= 21600000L){
							startTime = ""; //���ϰ�ǩ��ʱ��
						}else{
							endTime = ""; //���°�ǩ��ʱ��
						}
					}
				}catch(Exception e){
					System.out.println("ʱ��ת��ʧ��!!! time :" + workTimeList.toString());
					e.printStackTrace();
				}
			}else if(workTimeList.size() == 1){
				try{
					Long timeL = sdf.parse(workTimeList.get(0)).getTime();
					if(timeL >= 21600000L){
						startTime = ""; //���ϰ�ǩ��ʱ��
						endTime = workTimeList.get(0);
					}else{
						endTime = ""; //���°�ǩ��ʱ��
						startTime = workTimeList.get(0);
					}
				}catch(Exception e){
					System.out.println("ʱ��ת��ʧ��!!! time :" + workTimeList.toString());
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
			System.out.println("�������������쳣�� dayNum :" + dayNum);
			return null;
		}
	}
	
	public List<PersonWorkTimeInfo> parseWorkBook(Sheet readsheet){
		List<PersonWorkTimeInfo> personList = new ArrayList<>();
		PersonWorkTimeInfo personWorkTimeInfo = null;
		String timeRange = ""; //ͳ�Ƴ���ʱ�䷶Χ
		String startTime = "";  //ͳ�Ƴ��ڿ�ʼʱ��
		String endTime = ""; //ͳ�Ƴ��ڽ���ʱ��
		String tabulationTime = ""; //ͳ��ʱ��
		for(int i = 0; i < readsheet.getRows(); i++){
			for(int j = 0; j < readsheet.getColumns(); j++){
				Cell cell = readsheet.getCell(j, i);   
				String content = cell.getContents();
				switch(parseWorkBook(content)){ //��excel��ÿ����Ԫ���м���
					case JOB_NUMBER:
						//���ű�ʶ����ʼ���˹�ʱͳ��
						personWorkTimeInfo = new PersonWorkTimeInfo();
						personWorkTimeInfo.setStartTime(startTime);
						personWorkTimeInfo.setEndTime(endTime);
						personWorkTimeInfo.setTabulationTime(tabulationTime);
						j += 2;
						personWorkTimeInfo.setJobNumber(getCellContent(j, i, readsheet)); //��ȡԱ������
						break;
					case UNUSE_CELL: //�������cell,ֱ������
						break;
					case EMPLOYEE_NAME:
						j += 1;
						personWorkTimeInfo.setPersonName(getCellContent(j, i, readsheet)); //��ȡԱ������
						break;
					case DEPARTMENT_CLASS:
						j += 1;
						personWorkTimeInfo.setDepartment(getCellContent(j, i, readsheet)); //��ȡԱ������
						i += 1;  //��ת������������
						List<WorkOnOffTime> list = new ArrayList<WorkOnOffTime>();
						for(j = 1; j < readsheet.getColumns(); j++){
							WorkOnOffTime tempWk = getOnOffTime(i, j, readsheet);
							list.add(tempWk);
						}
						personWorkTimeInfo.setWorkOnOffTimeList(list);
						personList.add(personWorkTimeInfo);
						i += 1; 
						break;
					case ATTENDANCE_DATE: //����ͳ��ʱ�䷶Χ
						timeRange = content.split("��")[1];//���ͳ��ʱ�䷶Χ
						startTime = timeRange.split("��")[0];
						endTime = timeRange.split("��")[1];
						break;
					case TABULATION_TIME:
						tabulationTime = content.split("��")[1]; //����Ʊ�ʱ��
						break;
				}
			}
		}
		return personList;
	}
	
	
}
