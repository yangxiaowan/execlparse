package com.yangwan.util;

import java.io.FileInputStream;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ExcelUtil {

	
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
		if(Integer.valueOf(contentLine) >= 1 && Integer.valueOf(contentLine) <= 31 && LAST_FLAG != JOB_NUMBER){
			LAST_FLAG = DAY_CLASS;
		}else if(contentLine.equals("���ţ�")){
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
	
	public void parseWorkBook(Sheet readsheet){
		for(int i = 0; i < readsheet.getRows(); i++){
			for(int j = 0; j < readsheet.getColumns(); j++){
				Cell cell = readsheet.getCell(j, i);   
				String content = cell.getContents();
				switch(parseWorkBook(content)){
					
				}
			}
		}
	}
}
