package com.yangwan.test;


import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

import org.junit.Test;

import com.yangwan.bean.PersonWorkTimeInfo;
import com.yangwan.util.ExcelUtil;
import com.yangwan.util.WriteTable;

import common.Logger;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class WorkBookTest {

	public static final Logger logger = Logger.getLogger(WorkBookTest.class);
	
	@Test
	public void sayHello(){
		System.out.println("hello!");
	}
	
	@Test
	public void runParse(){
		String filePath = "C:\\test.xls";
		Workbook  readWb = null;
		try{
			readWb = ExcelUtil.getReadWorkBook(filePath);
		}catch(Exception e){
			logger.info("create readWorkBook error!!!!");
			e.printStackTrace();
		}
		Sheet readsheet = readWb.getSheet(0);
		ExcelUtil excelUtil = new ExcelUtil();
		WriteTable writeTable = new WriteTable();
		List<PersonWorkTimeInfo> personList = excelUtil.parseWorkBook(readsheet);
		writeTable.createSaveDir("C:\\testxsl");
		for(PersonWorkTimeInfo temp : personList){
			writeTable.writePersonWorktimeTable(temp);
		}
	}
	
	
	public int addj(int j){
		return j+1;
	}
	
	@Test
	public void jadd(){
		int j = 0;
		System.out.println(addj(j++));
	}
	
	@Test
	public void testRegex(){
		String workTime = "09:18";
		String regex = "^\\d\\d:\\d\\d$";
		Pattern pattern = Pattern.compile(regex);
		System.out.print(pattern.matcher(workTime).find());
	}
	
	@Test
	public void testSDF(){
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
		try {
			Date date = sdf.parse("14:00");
			System.out.print(date.getTime());
		} catch (ParseException e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testReplace(){
		String startTime = "2017-11-15";
		String temp  =  startTime.replaceAll("-", "");
		System.out.println(temp);
	}
}
