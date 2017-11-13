package com.yangwan.test;


import org.junit.Test;

import com.yangwan.util.ExcelUtil;

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
	public void testWorkBook(){
		String filePath = "C:\\test.xls";
		Workbook  readWb = null;
		try{
			readWb = ExcelUtil.getReadWorkBook(filePath);
		}catch(Exception e){
			logger.info("create readWorkBook error!!!!");
			e.printStackTrace();
		}
		Sheet readsheet = readWb.getSheet(0);
		logger.info("table total columns :" + readsheet.getColumns());
		logger.info("table total rows :" + readsheet.getRows());
		for(int i = 0; i < readsheet.getRows(); i++){
			for(int j = 0; j < readsheet.getColumns(); j++){
				Cell cell = readsheet.getCell(j, i);   
				
			}
		}
	}
}
