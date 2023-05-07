package utils;


import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	//static HSSFWorkbook workbook;
	//static HSSFSheet sheet;
	
	public ExcelUtils(String excelPath,String sheetName)
	{
		try {
			
			
		String excelpath = "./data/TestData.xlsx";
			
		 workbook = new XSSFWorkbook(excelpath);
			
			
		//InputStream file = new FileInputStream(excelPath);
		//workbook = new HSSFWorkbook(new POIFSFileSystem(file));
		 sheet = workbook.getSheet("Sheet1");
		 
		 
		}
		
		catch(Exception exp)
		{
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
		
	}
	
	
	
	public static void getCellData(int RowNum,int ColNum) throws IOException
	{
		
		
		
		
	
		
		DataFormatter formatter = new DataFormatter();
		Object value =formatter.formatCellValue(sheet.getRow(RowNum).getCell(ColNum));
		
	//String value =	sheet.getRow(1).getCell(0).getStringCellValue();
		//double value =	sheet.getRow(1).getCell(2).getNumericCellValue();
	System.out.println(value);
	}
	
	public static void getRowCount()
	{
		
		int RowCount=sheet.getPhysicalNumberOfRows();
		System.out.println("No. of Row count is: "+RowCount);
	
		
		}
	
	}


