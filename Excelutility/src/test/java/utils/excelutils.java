package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelutils {
	
	// function has local variables like workbook,sheet
	// hard coded values like this String value = sheet.getRow(1).getCell(0).getStringCellValue();
	

	public static void getRowCount() throws IOException{
		//function to get row count
		String Excelpath ="./data/TestData.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(Excelpath);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println("no of rows : "+rowcount);
		
		
		
	}
	
	public static void getCellData() throws IOException{

		//function to get row data	
		String Excelpath ="./data/TestData.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(Excelpath);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		String value = sheet.getRow(1).getCell(0).getStringCellValue();
		System.out.println(value);
		
		
		
	}
	
}
