package utils;

import java.io.IOException;

public class main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		String excelpath ="./data/TestData.xlsx";
		String sheetname = "sheet1";
		excelutils1 excel	= new excelutils1(excelpath,sheetname);
		
		excel.getCellData(1,2);
		excel.getRowCount();
		
	}

}
