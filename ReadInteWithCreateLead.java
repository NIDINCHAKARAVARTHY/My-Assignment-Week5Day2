package week5day2;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadInteWithCreateLead {
	@Test
	public static String[][] readdata(String datafile) throws IOException {
		
		
		XSSFWorkbook wb= new XSSFWorkbook("./data/"+datafile+".xlsx");
		XSSFSheet  sheet = wb.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		System.out.println("LAST ROW NO : " + lastRowNum);
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		System.out.println("LAST ROW NO : "+lastCellNum);
		
		String [][]data = new String [lastRowNum][lastCellNum];
		for (int i = 1; i <= lastRowNum; i++) {
			XSSFRow row = sheet.getRow(i);
			
			for (int j = 0; j < lastCellNum; j++) {
				XSSFCell cell = row.getCell(j);
				data[i-1][j] = cell.getStringCellValue();
			
				
				
			}
			
		}
		
		wb.close();
		return data;
	}

	
}
