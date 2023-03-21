package week5day2;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcel {
	
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb= new XSSFWorkbook("./data/CreateLead.xlsx");
		XSSFSheet  sheet = wb.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		System.out.println("LAST ROW NO : " + lastRowNum);
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		System.err.println("LAST ROW NO : "+lastCellNum);
		
		for (int i = 0; i <= lastRowNum; i++) {
			XSSFRow row = sheet.getRow(i);
			
			for (int j = 0; j < lastCellNum; j++) {
				XSSFCell cell = row.getCell(j);
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
				
				
			}
			
		}
		
		wb.close();
		
	}

}
