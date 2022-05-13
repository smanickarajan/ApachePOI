package readingexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		String excelpath="..\\ExcelUtility\\src\\main\\resources\\Details.xlsx";
		
		FileInputStream fis=new FileInputStream(excelpath);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet= wb.getSheet("Sheet1");
		
		int rows=sheet.getLastRowNum();
		
		int cols=sheet.getRow(1).getLastCellNum();
		
		for (int r=0;r<rows;r++) {
			XSSFRow row = sheet.getRow(r);
			for (int c=0;c<cols;c++) {
				XSSFCell col = row.getCell(c);
				
				switch (col.getCellType()) {
				case STRING: System.out.print(col.getStringCellValue());break;
				case NUMERIC: System.out.print(col.getNumericCellValue());break;
				case BOOLEAN: System.out.print(col.getBooleanCellValue());break;
				}
				
				System.out.print(" |  ");
			}
			
			System.out.println();
		}
		System.out.println("--------------------------------------------------------------------------");	
		Iterator<Row> rowitr = sheet.rowIterator();
		
		while (rowitr.hasNext()) {
			Row row = rowitr.next();
			
			Iterator<Cell> cellitr = row.cellIterator();
			
			while(cellitr.hasNext()) {
				
				Cell cell = cellitr.next();
				switch (cell.getCellType()) {
				case STRING: System.out.print(cell.getStringCellValue());break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
				}	
				System.out.print(" |  ");
			}
			System.out.println();
		}
		
	}

}
