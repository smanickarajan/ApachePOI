package writingexcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel2 {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Emp Info");

		Object empdata[][] = { { "EmpID", "Name", "Job" ,"IsMarried"}, { 101, "David", "Engineer",true }, { 102, "Scott", "Manager",false },
				{ 103, "Smith", "Analyst",true } };

		int rowcount=0;
		for (Object emp[]:empdata) {
			XSSFRow row = sheet.createRow(rowcount++);
			int colcount=0;
			for(Object value:emp) {
				XSSFCell cell = row.createCell(colcount++);
				
				if (value instanceof String)
					cell.setCellValue(value.toString());
				else if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				else if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
		}
		String excelpath = "..\\ExcelUtility\\src\\main\\resources\\Employee1.xlsx";

		FileOutputStream outstream = new FileOutputStream(excelpath);
		wb.write(outstream);
		outstream.close();

		System.out.println("Employee1 saved");
	}

}
