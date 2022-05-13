package hashmapexcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMaptoExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Emp Info");
		
		Map<String,String> studata=new HashMap<>();
		
		studata.put("101", "David");
		studata.put("102", "Smith");
		studata.put("103", "Graham");
		studata.put("104", "Jon");
		
		int rowno=0;
		
		for ( Entry<String, String> entry:studata.entrySet()) {
			XSSFRow row = sheet.createRow(rowno++);
			row.createCell(0).setCellValue(entry.getKey());
			row.createCell(1).setCellValue(entry.getValue());
			
		}
		String excelpath = "..\\ExcelUtility\\src\\main\\resources\\Student.xlsx";

		FileOutputStream outstream = new FileOutputStream(excelpath);
		wb.write(outstream);
		outstream.close();

		System.out.println("Student saved");
	}

}
