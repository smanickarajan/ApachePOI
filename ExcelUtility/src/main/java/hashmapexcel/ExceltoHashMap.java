package hashmapexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceltoHashMap {

	public static void main(String[] args) throws IOException {
		String excelpath = "..\\ExcelUtility\\src\\main\\resources\\Student.xlsx";

		FileInputStream fis = new FileInputStream(excelpath);

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Emp Info");

		int rows = sheet.getLastRowNum();
		Map<String, String> data = new HashMap<>();

		for (int r = 0; r <= rows; r++) {
			String key = sheet.getRow(r).getCell(0).getStringCellValue();
			String value = sheet.getRow(r).getCell(1).getStringCellValue();

			data.put(key, value);
		}

		for (Entry<String, String> entry : data.entrySet()) {

			System.out.println(entry.getKey() + "  " + entry.getValue());
		}

	}
}
