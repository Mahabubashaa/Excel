package excelutils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class ReadExcel {

	public static void main(String[] args) throws IOException {

		String path = "D:\\Arunsundar\\Docs\\work.xls";
		FileInputStream fin = new FileInputStream(path);

		HSSFWorkbook book = new HSSFWorkbook(fin);
		HSSFSheet sheet = book.getSheet("Sheet1");

		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			HSSFRow currentRow = sheet.getRow(i);
			for (int j = 0; j < currentRow.getPhysicalNumberOfCells(); j++) {
				Cell currentCell = currentRow.getCell(j);
				switch (currentCell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					System.out.print(currentCell.getStringCellValue() + " | ");
					break;
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(currentCell.getNumericCellValue() + "|");
					break;
				}
			}
			System.out.println("test");

		}

		List<HashMap<String, String>> mydata = new ArrayList<HashMap<String, String>>();
		HSSFRow HeaderRow = sheet.getRow(0);
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			HSSFRow currentRow = sheet.getRow(i);
			HashMap<String, String> currentHash = new HashMap<String, String>();

			for (int j = 0; j < currentRow.getPhysicalNumberOfCells(); j++) {
				Cell currentCell = currentRow.getCell(j);
				switch (currentCell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					currentHash.put(HeaderRow.getCell(j).getStringCellValue(),
							currentCell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					currentHash.put(HeaderRow.getCell(j).getStringCellValue(),String.valueOf(currentCell.getNumericCellValue()));
					break;
				}
			}
			mydata.add(currentHash);
		}
		System.out.println(mydata);
		HashMap<String, String> map = mydata.get(0);
		for (Entry<String, String> entry : map.entrySet()) {

			if (entry.getKey().equalsIgnoreCase("Name")) {
				System.out.println(entry.getValue());
		}
		}
	}
}
