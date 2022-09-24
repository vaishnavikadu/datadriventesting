package ExcelReading;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadingPractice {

	public void getSheetDetails() throws IOException {
		FileInputStream fis = new FileInputStream("E:\\Acceleration\\09Apr\\Day1\\AppData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// number of sheet present in workbook
		int sheetCount = workbook.getNumberOfSheets();
		// active sheet number
		int activeSheetNumber = workbook.getActiveSheetIndex();
		System.out.println("Total sheet : " + sheetCount);
		System.out.println("Active sheet : " + activeSheetNumber);
		for (int i = 0; i < sheetCount; i++) {
			System.out.println(workbook.getSheetName(i));
		}
	}

	public void getRowOperations() throws IOException {
		FileInputStream fis = new FileInputStream("E:\\Acceleration\\09Apr\\Day1\\AppData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("LoginDetails");
		int rowCount = sheet.getLastRowNum();
		System.out.println("Number of rows in LoginDetails : " + rowCount);
	}

	public void getCellOperator() throws IOException {
		FileInputStream fis = new FileInputStream("E:\\Acceleration\\09Apr\\Day1\\AppData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("LoginDetails");
		XSSFRow row = sheet.getRow(1);
		int cellCount = row.getLastCellNum();
		System.out.println("Last cell count : " + cellCount);
		for (int i = 0; i < cellCount; i++) {
			System.out.println(row.getCell(i).getStringCellValue());
		}
	}

	public void getDifferentCellValues() throws IOException {
		FileInputStream fis = new FileInputStream("E:\\Acceleration\\09Apr\\Day1\\AppData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("LoginDetails");
		XSSFRow row = sheet.getRow(1);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			XSSFCell cell = row.getCell(i);
			switch (cell.getCellType()) {
			case STRING:
				System.out.println(cell.getStringCellValue());
				break;
			case NUMERIC:
				System.out.println(cell.getNumericCellValue());
				break;
			case BOOLEAN:
				System.out.println(cell.getBooleanCellValue());
				break;

			case BLANK:
				break;
			default:
				throw new RuntimeException("There is support for this method");
			}
		}
	}

	public static void main(String[] args) {
		ExcelReadingPractice obj = new ExcelReadingPractice();
		try {
//			obj.getSheetDetails();
//			obj.getRowOperations();
//			obj.getCellOperator();
			obj.getDifferentCellValues();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
