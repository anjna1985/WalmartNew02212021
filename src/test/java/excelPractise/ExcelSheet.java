package excelPractise;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

public class ExcelSheet {

	public static void main(String[] args) throws IOException { // TestNG

		WebDriver driver;
		//Workbook Class
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet0 = workbook.createSheet("FirstSheet");

		for (int rows = 0; rows < 10; rows++) {
			Row row = sheet0.createRow(rows);

			for (int cols = 0; cols < 10; cols++) {
				Cell cell = row.createCell(cols);
				cell.setCellValue((int) (Math.random() * 100)); // 0.20, 0.23 - 20, 23, 24
			}

		}

		// Connecting with stream
		File f = new File("C:\\Users\\14012\\Desktop\\myExcel4.xlsx");
		FileOutputStream fo = new FileOutputStream(f);
		// FileInputStream
		workbook.write(fo);
		fo.close();
		System.out.println("File Created");

	}
}
