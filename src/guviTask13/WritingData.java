package guviTask13;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingData {

	public static void main(String[] args) {

		// Define the file path
		File src = new File("Utils\\Write&readData.xlsx");
		// Creating the workbook
		XSSFWorkbook workBook = new XSSFWorkbook();
		// Creating the Sheet
		XSSFSheet sheet1 = workBook.createSheet("Sheet1");
		// Creating the row
		XSSFRow row = sheet1.createRow(0);
		// Send the data to cell created
		row.createCell(0).setCellValue("Name");
		row.createCell(1).setCellValue("Age");
		row.createCell(2).setCellValue("Email");
		// Creating the row1
		XSSFRow row1 = sheet1.createRow(1);
		// Send the data to cell created
		row1.createCell(0).setCellValue("John Doe");
		row1.createCell(1).setCellValue("30");
		row1.createCell(2).setCellValue("john@test.com");
		// Creating the row2
		XSSFRow row2 = sheet1.createRow(2);
		// Send the data to cell created
		row2.createCell(0).setCellValue("Jane Doe");
		row2.createCell(1).setCellValue("28");
		row2.createCell(2).setCellValue("john@test.com");
		// Creating the row3
		XSSFRow row3 = sheet1.createRow(3);
		// Send the data to cell created
		row3.createCell(0).setCellValue("Bob Smith");
		row3.createCell(1).setCellValue("35");
		row3.createCell(2).setCellValue("jacky@example.com");
		// Creating the row4
		XSSFRow row4 = sheet1.createRow(4);
		// Send the data to cell created
		row4.createCell(0).setCellValue("Swapnil");
		row4.createCell(1).setCellValue("37");
		row4.createCell(2).setCellValue("Swapnil@example.com");

		// Write the workbook to the file
		try (FileOutputStream writeData = new FileOutputStream(src)) {
			workBook.write(writeData);
			System.out.println("Excel file written successfully.");
		} catch (IOException e) {
			System.out.println(" IOEception error  " +e);
		}
	}

}

// Output :
//	 Excel file written successfully
