package guviTask13;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingData {

	public static void main(String[] args) throws IOException {
		
		// Define the file path
		File src = new File("Utils\\Write&readData.xlsx");
		FileInputStream readfile = new FileInputStream(src);
		// Opening the workbook
		XSSFWorkbook workbook = new XSSFWorkbook(readfile);
		// Opening the sheet with index
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		// Get the number of rows in the sheet
		int sizeOfRow = sheet1.getPhysicalNumberOfRows();
		// Iterate through each row
		for (int i = 0; i < sizeOfRow; i++) {
			XSSFRow row = sheet1.getRow(i);
			// Get the number of cells in the current row
			int sizeOfCell = row.getPhysicalNumberOfCells();
			// Iterate through each cell in the row
			for (int j = 0; j < sizeOfCell; j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = getCellValue(cell);
				System.out.print(" " + cellValue);
			}
			System.out.println(" ");
		}
		workbook.close();

	}

	// Method to get the cell value as a String
	public static String getCellValue(XSSFCell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case STRING:
			return cell.getStringCellValue();
		default:
			return cell.getStringCellValue();
		}

	}
}


// Output :
//	 Name Age Email 
//	 John Doe 30 john@test.com 
//	 Jane Doe 28 john@test.com 
//	 Bob Smith 35 jacky@example.com 
//	 Swapnil 37 Swapnil@example.com 
//	 