package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelStuff {

	public static void main(String[] args) throws Exception {

		// printAllSheetData();

		String[][] result = getAllSheetDate("MOCK_DATA.xlsx", "data");

		System.out.println(Arrays.deepToString(result));

	}

	public static void printAllSheetData() throws Exception {

		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);

		Sheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getLastCellNum();

		for (int i = 0; i < rowCount; i++) {

			System.out.println(" row number : " + (i + 1));

			for (int j = 0; j < colCount; j++) {

				Cell cell = sheet.getRow(i).getCell(j);
				System.out.print(cell.toString() + " | ");

			}
			System.out.println();

		}
		wb.close();

	}

	public static String[][] getAllSheetDate(String filePath, String SheetName) throws Exception {

		// File excelFile = new File("MOCK_DATA.xlsx") ;
		FileInputStream fis = new FileInputStream(filePath);
		Workbook wb = WorkbookFactory.create(fis);

		// Sheet sheet = wb.getSheetAt(0);
		Sheet sheet = wb.getSheet(SheetName);
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getLastCellNum();

		// String[][] data = new String[11][11] ;
		String[][] data = new String[rowCount][colCount];

		for (int i = 0; i < rowCount; i++) {

			// System.out.println(" row number : " + (i + 1));

			for (int j = 0; j < colCount; j++) {

				Cell cell = sheet.getRow(i).getCell(j);
				data[i][j] = cell.toString();
				// System.out.print(cell.toString() + " | ");

			}
			// System.out.println();

		}
		fis.close();
		wb.close();

		return data;

	}
	
	public static String getCellData(String filePath, String sheetName, int rowIndex, int colIndex) throws Exception{
		  
		  String[][] result = getAllSheetDate(filePath, sheetName); 
		  return result[rowIndex][colIndex] ; 
		  
		}

}