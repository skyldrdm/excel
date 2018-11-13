package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WorkingExcel {

	public static void main(String[] args) throws Exception {

		printAllSheetData();

		// Workbook --> Sheet ---> Row--> Cell 
		//degisik

		// Eealier version of poi library
		// have 2 different set of classes to work with xls , xlsx files
		/*
		 * xls files --- MS Excel 97-2003 HSSFWorkbook , HSSFSheet , HSSFRow , HSSFCell
		 * xlsx XSSFWorkbook , XSSFSheet , XSSFRow , XSSFCell
		 */
		/*
		 * The main difference between these two file extensions is that the XLS is
		 * created on the version of Excel prior to 2007 while XLSX is created on the
		 * version of Excel 2007 and onward. They are also different on the basis of
		 * information storing way. XLS is a binary format while that XLSX is Open XML
		 * format.
		 */

		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);

		System.out.println(wb.getNumberOfSheets());

		// Sheet sh = wb.getSheet("data");
		Sheet sh = wb.getSheetAt(0);
		Row row1 = sh.getRow(1);
		Cell c1 = row1.getCell(0);
		System.out.println(c1);

		int columnCountInFirstRow = row1.getLastCellNum();
		System.out.println(columnCountInFirstRow);

		int rowCount = sh.getLastRowNum();
		System.out.println(rowCount);

		// getPhysicalNumberOfRows will return actual rowNumber
		// whether you have empty value row or not

		int actualRowCount = sh.getPhysicalNumberOfRows();
		System.out.println(actualRowCount);

		for (int i = 0; i < actualRowCount; i++) {
			System.out.println("ROW NUMBER : " + (i + 1));

			Row row = sh.getRow(i);

			for (int j = 0; j < columnCountInFirstRow; j++) {

				Cell cell = row.getCell(j);
				System.out.print(cell + "---");

			}
			System.out.println();
		}
	}
	// Task 1

	// Create a utility method to store all sheetData
	// in two dimensional String Array
	// method name : getAllSheetData
	// return type : none
	// params : ()
	// logic , print everything in nice format

	public static void printAllSheetData() throws Exception {

		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);

		Sheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getLastCellNum();

		for (int i = 0; i < rowCount; i++) {

			System.out.println(" row number : " + i);

			for (int j = 0; j < colCount; j++) {

				Cell cell = sheet.getRow(i).getCell(j);
				System.out.print(cell.toString() + " | ");

			}
			System.out.println();

		}

		wb.close();
	}
	// Task 2

	// Create a utility method to store all sheetData
	// in two dimensional String Array
	// method name : getAllSheetData
	// return type : String[][]
	// params : FileName as String , SheetName

	public static String[][] getAllSheetData(String filePath, String SheetName) throws Exception {

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
	
	//create a method called getCellData(String filePath, String sheetName, int rowIndex, int colIndex)
	//return value as String
	public String getCellData(String filePath, String sheetName, int rowIndex, int colIndex) throws Exception {
		
	String[][]result=	getAllSheetData(filePath, sheetName);		
		
		return result[rowIndex][colIndex];
		
		

		
		
	}
}
