/**
 * 
 */
package appium.Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import io.appium.java_client.android.AndroidDriver;

/**
 * @author TrungTH_CA
 *
 */
public class ReadExcel {

	static AndroidDriver<WebElement> driver;
	static DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
	static LocalDateTime now = LocalDateTime.now();
	static boolean isAndroid=false;
	
	public static AndroidDriver<WebElement> ReadAll(String fileNameConfig) throws NumberFormatException, Exception{
		String filePathConfig = System.getProperty("user.dir") + "//Test_Input";
		String sheetNameConfig = "TestConfig";
		
		for(int i=Constant.RUN_Android_COL_NUM;i<=Constant.RUN_IOs_COL_NUM;i++){
			ArrayList<String> runColumn = ReadExcel.readExcelFileAtColumn(filePathConfig, fileNameConfig,
					sheetNameConfig, i);
			System.out.println(Constant.narrow+ "\n"+runColumn.toString());
			for (int j = 1; j < runColumn.size(); j++) {
				String sheet = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, j,
						Constant.SHEET_COL_NUM);
				if (sheet.equalsIgnoreCase("")) {
					int temp = j;
					while (sheet.equalsIgnoreCase("")) {
						temp--;
						sheet = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, temp,
								Constant.SHEET_COL_NUM);
						
					}
				}
				String fromRow = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, j,
						Constant.FROM_ROW_COL_NUM);
				String toRow = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, j,
						Constant.TO_ROW_COL_NUM);
				//System.out.println("runcolumn"+ runColumn.toString()+";"+runColumn.size() + "j: " +j);
				if(runColumn.get(j).toLowerCase().equalsIgnoreCase(Constant.yes)){
					driver = readExcelFile(null, j, filePathConfig, fileNameConfig, sheet,
							Integer.parseInt(fromRow), Integer.parseInt(toRow), true,
							Constant.TOTAL_COLUMN_NUMBER);
				}
			}
		}
		return driver;
	}
	
	
	public static ArrayList<String> readExcelFileAtColumn(String filePath, String fileName, String sheetName,
			int column) throws IOException {
		ArrayList<String> columnData = new ArrayList<String>();
		Workbook wb = newWorkbook(filePath, fileName);
		Sheet sheet = wb.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 0; i < rowCount + 1; i++) {
			try {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(column);
				cell.setCellType(cell.CELL_TYPE_STRING);
				columnData.add(cell.getStringCellValue());
			} catch (Exception e) {
				// if row(i) = null / empty
				columnData.add("");
			}
		}
		return columnData;
	
	}
	
	
	public static Workbook newWorkbook(String filePath, String fileName) throws IOException {
		File file = new File(filePath + "//" + fileName);
		FileInputStream inputStream = new FileInputStream(file);
		Workbook wb = null;
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		if (fileExtensionName.equalsIgnoreCase(".xlsx")) {
			wb = new XSSFWorkbook(inputStream);
		} else if (fileExtensionName.equalsIgnoreCase(".xls")) {
			wb = new HSSFWorkbook(inputStream);
		}
		return wb;
	}
	
	public static String readExcelFileAtCell(String filePath, String fileName, String sheetName, int row, int column)
			throws IOException {
		String CellData;
		String columnRefer;
		String rowRefer;
		CellData= getDataAtCell(filePath, fileName, sheetName, row, column);
		
		// check if Data is reference from another cell in this sheet
		if (CellData.startsWith("!")) {
			// read reference data then refill it into current cell
			columnRefer = CellData.substring(1, CellData.indexOf(":"));
			rowRefer = CellData.substring(CellData.indexOf(":") + 1, CellData.length());
			CellData = readExcelFileAtCell(filePath, fileName, sheetName, Integer.parseInt(rowRefer),
					Integer.parseInt(columnRefer));	
			
		}

		return CellData;
	}
	
	
	@SuppressWarnings("deprecation")
	public static String getDataAtCell(String filePath, String fileName, String sheetName, int row, int column) throws IOException{
		Workbook wb = newWorkbook(filePath, fileName);
		String CellData;
		Sheet sheet = wb.getSheet(sheetName);		
		try {
			Row rowExcel = sheet.getRow(row);
			Cell cell = rowExcel.getCell(column);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			CellData = cell.getStringCellValue();
		} catch (Exception e) {
			CellData = "";
		}
		return CellData;
	}
	
	
	/*
	 * read excel test case with sheetName
	 */
	public static AndroidDriver<WebElement> readExcelFile(WebDriver localDriver, int browserColumn, String filePath, String fileName,
			String readSheetName, int fromRow, int toRow, boolean haveRelated, int maxColumn) throws Exception {
		// Clear previous result on test case file

		System.out.println(Constant.narrow);
		for (int i = fromRow; i < toRow ; i++) {
			try {
				ArrayList<String> eachRow = ReadExcel.readExcelFileAtRow(filePath, fileName, readSheetName, i, 0,
						maxColumn);
				/*	keyword_executor(localDriver, browserColumn, eachRow, filePath, fileName, readSheetName, i,
							maxColumn, haveRelated);*/
				System.out.println(i+" "+eachRow);
			} catch (Exception e) {
				System.out.println(e.getMessage());
			/*	ExcelUtils.writeExcelFileWithResultTest(browserColumn, filePath, fileName, readSheetName, i, "FAILED",
						dtf.format(now), e.getMessage(), maxColumn);*/
				System.out.println("error gi do");
			}
		}
		return driver;
	}
	
	
	// read data at specific row from 'startColumn' to 'endColumn'
		@SuppressWarnings({ "deprecation", "static-access" })
		public static ArrayList<String> readExcelFileAtRow(String filePath, String fileName, String sheetName, int row,
				int startColumn, int endColumn) throws IOException {
			ArrayList<String> rowData = new ArrayList<String>();
			try {
				Workbook wb = newWorkbook(filePath, fileName);
				Sheet sheet = wb.getSheet(sheetName);
				Row rowExcel = null;
				rowExcel = sheet.getRow(row);
				for (int i = startColumn; i <= endColumn; i++) {
					try {
						Cell cell = rowExcel.getCell(i);
						cell.setCellType(cell.CELL_TYPE_STRING);
						rowData.add(cell.getStringCellValue());
					} catch (Exception e) {
						// if row(i) = null / empty
						rowData.add("");
					}
				}
			} catch (Exception e) {
				System.out.println("Cannot read data at row:" + row);
			}
			return rowData;
		}
}
