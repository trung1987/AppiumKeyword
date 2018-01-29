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
	
	public static AndroidDriver<WebElement> ReadAll(String fileNameConfig) throws IOException{
		String filePathConfig = System.getProperty("user.dir") + "//Test_Input";
		String sheetNameConfig = "TestConfig";
		
		for(int i=Constant.RUN_Android_COL_NUM;i<=Constant.Run_IOS;i++){
			ArrayList<String> runColumn = ReadExcel.readExcelFileAtColumn(filePathConfig, fileNameConfig,
					sheetNameConfig, i);
			System.out.println(Constant.narrow+ "\n"+runColumn.toString());
			
			String sheet = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, i,
					Constant.SHEET_COL_NUM);
			if (sheet.equalsIgnoreCase("")) {
				int temp = i;
				while (sheet.equalsIgnoreCase("")) {
					temp--;
					sheet = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, temp,
							Constant.SHEET_COL_NUM);
					
				}
			}
			
			
			String fromRow = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, i,
					Constant.FROM_ROW_COL_NUM);
			String toRow = ReadExcel.readExcelFileAtCell(filePathConfig, fileNameConfig, sheetNameConfig, i,
					Constant.TO_ROW_COL_NUM);
			
			
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
}
