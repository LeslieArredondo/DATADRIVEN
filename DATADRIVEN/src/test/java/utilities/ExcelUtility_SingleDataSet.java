package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility_SingleDataSet {
	
	private static XSSFWorkbook ExcelWBook;
	private static XSSFSheet ExcelWSheet;
	private static XSSFCell Cell;
	private static XSSFRow Row;
	
	
	public static void setExcelFile(String Path, String SheetName) throws Exception {
		try {
			//Open the excel file
			FileInputStream ExcelFile = new FileInputStream(Path);
			
			//Access the excel data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			ExcelWSheet = ExcelWBook.getSheet(sheetName);
		} catch (Exception e) {
			throw (e);
		}
	}
	
	public static String getCellData(int RowNum, int ColNum) throws Exception {
		try {
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String cellData = Cell.getStringCellValue();
			return cellData;
		} catch (Exception e) {
			return "";
		}
	}
	
	
	public static String getDateCellData(int RowNum, int ColNum) throws Exception {
		try {
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			
			Date dateValue = Cell.getDateCellValue();
			String dateStringFormat = df.format(dateValue);
			
			return dateStringFormat;
		} catch (Exception e) {
			return "";
		}
	}
	
	
	//Write in the Excel cell, String Result
	
	public static void setCellData(String Result, int RowNum, int ColNum) throws Exception {
		try {
			Row = ExcelWSheet.getRow(RowNum);
			Cell = Row.getCell(ColNum);
			if (Cell == null) {
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);
			} else {
				Cell.setCellValue(Result);
			}
			
			//Open the file to write the results
			FileOutputStream fileOut = new FileOutputStream(
					Constants.File_Path + Constants.File_Name);
			
			ExcelWBook.write(fileOut);
			fileOut.flush();
			fileOut.close();
			} catch (Exception e) {
				throw (e);
			}
	}
	
	//Write in the Excel cell, double Result
	
	public static void setCellData(double Result, int RowNum, int ColNum) throws Exception {
	try {
		Row = ExcelWSheet.getRow(RowNum);
		Cell = Row.getCell(ColNum);
		if (Cell == null) {
			Cell = Row.createCell(ColNum);
			Cell.setCellValue(Result);
		}else {
			Cell.setCellValue(Result);
		}
		
		//Open the file to write the results
		FileOutputStream fileOut = new FileOutputStream(
				Constants.File_Path + Constants.File_Name);
		
		ExcelWBook.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}catch (Exception e) {
		throw (e);
		}
	}
	
	
	
	
}
