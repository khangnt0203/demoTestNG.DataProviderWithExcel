package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
static XSSFWorkbook workBook;
static XSSFSheet sheet;

//kết nối excel tương tự database
public ExcelUtils(String excelPath, String sheetName) {
	try {
		workBook=new XSSFWorkbook(excelPath);
		sheet = workBook.getSheet(sheetName);
	}catch (Exception e) {
		// TODO: handle exception
		e.printStackTrace();
	}
}

//lấy số dòng
public static int getRowCount() {
	int rowCount =0;
	try {
		rowCount = sheet.getPhysicalNumberOfRows();
		//System.out.println("No of rows: "+rowCount);
	}catch (Exception e) {
		// TODO: handle exception
		System.out.println(e.getMessage());
		System.out.println(e.getCause());
		e.printStackTrace();
	}
	return rowCount;
}

//lấy số cột
public static int getColCount() {
	int colCount=0;
	try {
		colCount=sheet.getRow(0).getPhysicalNumberOfCells();
		//System.out.println("No of cols: "+colCount);
	}catch (Exception e) {
		// TODO: handle exception
		System.out.println(e.getMessage());
		System.out.println(e.getCause());
		e.printStackTrace();
	}
	return colCount;
}

//lấy giá trị dạng string từng dòng
public static String getCellDataString(int row, int col) {
	String cellData=null;
	try {
		cellData=sheet.getRow(row).getCell(col).getStringCellValue();
		//System.out.println(cellData);
	}catch (Exception e) {
		// TODO: handle exception
		System.out.println(e.getMessage());
		System.out.println(e.getCause());
		e.printStackTrace();
	}
	return cellData;
}
}
