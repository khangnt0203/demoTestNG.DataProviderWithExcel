package utils;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {

//tập data	
public Object[][] testData(String excelPath, String sheetName) {
	
	ExcelUtils excel = new ExcelUtils(excelPath,sheetName);
	
	int rowCount = excel.getRowCount();
	int colCount = excel.getColCount();
	
	Object data[][]=new Object[rowCount-1][colCount];
	
	for(int i=1; i<rowCount; i++) {
		for(int j=0; j<colCount;j++) {
			String cellData = excel.getCellDataString(i,j);
			System.out.print(cellData+" | ");
			data[i-1][j]=cellData;
		}
	}
	return data;
}

//lấy tập dữ liệu data
@DataProvider(name="test1Data")
public Object[][] getData() {
	String excelPath="D:\\FPT-U\\Kì 4\\PRJ321_3w\\TestNG_ASM_3w\\LoginTestNG\\excel\\data.xlsx";
	
	Object data[][]=testData(excelPath, "Sheet1");
	
	return data;
}

//chạy test
@Test(dataProvider = "test1Data")
public void test1(String username, String password) {
	System.out.println(username+" | "+password);
}

}
