package library;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelLib {
public static String readExcelDataString(String path,String sheet,int row,int cell) 
{
	String v=" ";
	try {
		FileInputStream file= new FileInputStream(path);
		Workbook workbook=WorkbookFactory.create(file);
		v=workbook.getSheet(sheet).getRow(row).getCell(cell).toString();
	} catch (Exception e) {
		
	}
//	return v;
}

public static void writeExcelDataString(String path,String sheet,int row,int cell,String cellData) 
{
	try {
		FileInputStream file= new FileInputStream(path);
		Workbook workbook=WorkbookFactory.create(file);
		workbook.getSheet(sheet).createRow(row).createCell(cell).setCellValue(cellData);
		
	FileOutputStream file2=new FileOutputStream(path);
	workbook.write(file2);
	} catch (Exception e) {		
	}
	
}

public static int getRowCount(String path,String sheet) 
{
	int rowCount=0;
	try {
		FileInputStream file= new FileInputStream(path);
		Workbook workbook=WorkbookFactory.create(file);
		rowCount=workbook.getSheet(sheet).getLastRowNum();
	} catch (Exception e) {
		
	}
	return rowCount;
}

public static int getCellCount(String path,String sheet,int row) 
{
	int cellCount=0;
	try {
		FileInputStream file= new FileInputStream(path);
		Workbook workbook=WorkbookFactory.create(file);
		cellCount=workbook.getSheet(sheet).getRow(row).getLastCellNum();
	} catch (Exception e) {
		
	}
	return cellCount;
}

}
