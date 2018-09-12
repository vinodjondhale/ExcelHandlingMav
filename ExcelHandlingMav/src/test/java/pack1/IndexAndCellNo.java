package pack1;

import java.io.FileInputStream;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class IndexAndCellNo {
public static void main(String[] args) throws Exception {
	String xlpath="./ExcelFile/Data.xlsx";
	FileInputStream stream= new FileInputStream(xlpath);
	Workbook wb = WorkbookFactory.create(stream);
	Sheet s = wb.getSheet("Sheet3");	
	
	int rowCount=s.getLastRowNum();// index of last row
	System.out.println(rowCount);
	
int cellCountRow=s.getRow(0).getLastCellNum();//no of cells
System.out.println(cellCountRow);
}
}
