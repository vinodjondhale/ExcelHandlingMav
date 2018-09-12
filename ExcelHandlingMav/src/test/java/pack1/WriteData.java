package pack1;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteData {
public static void main(String[] args) throws Exception {
	String xlpath="./ExcelFile/Data.xlsx";
	FileInputStream stream= new FileInputStream(xlpath);
	Workbook wb = WorkbookFactory.create(stream);
	Sheet s = wb.getSheet("Sheet2");
	Row r = s.createRow(0);
	Cell c = r.createCell(0);
	c.setCellValue("vinod");
	FileOutputStream stream2=new FileOutputStream(xlpath);
	wb.write(stream2);
	System.out.println("Cell Value is entered");
	
}
}
