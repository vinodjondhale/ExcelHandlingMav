package pack1;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ReadData {
public static void main(String[] args) throws Exception {
String xlpath="./ExcelFile/Data.xlsx";
FileInputStream stream= new FileInputStream(xlpath);
Workbook wb = WorkbookFactory.create(stream);
Sheet s = wb.getSheet("Sheet1");
Row r=s.getRow(0);
Cell c=r.getCell(0);
String v=c.toString();
System.out.println(v);
//optimized code
System.out.println(wb.getSheet("Sheet1").getRow(0).getCell(0).toString());

}
}
