package pack1;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PrintMatrix {
	public static void main(String[] args) throws Exception {
		String xlpath="./ExcelFile/Data.xlsx";
		FileInputStream stream= new FileInputStream(xlpath);
		Workbook wb = WorkbookFactory.create(stream);
		Sheet s = wb.getSheet("Sheet3");	
		Cell c = null;
		for (int i = 0; i < 3; i++) {
			for (int j = 0; j < 3; j++) {
			c.getRow().getCell(j);
			System.out.println(c+" ");
			}
			System.out.println();
		}
			
		
	}
	
}
