package pack1;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PrintEntireExcelSheet {
	public static void main(String[] args) throws Exception {
		String xlpath="./ExcelFile/Data.xlsx";
		FileInputStream stream= new FileInputStream(xlpath);
		Workbook wb = WorkbookFactory.create(stream);
		Sheet s = wb.getSheet("Sheet3");	
		
		int rowCount=s.getLastRowNum();// index of last row
		for (int i = 0; i < rowCount+1; i++) {
			int cellCount=s.getRow(i).getLastCellNum();
			for (int j = 0; j < cellCount; j++) {
				System.out.println(s.getRow(i).getCell(j)+" ");
			}
		}
	}
}
