package Pack;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	XSSFSheet ss;
	public String readData(int i, int j) {
		Row r = ss.getRow(i);
		Cell c= r.getCell(j);
		int celltype =c.getCellType();  // to get the return value from cell; 0 means numeric, 1 means string
		switch(celltype)
		{
		case Cell.CELL_TYPE_NUMERIC:
		{
			double a = c.getNumericCellValue();
			return String.valueOf(a);
		}
		case Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
		}
		}		
		return null;
	}
	
	public Excel() throws IOException{
		FileInputStream f= new FileInputStream ("C:\\Users\\bobyp\\Documents\\Test.xlsx");
		XSSFWorkbook w=new XSSFWorkbook(f);
		ss=w.getSheet("Sheet1");
		
	}
}
