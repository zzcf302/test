
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class Test {

	public static void main(String[] args) throws Exception {
		Workbook wb = new HSSFWorkbook();
		//Sheet
		Sheet s = wb.createSheet("First");
		for (int i = 0; i < 9; i++) {
			Row row = s.createRow(i);
			for (int j = 0; j <= i; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue((i+1) + " x " + (j+1) + " = " + (i+1)*(j+1));
			}
		}
		
		//Cell
		
		wb.write(new FileOutputStream("e:/workbook.xls"));
		
	}

}
