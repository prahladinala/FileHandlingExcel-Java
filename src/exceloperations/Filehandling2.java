package exceloperations;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Filehandling2 {

	public static void main(String[] args) {

		try {
			File src = new File(".\\datafiles\\Animals.xlsx");
			FileInputStream fis = new FileInputStream(src);
			
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			
			XSSFSheet sh = wb.getSheet("Sheet1");
			
			Row row = sh.getRow(9);
			
			//DELETE CELL
			Cell cell = row.getCell(7);
			row.removeCell(cell);
			
			//UPDATE CELL
			Cell cell1 = row.getCell(2);
			cell1.setCellValue("Susu Dolphin");
			
			FileOutputStream fos = new FileOutputStream(src);
			wb.write(fos);
			wb.close();
			fos.close();
			System.out.println("File Updated Successfully");
		}catch(java.io.IOException e) {
			System.out.println(e);
		}

	}

}
