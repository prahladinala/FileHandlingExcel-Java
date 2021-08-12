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
public class FileHandling {

	public static void main(String[] args) {
		
		try {
			File src = new File(".\\datafiles\\Animals.xlsx");
			FileInputStream fis = new FileInputStream(src);
			
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			
			XSSFSheet sh = wb.getSheet("Sheet1");
			
			Row row = sh.createRow(9);
			
			Cell cell0  = row.createCell(0);
			cell0.setCellValue(cell0.getStringCellValue());
			cell0.setCellValue("Golden Frog");
			
			Cell cell1  = row.createCell(1);
			cell1.setCellValue(cell1.getStringCellValue());
			cell1.setCellValue("One Horned Rhino");
			
			Cell cell2  = row.createCell(2);
			cell2.setCellValue(cell2.getStringCellValue());
			cell2.setCellValue("Dolphin");
			
			Cell cell3  = row.createCell(3);
			cell3.setCellValue(cell3.getStringCellValue());
			cell3.setCellValue("Sabre tooth");
			
			Cell cell4  = row.createCell(4);
			cell4.setCellValue(cell4.getStringCellValue());
			cell4.setCellValue("Dinosaur");
			
			Cell cell5  = row.createCell(5);
			cell5.setCellValue(cell5.getStringCellValue());
			cell5.setCellValue("Mammoth");
			
			Cell cell6  = row.createCell(6);
			cell6.setCellValue(cell6.getStringCellValue());
			cell6.setCellValue("Neanderthal");
			
			Cell cell7  = row.createCell(7);
			cell7.setCellValue(cell7.getStringCellValue());
			cell7.setCellValue("Tiger");
			
			FileOutputStream fos = new FileOutputStream(".\\datafiles\\Animals.xlsx");
			wb.write(fos);
			wb.close();
			fos.close();
			System.out.println("File Created Successfully");
		}catch(java.io.IOException e) {
			System.out.println(e);
		}

	}

}
