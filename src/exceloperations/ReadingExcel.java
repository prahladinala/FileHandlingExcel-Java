package exceloperations;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;
public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub 

		String excelFilePath = ".\\datafiles\\countries.xlsx";
		
		// FILE INPUT STREAM TO OPEN FILE IN READING MODE
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		// GET WORKBOOK FROM THIS FILE
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		// GET SHEET FROM WORKBOOK
		// XSSFSheet sheet = workbook.getSheet("Sheet1"); // BY SHEET NAME
		XSSFSheet sheet = workbook.getSheetAt(0); // BY SHEET INDEX
		
		// READ DATA (ALL ROWS AND COLUMNS)
		
		// USING FOR LOOP
		
		// GET NUMBER OF ROWS AND COLUMNS
		
		// GET NUMBER OF ROWS
		int rows = sheet.getLastRowNum();
		//GET ROW FROM SHEET (GET NUMBER OF COLUMS)
		int cols = sheet.getRow(1).getLastCellNum();
		
		System.out.println("################FOR LOOP#######################");
		//FOR LOOP
		for(int r=0; r<=rows; r++)
		{
			// GET ROW FROM SHEET EXAMPLE in first loop we get first row and so on
			XSSFRow row = sheet.getRow(r);
			for(int c=0; c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				// FIND TYPE OF CELL => String or Int or Boolean or Formula etc
				// cell.getCellType();
				// IF STRING => getStringCellValue method
				// IF INT => getNumbericCellValue method
				// etc
				switch(cell.getCellType())
				{
//				case STRING: System.out.println(cell.getStringCellValue());
//				break;
//				
//				case NUMERIC: System.out.println(cell.getNumericCellValue());
//				break;
//				
//				case BOOLEAN: System.out.println(cell.getBooleanCellValue());
//				break;
				
				case STRING: System.out.print(cell.getStringCellValue());
				break;
				
				case NUMERIC: System.out.print(cell.getNumericCellValue());
				break;
				
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
				break;
				
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
		
		System.out.println("################ITERATOR#######################");

		// USING ITERATOR
		Iterator iterator = sheet.iterator();
		while(iterator.hasNext())
		{
			//RETURNS ALL THE ROWS EX 1st iterator it gets first row
			XSSFRow row = (XSSFRow) iterator.next();
			
			Iterator cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext())
			{
				XSSFCell cell =(XSSFCell) cellIterator.next();
				
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue());
				break;
				
				case NUMERIC: System.out.print(cell.getNumericCellValue());
				break;
				
				case BOOLEAN: System.out.print(cell.getBooleanCellValue());
				break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
	}

}
