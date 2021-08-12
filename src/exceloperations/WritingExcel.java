package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel {

	// WORKBOOK -> SHEET -> ROWS -> CELLS
	public static void main(String[] args) throws IOException {
		
		// WORKBOOK		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		// SHEET
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		// 
		Object empdata[][] = {
				{"EmpID", "Name", "Job"},
				{101, "David", "Engineerss"},
				{102, "Prahlad", "Developer"},
				{103, "Raghu", "Analyist"},
				{104, "Venkatesh", "Engineer"},
				{105, "Smith", "Manager"}
				};
		
		// USING FOR LOOP
//		int rows = empdata.length;
//		int cols = empdata[0].length;
//		
//		System.out.println(rows); //6
//		System.out.println(cols); //3
//		
//		//WRITING ROWS
//		for(int r = 0; r<rows; r++)
//		{
//			//CREATE A ROW
//			XSSFRow row = sheet.createRow(r);
//			//WRITING CELLS
//			for(int c = 0; c<cols; c++)
//			{
//				// CREATE A CELl
//				XSSFCell cell = row.createCell(c);
//				
//				//WRITE THE DATA
//				Object value = empdata[r][c];
//				
//				// UPDATE IN EXCEL
//				if(value instanceof String)
//					//IF value CONTAINS STRING
//					cell.setCellValue((String)value);
//				
//				if(value instanceof Integer)
//					cell.setCellValue((Integer)value);
//				
//				if(value instanceof Boolean)
//					cell.setCellValue((Boolean)value);
//				
//				
//			}
//		}
		
		
		// USING FOR EACH LOOP
		int rowCount = 0;
		for(Object emp[]:empdata)
		{
			//CREATE A ROW
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount = 0;
			for(Object value:emp)
			{
				XSSFCell cell = row.createCell(columnCount++);
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
		
		String filePath = ".\\datafiles\\emp.xlsx";
		
		FileOutputStream outStream = new FileOutputStream(filePath);
		
		workbook.write(outStream);
		
		outStream.close();
		System.out.println("Employee File Created Succesfully");
	}

}
