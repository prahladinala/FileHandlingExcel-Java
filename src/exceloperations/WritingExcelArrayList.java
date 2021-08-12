package exceloperations;

import java.io.FileOutputStream;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelArrayList {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// WORKBOOK		
		XSSFWorkbook workbook = new XSSFWorkbook();
				
		// SHEET
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
		empdata.add(new Object[]{"EmpID", "Name", "Job"});
		empdata.add(new Object[]{101, "David", "Engineers"});
		empdata.add(new Object[]{102, "Prahlad", "Developer"});
		empdata.add(new Object[]{103, "Raghu", "Analyist"});

		
		// USING FOR EACH LOOP
		int rownum = 0;
		for(Object[] emp:empdata)
		{
			//CREATE A ROW
			XSSFRow row = sheet.createRow(rownum++);
					
			int cellnum = 0;
			for(Object value:emp)
			{
				XSSFCell cell = row.createCell(cellnum++);
						
				if(value instanceof String)
					cell.setCellValue((String)value);
						
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
						
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
				
		String filePath = ".\\datafiles\\empArrayList2.xlsx";
		
		FileOutputStream outStream = new FileOutputStream(filePath);
				
		workbook.write(outStream);
				
		outStream.close();
		System.out.println("Employee File Created Succesfully");
	}

}
