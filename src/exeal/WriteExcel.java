package exeal;


import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException 
	{
		//data
		//create workbook and spreadsheet
		//logic
		//write
//Data		
		int[] serial = new int[10];     //for column 0
		for(int i=0; i<serial.length; i++) 
		{
			serial[i] = i + 1;
		}
		
		//for column 1
		String[] name = new String[10];
		name[0] = "Student A";
		name[1] = "Student B";
		name[2] = "Student C";
		name[3] = "Student D";
		name[4] = "Student E";
		name[5] = "Student F";
		name[6] = "Student G";
		name[7] = "Student H";
		name[8] = "Student I";
		name[9] = "Student K";
		
		//for column 2
		String[] result = new String[10];
		result[0] = "Pass";
		result[1] = "Pass";
		result[2] = "Fail";
		result[3] = "Pass";
		result[4] = "Pass";
		result[5] = "Fail";
		result[6] = "Pass";
		result[7] = "Fail";
		result[8] = "Pass";
		result[9] = "Pass";

//create workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		
//create spreadsheet
		XSSFSheet sheet = wb.createSheet("Sheet1");
		
// create rows		
		XSSFRow row;
		row = sheet.createRow(0);
		
		XSSFCell cell0 = row.createCell(0);
		XSSFCell cell1 = row.createCell(1);
		XSSFCell cell2 = row.createCell(2);
		
// logic 
		for (int i=0; i<serial.length; i++) 
		{
			row = sheet.createRow(i+1);
			for(int j=0; j<3;j++) 
			{
				XSSFCell cell = row.createCell(j);
				
				if(cell.getColumnIndex()==0) 
				{
					cell.setCellValue(serial[i]);
				}
				else if(cell.getColumnIndex()==1) 
				{
					cell.setCellValue(name[i]);
				}
				else if(cell.getColumnIndex()==2) 
				{
					cell.setCellValue(result[i]);
				}
			}
		}
//Write in excel sheet
		String path = "D:\\DriverForSelenium\\Tes.xlsx";
		FileOutputStream out = new FileOutputStream(path);
		wb.write(out);
		
		System.out.println("File is Generated.....");
		
		out.close();

	}

}
