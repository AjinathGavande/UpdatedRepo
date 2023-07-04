package exeal;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExeal {

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		String path = "D:\\DriverForSelenium\\Testdata.xlsx";
		//FileInputStream is class to fetch the data from Excel sheet
		FileInputStream file = new FileInputStream(path);
		
		String message1 = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
		
		System.out.println(message1);
		
		FileInputStream file1 = new FileInputStream(path);
        double message2 = WorkbookFactory.create(file1).getSheet("Sheet1").getRow(1).getCell(1).getNumericCellValue();
		
		System.out.println(message2);
		
		FileInputStream file3 = new FileInputStream(path);
        CellType cellType = WorkbookFactory.create(file3).getSheet("Sheet1").getRow(0).getCell(0).getCellType();
		
		System.out.println(cellType);

	}

}
