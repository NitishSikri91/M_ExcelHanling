package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClassOne {

	public static void main(String[] args) throws InvalidFormatException, IOException {

	File fi = new File("C:\\Users\\Nitish\\Desktop\\Workspace\\Learning Tracker\\Input.xlsx");
	FileInputStream fil = new FileInputStream(fi); // File Input Stream open
	XSSFWorkbook xwb = new XSSFWorkbook(fil); // Creating Workbook
	XSSFSheet xsh = xwb.getSheetAt(0); // Creating sheet
	
	int r = xsh.getPhysicalNumberOfRows();
	int c = xsh.getRow(0).getPhysicalNumberOfCells();
	System.out.println("Total number of rows are:"+r);
	System.out.println("Total number of cols are:"+ c);
	for (int i=0;i<r;i++) // Outer loop for row
	{
		XSSFRow xrw = xsh.getRow(i);
		for (int j=0;j<xrw.getPhysicalNumberOfCells();j++)
		{
			XSSFCell xcl = xrw.getCell(j);
			System.out.print(xcl);
			System.out.print(" ");
		}
		System.out.println("");
	}

	
	
	
	}

}
