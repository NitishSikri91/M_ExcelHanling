package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Temp {

	public void ReadCell(int i , int j, FileInputStream fir , File fi ) throws IOException
	{
		fir = new FileInputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook(fir);
		XSSFSheet sh = wb.getSheetAt(0);
		System.out.println("Row num in "+ i);
		System.out.println("Coln num is "+ j);
		XSSFRow rw =sh.getRow(i);
		XSSFCell cl = rw.getCell(j);
		System.out.println("Value in cell is "+ cl);
		
	}
	
	public static void main(String[] args) throws IOException {
		File fi = new File("C:\\Users\\Nitish\\Desktop\\Workspace\\Learning Tracker\\Input.xlsx");
		FileInputStream fir = new FileInputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook(fir);
		XSSFSheet sh = wb.getSheetAt(0);
		Scanner scan = new Scanner(System.in);
		Temp obj = new Temp();
		int r=sh.getPhysicalNumberOfRows();
		int c=sh.getRow(0).getPhysicalNumberOfCells();
		
		System.out.println("Total Number of Rows "+r);
		System.out.println("Total Number of Columns "+c);
		System.out.println("Please enter the row you want to read");
		
		int a = scan.nextInt();
		if (a>r || a<0)
		{
			System.out.println("Please enter valid number of row");
			a=0;
			return;
		}
			
		System.out.println("Please enter the column you want to read");
		int b = scan.nextInt();
		if (b>c || b<0)
		{	System.out.println("Please enter valid number of row");
			b=0;
			return;
		}
		
				obj.ReadCell(a, b , fir,fi);
	}


	
}
