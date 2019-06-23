package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AssignmentTwo {

	public void Readline(int a, File fi) throws IOException
	{	FileInputStream fip = new FileInputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook(fip);
		XSSFSheet sh = wb.getSheetAt(0);
		XSSFRow rw = sh.getRow(a);
		for (int x=0;x<rw.getPhysicalNumberOfCells();x++)
		{
			XSSFCell cl = rw.getCell(x);
			System.out.print(cl+" ");
		}
	}
	
	public static void main(String[] args) throws IOException {
		File fi = new File("C:\\Users\\Nitish\\Desktop\\Workspace\\Learning Tracker\\Input.xlsx");
		FileInputStream fip = new FileInputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook(fip);
		XSSFSheet sh = wb.getSheetAt(0);
		Scanner scan = new Scanner(System.in);
		AssignmentTwo obj = new AssignmentTwo();
		int a=0;
		int rows = sh.getPhysicalNumberOfRows();
		int cols = sh.getRow(0).getPhysicalNumberOfCells();
		
		System.out.println("No of rows are :"+rows);
		System.out.println("No of cols are :"+cols);
		
		System.out.println("Please enter the row you want to read.");
		if(scan.hasNextInt())
			{ a =scan.nextInt();}
		else 
			System.out.println("Please enter a number value");
	
		if (a>rows || a<0)
		{
			System.out.println("Please enter valid value");
		}
	
		
		obj.Readline(a-1,fi);
		
	}

}
