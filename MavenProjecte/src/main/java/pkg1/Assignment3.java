package pkg1;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment3 {


	public static void main(String[] args) throws IOException {
		
		File fi = new File("C:\\Users\\Nitish\\Desktop\\Workspace\\Learning Tracker\\Output2.xlsx");
		FileOutputStream fiw = new FileOutputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sh = wb.createSheet("Sheet 1");
		Scanner scan = new Scanner(System.in);
		for(int i =0;i<3;i++)
		{
			XSSFRow rw = sh.createRow(i);
			for(int j=0;j<3;j++)
			{
				XSSFCell cl = rw.createCell(j);
				System.out.println("Please enter a single string value");
				String str = scan.next();
				cl.setCellValue(str);
			}
			
		}
		
		System.out.println("Output Generated");
		wb.write(fiw);
		wb.close();
		fiw.flush();
		fiw.close();
	
		
		
	}
	

}
