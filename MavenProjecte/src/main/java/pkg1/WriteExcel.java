package pkg1;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {

		File fi = new File("C:\\Users\\Nitish\\Desktop\\Workspace\\Learning Tracker\\Output.xlsx");
		FileOutputStream fio = new FileOutputStream(fi);
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sh = wb.createSheet("Sheet 1");
		
		for (int i=0;i<3;i++)
		{
			XSSFRow rw = sh.createRow(i);
			for(int j=0;j<3;j++)
			{
				XSSFCell xc =rw.createCell(j);
				xc.setCellValue("Cell no "+j);
			}		
		}
		wb.write(fio);
		fio.flush();
		fio.close();
		
		System.out.println("Output Generated !!!!!!");
		
		
		
	}

}
