package Apachepoiproj.Apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheReaddataForGivenRange {
	public void ReadDataForGivenRowRange(int startrow, int endrow) throws IOException
	{
		File f=new File("../Apachepoi/apchedeom.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xw=new XSSFWorkbook(fis);
		XSSFSheet xs=xw.getSheetAt(0);
		int a=xs.getPhysicalNumberOfRows();
		System.out.println(a);
		for(int i=1;i<a;i++)//excluding heading row
		{
			XSSFRow xr=xs.getRow(i);
			int b=xr.getPhysicalNumberOfCells();
			for(int j=0;j<b;j++)
			{
				if(i<=endrow)
				{
					XSSFCell xc=xr.getCell(j);
					System.out.print( xc.getStringCellValue());
				
				}
				}
			
			
		}		
		
	}

	public static void main(String[] args) throws IOException 
	{
        Scanner s = new Scanner(System.in);
		System.out.println("Please enter start row no");
		int c=s.nextInt();
		System.out.println("Please enter end row no");
		int d=s.nextInt();
		ApacheReaddataForGivenRange ad=new ApacheReaddataForGivenRange();
		ad.ReadDataForGivenRowRange(c,d);
		

	}
}


