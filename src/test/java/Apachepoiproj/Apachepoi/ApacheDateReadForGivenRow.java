package Apachepoiproj.Apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheDateReadForGivenRow {
	public void ReadDataForGivenRow(int rowno) throws IOException
	{
		File f=new File("../Apachepoi/apchedeom.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xw=new XSSFWorkbook(fis);
		XSSFSheet xs=xw.getSheetAt(0);
		int a=xs.getPhysicalNumberOfRows();
		
		for(int i=0;i<a;i++)
		{
			XSSFRow xr=xs.getRow(i);
			int b=xr.getPhysicalNumberOfCells();
			for(int j=0;j<b;j++)
			{
				if(i==rowno)
				{
					XSSFCell xc=xr.getCell(j);
					System.out.println("Cell value " + xc.getStringCellValue());
				
				}
				}
			
			
		}		
		
	}
	public static void main(String[] args) throws IOException {
    try (Scanner s = new Scanner(System.in)) {
		System.out.println("Please enter row no");
		int c=s.nextInt();
		ApacheDateReadForGivenRow ad=new ApacheDateReadForGivenRow();
		ad.ReadDataForGivenRow(c);
	}


	}

}
