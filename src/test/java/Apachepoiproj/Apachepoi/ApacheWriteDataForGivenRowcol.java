package Apachepoiproj.Apachepoi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheWriteDataForGivenRowcol {


		public void WriteDataForGivenRowRange(int row, int coln) throws IOException
		{
			File f=new File("../Apachepoi/abc.xlsx");
			FileOutputStream fos=new FileOutputStream(f);
			XSSFWorkbook xw=new XSSFWorkbook();
			XSSFSheet xs=xw.createSheet("MSS");
			String st;
			System.out.println("Please enter data");
			Scanner s=new Scanner(System.in);
			
			for(int i=0;i<row;i++)
			{
				XSSFRow xr=xs.createRow(i);
				for(int j=0;j<coln;j++)
				{
					    st=s.next();
						XSSFCell xc=xr.createCell(j);
						xc.setCellValue(st);
						System.out.println("data writing done");
	
					}
			}		
			xw.write(fos);
			fos.close();
			fos.flush();	
				
			}	

		public static void main(String[] args) throws IOException 
		{
	        Scanner s1= new Scanner(System.in);
			System.out.println("Please enter no of row ");
			int c=s1.nextInt();
			System.out.println("Please enter no of column");
			int d=s1.nextInt();
			ApacheWriteDataForGivenRowcol ad=new ApacheWriteDataForGivenRowcol();
			ad.WriteDataForGivenRowRange(c,d);
	}

}
