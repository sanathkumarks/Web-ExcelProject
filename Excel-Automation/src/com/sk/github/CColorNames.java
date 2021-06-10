package com.sk.github;
import java.util.Scanner;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class CColorNames 
{
	public static void main(String[] args) 
	{
		writeColor();
	}
	static void writeColor()
	{
		Scanner sc=new Scanner(System.in);
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try
		{
			
			System.out.println("Enter number of Color you insert into sheet");
			int n=sc.nextInt();
			wb=new XSSFWorkbook();
			sh=wb.createSheet("Colors");
			row=sh.createRow(10-1);
			for(int i=0;i<n;i++)
			{
				
				cell=row.createCell(i);
				System.out.println("Enter color");
				String color=sc.next();
				cell.setCellValue(color);
			}
			fout=new FileOutputStream("D:\\EXCEL\\color1.xlsx");
			wb.write(fout);
			System.out.println("Successfully created Excel sheet for colors ");
			sc.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fout.close();
				wb.close();
				sh=null;
				row=null;
				cell=null;
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}
}
