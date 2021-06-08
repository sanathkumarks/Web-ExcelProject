package com.sk.excelassignmt;
import java.util.Scanner;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class DCityNameIn5thColumn 
{
	public static void main(String[] args) 
	{
		 writeCityName();
	}
	static void writeCityName()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try
		{
			Scanner sc=new Scanner(System.in);
			wb=new XSSFWorkbook();
			sh=wb.createSheet("CityName");
			System.out.println("Enter number of city names you want insert into sheet");
			int n=sc.nextInt();
			for(int i=0;i<n;i++)
			{
				row=sh.createRow(i);
				cell=row.createCell(5-1);
				System.out.println("Enter city name");
				String cityname=sc.next();
				cell.setCellValue(cityname);
			}
			fout=new FileOutputStream("D:\\EXCEL\\cityName.xlsx");
			wb.write(fout);
			System.out.println("Successfull created Excel sheet for citynames");
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
