package com.sk.excelassignmt;
import java.util.Scanner;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
public class AMonthNamesDiagonally {

	public static void main(String[] args) 
	{
		writeMonthNmes();
	}
	static void writeMonthNmes()
	{
		Scanner sc=new Scanner(System.in);
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try
		{
			wb=new XSSFWorkbook();
			sh=wb.createSheet("Months");
			System.out.println("Enter Number of months");
			int n=sc.nextInt();
			for(int i=0;i<n;i++)
			{
				row=sh.createRow(i);
				cell=row.createCell(i);
				System.out.println("Enter Month");
				String month=sc.next();
				cell.setCellValue(month);
			}
			fout=new FileOutputStream("D:\\EXCEL\\month.xlsx");
			wb.write(fout);
			System.out.println("Excel Sheet Created Successfully");
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
				cell=null;
				row=null;
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}

}
