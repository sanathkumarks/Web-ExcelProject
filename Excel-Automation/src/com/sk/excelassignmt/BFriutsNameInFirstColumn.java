package com.sk.excelassignmt;
import java.util.Scanner;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;;
public class BFriutsNameInFirstColumn {

	public static void main(String[] args)
	{
		writeContent();
	}
	static void writeContent()
	{
		Scanner sc=new Scanner(System.in);
		FileOutputStream fout= null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try
		{
			System.out.println("enter number of rows should be created");
			int n=sc.nextInt();
			wb=new XSSFWorkbook();
			sh=wb.createSheet("Fruits1");
			for(int i=0;i<n;i++)
			{
				row=sh.createRow(i);
				cell=row.createCell(0);
				System.out.println("Enter fruits name");
				String fruitsname=sc.next();
				cell.setCellValue(fruitsname);
			}
			fout=new FileOutputStream("D:\\EXCEL\\sks.xlsx");
			wb.write(fout);
			System.out.println("Successfully created excel sheet for Fruits");
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
