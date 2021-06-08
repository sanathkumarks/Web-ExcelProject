package com.sk.excelassignmt;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GReadWriteReverseOrder
{
	public static void main(String[] args) 
	{
		readWriteContent();
	}
	static void readWriteContent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowsh1=null;
		Row rowsh2=null;
		Cell cellsh1=null;
		Cell cellsh2=null;
		try
		{
			fin = new FileInputStream("D:\\EXCEL\\ReverseOrder.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
				
			}
			int r=sh1.getPhysicalNumberOfRows();
			int k=1;
			for(int i=r-1;i>=0;i--)
			{
				rowsh1=sh1.getRow(i);
				rowsh2=sh2.getRow(k);
				if(rowsh2==null)
				{
					rowsh2=sh2.createRow(k);
				}
				int c=rowsh1.getPhysicalNumberOfCells();
				for(int j=0;j<c;j++)
				{
					cellsh1=rowsh1.getCell(j);
					cellsh2=rowsh2.getCell(j);
					if(cellsh2==null)
					{
						cellsh2=rowsh2.createCell(j);
					}
					String value=cellsh1.getStringCellValue();
					cellsh2.setCellValue(value);
				}
				k++;
				fout=new FileOutputStream("D:\\EXCEL\\ReverseOrder.xlsx");
				wb.write(fout);
			}
			System.out.println("successfully data written to sheet2");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fin.close();
				fout.close();
				wb.close();
				sh1=null;
				sh2=null;
				rowsh1=null;
				rowsh2=null;
				cellsh1=null;
				cellsh2=null;
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
		}
		
	}
}
