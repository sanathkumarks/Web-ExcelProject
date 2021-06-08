package com.sk.excelread;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFileContent {

	public static void main(String[] args) {
		writeContent();

	}
	
	static void writeContent()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try 
		{
			wb=new XSSFWorkbook();
			sh=wb.createSheet("Information");
			row=sh.createRow(0);
			cell=row.createCell(0);
			cell.setCellValue("UserName");
			cell=row.createCell(1);
			cell.setCellValue("Password");
			
			row=sh.createRow(1);
			cell=row.createCell(0);
			cell.setCellValue("DemoUser1");
			cell=row.createCell(1);
			cell.setCellValue("Welcome123");
			
			fout=new FileOutputStream("D:\\EXCEL\\Welcome.xlsx");
			wb.write(fout);
		}catch(Exception e)
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
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}

}