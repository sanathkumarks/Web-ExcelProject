package com.sk.excelread;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcelFileContent {

	public static void main(String[] args) {
		readWriteContent();

	}
	
	static void readWriteContent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowSh1=null;
		Row rowSh2=null;
		Cell cellSh1=null;
		Cell cellSh2=null;
		try
		{
			fin=new FileInputStream("D:\\EXCEL\\Test.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
			
			int rc=sh1.getPhysicalNumberOfRows();
			for(int r=0;r<rc;r++)
			{
				rowSh1=sh1.getRow(r);
				rowSh2=sh2.getRow(r);
				if(rowSh2==null)
				{
					rowSh2=sh2.createRow(r);
				}
				
				int cc=rowSh1.getPhysicalNumberOfCells();
				for(int c=0;c<cc;c++)
				{
					cellSh1=rowSh1.getCell(c);
					cellSh2=rowSh2.getCell(c);
					if(cellSh2==null)
					{
						cellSh2=rowSh2.createCell(c);
					}
					String data=cellSh1.getStringCellValue();
					cellSh2.setCellValue(data);
				}
				fout=new FileOutputStream("D:\\EXCEL\\Test.xlsx");
				wb.write(fout);
			}
		}catch(Exception e)
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
				rowSh1=null;
				rowSh2=null;
				cellSh1=null;
				cellSh2=null;
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}

}

