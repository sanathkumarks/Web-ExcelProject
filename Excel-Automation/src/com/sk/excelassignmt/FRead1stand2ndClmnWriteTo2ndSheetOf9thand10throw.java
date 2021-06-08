package com.sk.excelassignmt;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FRead1stand2ndClmnWriteTo2ndSheetOf9thand10throw 
{
	public static void main(String[] args)
	{
		readWritecontent();
	}
	static void readWritecontent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowsh1=null;
		Row rowsh2=null;
		Row rowsh3=null;
		Cell cellsh1=null;
		Cell cellsh2=null;
		Cell cellsh3=null;
		try
		{
			fin =new FileInputStream("D:\\EXCEL\\Flower.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
			int r=sh1.getPhysicalNumberOfRows();
			rowsh2=sh2.createRow(9);
			rowsh3=sh2.createRow(10);
			int k=0;
			int l=0;
			for(int i=0;i<r;i++)
			{
				rowsh1=sh1.getRow(i);
				/*rowsh2=sh2.getRow(i);
				if(rowsh2==null) 
				{
					rowsh2=sh2.createRow(i);
				}*/
				int c=rowsh1.getPhysicalNumberOfCells();
				for(int j=0;j<c;j++)
				{
					cellsh1=rowsh1.getCell(j);
					if(j==0)
					{
						cellsh2=rowsh2.getCell(k);
						if(cellsh2==null)
						{
							cellsh2=rowsh2.createCell(k);
						}
						String data=cellsh1.getStringCellValue();
						cellsh2.setCellValue(data);
						k++;
					}
					else if(j==1)
					{
						cellsh3=rowsh3.getCell(l);
						if(cellsh3==null)
						{
							cellsh3=rowsh3.createCell(l);
						}
						String data=cellsh1.getStringCellValue();
						cellsh3.setCellValue(data);
						l++;
					}
				}
				fout=new FileOutputStream("D:\\EXCEL\\Flower.xlsx");
				wb.write(fout);
			}
			System.out.println("Successfull written data to 9th and 10th row of the 2nd sheet!!!");
			
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
