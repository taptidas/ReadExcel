package xslsx;

import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;


import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class ReadExcel {

	public static void main(String[] args)   
	{ 
		ReadExcel rc=new ReadExcel();  
		try  
		{  
			XSSFWorkbook wb = new XSSFWorkbook("E:\\employee.xlsx");
			OutputStream os = new FileOutputStream("E:\\incentivcal.xlsx"); 


			double incentive = 0;
			XSSFSheet sheet = wb.getSheetAt(0); 
			System.out.println(wb.getSheetAt(0).getRow(3).getCell(4)); 
			System.out.println(sheet.getRow(0).getPhysicalNumberOfCells());			
			for(int i=0;i<=sheet.getLastRowNum();i++) {

				Row row = sheet.createRow(i);  
				Cell cell = row.createCell(5); 

				if(i==0)
				{
					cell.setCellValue("Incentive");
				}
				else
				{	if(rc.ReadCellData(i, 3)==1)
					incentive =((95*rc.ReadCellData(i, 4))/100);
				else if(rc.ReadCellData(i, 3)==2)
					incentive=((65*rc.ReadCellData(i, 4))/100);
				else if(rc.ReadCellData(i, 3)==3)
					incentive=(35*rc.ReadCellData(i, 4))/100; 
				cell.setCellValue(incentive); }


			}wb.write(os); 
			os.close();

			System.out.println("rows:"+sheet.getLastRowNum());
			System.out.println("columns:"+sheet.getRow(0).getPhysicalNumberOfCells());

		} 
		catch(Exception e)  
		{  
			e.printStackTrace();  
		}  

	}


	public double ReadCellData(int vRow, int vColumn)  
	{  
		double value;            
		Workbook wb=null;            
		try  
		{  

			FileInputStream fis=new FileInputStream("E:\\employee.xlsx");  

			wb=new XSSFWorkbook(fis); 
			fis.close();
		}  
		catch(FileNotFoundException e)  
		{  
			e.printStackTrace();  
		}  
		catch(IOException e1)  
		{  
			e1.printStackTrace();  
		}  
		Sheet sheet=wb.getSheetAt(0);   //getting the XSSFSheet object at given index  
		Row row=sheet.getRow(vRow); //returns the logical row  
		Cell cell=row.getCell(vColumn); //getting the cell representing the given column  
		value=cell.getNumericCellValue();   //getting cell value  
		return value;               //returns the cell value  

	}  
}  

