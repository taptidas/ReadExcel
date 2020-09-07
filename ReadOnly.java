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

public class ReadOnly {
	
	public static void main(String[] args)   
	{  
		ReadOnly rc=new ReadOnly();
	try  
	{  
	File file = new File("E:\\employee.xlsx");   //creating a new file instance  
	FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
	//creating Workbook instance that refers to .xlsx file  
	XSSFWorkbook wb = new XSSFWorkbook(fis);  
	//fis.close();
	XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
	Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
	OutputStream os = new FileOutputStream("E:\\incentivcal.xlsx"); 
	int r=0;
	while (itr.hasNext())                 
	{  	Row row = itr.next();  
	//Row row1 = sheet.createRow(r);  
	r++;
	//Cell cell = row.createCell(j);
	//cell.setCellValue(rc.ReadCellData(i, j));
	Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column 
	int count=0;
	while (cellIterator.hasNext())   
	{  count++;
	Cell cell = cellIterator.next();  
	//Cell cell1 = row1.createCell(count);
	switch (cell.getCellType())               
	{  
	case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
	System.out.print(cell.getStringCellValue() + "\t\t\t");  
	cell.setCellValue(cell.getStringCellValue());
	break;  
	case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
	System.out.print(cell.getNumericCellValue() + "\t\t\t");  
	cell.setCellValue(cell.getNumericCellValue());
	break;  
	default:  
	}  
	}  
	System.out.println(""); 
	}System.out.println("row:"+sheet.getLastRowNum());

/**  XSSFWorkbook wb1 = new XSSFWorkbook("E:\\incentivcal.xlsx"); 
  XSSFSheet sheet1 = wb.getSheetAt(0);  	double incentive=0;
	for(int i=0;i<=sheet.getLastRowNum();i++) {

		Row row1 = sheet.createRow(i);  
		Cell cell1 = row1.createCell(5); 

		if(i==0)
		{
			cell1.setCellValue("Incentive");
		}
		else
		{	if(rc.ReadCellData(i, 3)==1)
			incentive =((95*rc.ReadCellData(i, 4))/100);
		else if(rc.ReadCellData(i, 3)==2)
			incentive=((65*rc.ReadCellData(i, 4))/100);
		else if(rc.ReadCellData(i, 3)==3)
			incentive=(35*rc.ReadCellData(i, 4))/100; 
		cell1.setCellValue(incentive); }


	}**/

	  wb.write(os);
	os.close();
	OutputStream os1 = new FileOutputStream("E:\\incentivcal.xlsx"); 
	//XSSFWorkbook wb1 = new XSSFWorkbook("E:\\incentivcal.xlsx"); 
	 /// XSSFSheet sheet1 = wb.getSheetAt(0);  

		double incentive=0;
	for(int i=0;i<=sheet.getLastRowNum();i++) {

		Row row1 = sheet.createRow(i);  
		Cell cell1 = row1.createCell(8); 

		if(i==0)
		{
			cell1.setCellValue("Incentive");
		}
		else
		{	if(rc.ReadCellData(i, 3)==1)
			incentive =((95*rc.ReadCellData(i, 4))/100);
		else if(rc.ReadCellData(i, 3)==2)
			incentive=((65*rc.ReadCellData(i, 4))/100);
		else if(rc.ReadCellData(i, 3)==3)
			incentive=(35*rc.ReadCellData(i, 4))/100; 
		cell1.setCellValue(incentive); }


	}wb.write(os1);
	os1.close();
	
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

	

