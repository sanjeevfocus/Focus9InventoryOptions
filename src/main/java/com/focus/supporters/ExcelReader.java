package com.focus.supporters;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import javax.sound.midi.SysexMessage;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.apache.poi.xwpf.usermodel.Borders;
import org.testng.Assert;

import com.focus.base.BaseEngine;

//import com.focus.base.BaseEngine;

public class ExcelReader extends BaseEngine
{
	private String filePath;
	private FileInputStream fip;
	private FileOutputStream fop;
	private Workbook workbook;
	private Sheet sheet;
	private Cell cell;
	private CellStyle style;
	private Row row;

	
	public ExcelReader(String filePath) throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		this.filePath=filePath;
		fip=new FileInputStream(filePath);
		workbook=WorkbookFactory.create(fip);
	}
	

	public Sheet getSheet(String sheetName) 
	{
		sheet = workbook.getSheet(sheetName);
		return sheet;
	}
	
	public Sheet getSheet(int sheetIndex)
	{
		sheet=workbook.getSheetAt(sheetIndex);
		return sheet;
	}
	
	public Row getRow(String sheetName, int rowNum)
	{
		Row row=getSheet(sheetName).getRow(rowNum);
		return row;
	}
	
	public Row getRow(int sheetIndex, int rowNum)
	{
		row=getSheet(sheetIndex).getRow(rowNum);
		return row;
	}
	
	public Cell getCell(String sheetName,int rowNum, int cellNum)
	{
    
		cell =getSheet(sheetName).getRow(rowNum).getCell(cellNum);
		return cell;
		
	}
	
	
	
	public String getCellData(String sheetName, int rowNum,int cellNum)
	{
		String data=null;
		cell=getSheet(sheetName).getRow(rowNum).getCell(cellNum);
		if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
		{
			data=cell.getNumericCellValue()+"";
		}
		else if(cell.getCellType()==Cell.CELL_TYPE_STRING)
		{
			data=cell.getStringCellValue();
		}
		return data;
	}
	
	
   public ArrayList<String> getTotalExcelData(String xlfile,String sheetName) throws IOException
   {
	   FileInputStream fi = new FileInputStream(xlfile);
		XSSFWorkbook wb = new XSSFWorkbook(fi);
	   
	   String data=null;
	   
	   sheet=wb.getSheet(sheetName);
	   
	   System.err.println("ROWS : "+sheet.getLastRowNum());
	   
	   ArrayList<String> excelData=new ArrayList<String>();
	   
	   for(int i=0;i<sheet.getLastRowNum();i++)
	   {
		   row=sheet.getRow(i);
		   
		   System.err.println("Row : "+i+" Column "+row.getLastCellNum());
		   
		   for(int j=0;j<row.getLastCellNum();j++)
		   {
			   System.err.println("Entered row and getting data");
			   
			   
			   if (cell != null) 
			   {
				   cell=row.getCell(j);
				   System.err.println("CELL DATA GETTING");
				   
				   if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
				   {
					   data=cell.getNumericCellValue()+"";
					   excelData.add(data);
				   }
				   else if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				   {
					   data=cell.getStringCellValue();
					   excelData.add(data);
				   }
				   else if(cell.getCellType()==cell.CELL_TYPE_BOOLEAN)
				   {
					   boolean text = cell.getBooleanCellValue();
					   data=Boolean.toString(text);
					   excelData.add(data);
				   }
				 
			   }
			   
		   }
	   }
	   System.err.println("EXCEL DATA :"+excelData);
	   return excelData;
   }
   
  
   
   public   ArrayList<String[]> readContentFromExcel(String xlfile,String sheetName) throws IOException
   {
	    FileInputStream file = new FileInputStream(xlfile);
	    ArrayList<String[]> ExcelRows = new ArrayList<String[]>();
	    XSSFWorkbook  wb = new XSSFWorkbook(file);
	    sheet=wb.getSheet(sheetName);
	   int rowCount = sheet.getLastRowNum();
	   
	   int columnCount = row.getLastCellNum();
	    
	   for(int i=0;i<sheet.getLastRowNum();i++)
	   {
		   row=sheet.getRow(i);
		   
		   System.err.println("Row : "+i+" Column "+row.getLastCellNum());
		   
		   for(int j=0;j<row.getLastCellNum();j++)
		   {
			   System.err.println("Entered row and getting data");
	    
	    /*Iterator<Row> rowIterator = sheet.iterator();*/
	    int cellIndex = 0;
	   /* while (rowIterator.hasNext())
	    {
	      Row row = rowIterator.next();*/
	      Iterator <Cell> cellIterator = row.cellIterator();
	      String[] cellValues  = new String[columnCount];
	      while (cellIterator.hasNext()) 
	      {
	           Cell cell = cellIterator.next();
	           if (cell != null) 
			   {
				   
				   if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
				   {
					   int text = (int) cell.getNumericCellValue();
					   cellValues[cellIndex++] = Integer.toString(text);
			           System.out.print(cell.getStringCellValue() + "\t\t");
				   }
				   else if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				   {
					   cellValues[cellIndex++] = cell.getStringCellValue();
			           System.out.print(cell.getStringCellValue() + "\t\t");
				   }
				   else if(cell.getCellType()==cell.CELL_TYPE_BOOLEAN)
				   {
					   boolean text = cell.getBooleanCellValue();
					   cellValues[cellIndex++] = Boolean.toString(text);
			           System.out.print(cell.getStringCellValue() + "\t\t");
				   }
				 
			   }
	      }
	     ExcelRows.add(cellValues);
	     cellIndex=0;
	   }
	   }

	  file.close();
	  return ExcelRows;
	}
	   
   
   
   
   public List<String> getRowData(String sheetName, int rowNum)
   {
	   row=getSheet(sheetName).getRow(rowNum);
	   List<String> excelRowData=null;
	   String data=null;
	   
	   for(int i=0;i<row.getLastCellNum();i++)
	   {
		   cell=row.getCell(i);
		   
		   if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
		   {
			   data=cell.getNumericCellValue()+"";
			   excelRowData.add(data);
		   }
		   else if(cell.getCellType()==Cell.CELL_TYPE_STRING)
		   {
			   data=cell.getStringCellValue();
			   excelRowData.add(data);
		   }   
		   else if (cell.getCellType()==cell.CELL_TYPE_BOOLEAN) 
		   {
			   boolean value=cell.getBooleanCellValue();
			   data=Boolean.toString(value);
			   excelRowData.add(data);
			}
	   }
	   return excelRowData;
   }
   
   
   //To Write the data in the excel sheet
  
   public void setCellData(String filePath,String sheetName,int rowNum,int cellNum,String data) throws IOException, EncryptedDocumentException, InvalidFormatException
   {
       if(data.equalsIgnoreCase("pass") || data.equalsIgnoreCase("fail") )
       {
           if(data.equalsIgnoreCase("pass") )
           {
               this.filePath       =  filePath;
               FileInputStream fip = new FileInputStream(filePath);             
               workbook            = WorkbookFactory.create(fip);             
               Sheet sheet         = workbook.getSheet(sheetName);           
               Row row             = sheet.getRow(rowNum);           
               Cell cell           = row.createCell(cellNum);
               
               cell.setCellValue(data);
              
               style               = workbook.createCellStyle();
               style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
               style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                
               cell.setCellStyle(style);
    
               FileOutputStream fop = new FileOutputStream(filePath);
    
               workbook.write(fop);
               
               fip.close();
               
               fop.close();
           }
           else
           {
               this.filePath        = filePath;
               FileInputStream fip  = new FileInputStream(filePath);
               workbook             = WorkbookFactory.create(fip);
               Sheet sheet          = workbook.getSheet(sheetName);
               Row row              = sheet.getRow(rowNum);
               Cell cell            = row.createCell(cellNum);
               
               cell.setCellValue(data);
               
               style                = workbook.createCellStyle();
               style.setFillForegroundColor(IndexedColors.RED.getIndex());
               style.setFillPattern(FillPatternType.BIG_SPOTS);
                
               cell.setCellStyle(style);

               FileOutputStream fop = new FileOutputStream(filePath);

               workbook.write(fop);
               
               fip.close();
               
               fop.close();
           }
       }
       else
       {
           this.filePath            = filePath;
           FileInputStream fip      = new FileInputStream(filePath);
           workbook                 = WorkbookFactory.create(fip); 
           Sheet sheet              = workbook.getSheet(sheetName);
           Row row                  = sheet.getRow(rowNum);          
           Cell cell                =row.getCell(cellNum);
        
           cell.setCellValue(data);
           
           FileOutputStream fop     = new FileOutputStream(filePath);

           workbook.write(fop);
           
           fip.close();
           
           fop.close();
       }
     }
   
   
   public void copyExcelFile() throws IOException, EncryptedDocumentException, InvalidFormatException, InterruptedException
   {
   	   //Provide the Path of excel file which we want to copy
   	   File inputFile=new File(getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestDataExcelBackup\\FocusTest.xlsx");
   	   File outputFile=new File(getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx");
   	   
   	   if(outputFile.exists())
   	   {
   	     outputFile.delete();
   	     System.out.println("The Backup File Previous Exits and Delete The Backup Test Data File");
   	   }
   	   
   	   Thread.sleep(3000);
   	   Files.copy(inputFile.toPath(), outputFile.toPath());
 	   System.out.println("Test Data Back Up File Copied Successfully");
   }
   
   
   
   public void setExceptionInExcel(String filePath,String sheetName,int rowNum,int cellNum,String data) throws IOException, EncryptedDocumentException, InvalidFormatException
   {
	   this.filePath=filePath;

	   FileInputStream fip=new FileInputStream(filePath);
	   workbook =WorkbookFactory.create(fip);
	   
	   Sheet sheet=workbook.getSheet(sheetName);
	   
	   Row row=sheet.getRow(rowNum);

	   Cell cell =row.createCell(cellNum);
	   cell.setCellValue(data);
	   
	   style = workbook.createCellStyle();
	   style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	   style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	   
	   cell.setCellStyle(style);
			
	   FileOutputStream fop=new FileOutputStream(filePath);
			
	   workbook.write(fop);
	   
	   fip.close();
	   
	   fop.close();
   }
   
  
   
   
   /*
	This method compares if both excel files have same number of sheets and same
	name of sheets
	*/
	public boolean verifyIfExcelFilesHaveSameNumberAndNameOfSheets(Workbook workbook1, Workbook workbook2) {
		System.out.println("Verifying if both work books have same number of sheets.............");
		// Get total sheets count from first excel file
		int sheetsInWorkbook1 = workbook1.getNumberOfSheets();
		// Get total sheets count from second excel file
		int sheetsInWorkbook2 = workbook2.getNumberOfSheets();
		// Compare if both excel files have same number of sheets
		Assert.assertEquals(sheetsInWorkbook1, sheetsInWorkbook2,
				"Excel work books have different number of sheets. \n "
				+ "Sheets in work book 1 : "+sheetsInWorkbook1 +"\n "
						+ "Number of sheets in work book 2 : "+sheetsInWorkbook2);
		// Printing number of sheets in each work book
		System.out.println("Sheets in first work book : "+sheetsInWorkbook1);
		System.out.println("Sheets in second work book : "+sheetsInWorkbook2);
		System.out.println("Both work books have same number of sheets.........................");
		
		// Verify if sheets have same name in both workbooks
		// Sheets may not be in same order in both excel. So I am relaxing order of sheets condition.
		// Change i as required.
		System.out.println("Verifying if both work books have same name of sheets.............");
		List<String> sheetsNameOfWb1 = new ArrayList<>();
		List<String> sheetsNameOfWb2 = new ArrayList<>();
		// Since we have already verified that both work books have same number of sheets so iteration can be done against any workbook sheet count
		for (int i = 0; i < sheetsInWorkbook1; i++) {
			// Retrieving sheet names from both work books and adding to different lists
			sheetsNameOfWb1.add(workbook1.getSheetName(i));
			sheetsNameOfWb2.add(workbook2.getSheetName(i));
		}
		// Since I am relaxing same sequence of sheets. 
		Collections.sort(sheetsNameOfWb1);
		Collections.sort(sheetsNameOfWb2);
		boolean res = sheetsNameOfWb1.equals(sheetsNameOfWb2);
		System.out.println("Sheet Names in first work book : "+sheetsNameOfWb1);
		System.out.println("Sheet Names in second work book : "+sheetsNameOfWb2);
		System.out.println("Both work books have same name of sheets.........................");
		return res;
	}

	// This method compares if both excel files have same number of rows and corresponding columns
	public void verifySheetsInExcelFilesHaveSameRowsAndColumns(Workbook workbook1, Workbook workbook2) {
		System.out.println(
				"Verifying if both work books have same number of rows and columns in all sheets.............");
		int sheetCounts = workbook1.getNumberOfSheets();
		for (int i = 0; i < sheetCounts; i++) {
			Sheet s1 = workbook1.getSheetAt(i);
			Sheet s2 = workbook2.getSheetAt(i);
			int rowsInSheet1 = s1.getPhysicalNumberOfRows();
			int rowsInSheet2 = s2.getPhysicalNumberOfRows();
			Assert.assertEquals(rowsInSheet1, rowsInSheet2, "Sheets have different count of rows..");

			Iterator<Row> rowInSheet1 = s1.rowIterator();
			Iterator<Row> rowInSheet2 = s2.rowIterator();
			while (rowInSheet1.hasNext()) {
				int cellCounts1 = rowInSheet1.next().getPhysicalNumberOfCells();
				int cellCounts2 = rowInSheet2.next().getPhysicalNumberOfCells();
				Assert.assertEquals(cellCounts1, cellCounts2, "Sheets have different count of columns..");
			}
		}
	}
	boolean res =false;
	
	public boolean verifyDataInExcelBookAllSheets(Workbook workbook1, Workbook workbook2)
	{
		
		System.out.println("Verifying if both work books have same data.............");
		// Since we have already verified that both work books have same number of sheets so iteration can be done against any workbook sheet count
		int sheetCounts = workbook1.getNumberOfSheets();
		// So we will iterate through sheet by sheet
		for (int i = 0; i < sheetCounts; i++) {
			// Get sheet at same index of both work books
			Sheet s1 = workbook1.getSheetAt(i);
			Sheet s2 = workbook2.getSheetAt(i);
			System.out.println("*********** Sheet Name : "+s1.getSheetName()+"*************");
			// Iterating through each row
			int rowCounts = s1.getPhysicalNumberOfRows();
			System.err.println("Row Coun : "+rowCounts);
			for (int j = 0; j < rowCounts; j++) {
				// Iterating through each cell
				int cellCounts = s1.getRow(j).getPhysicalNumberOfCells();
				for (int k = 0; k < cellCounts; k++) {
					// Getting individual cell
					Cell c1 = s1.getRow(j).getCell(k);
					Cell c2 = s2.getRow(j).getCell(k);
					
					
					
						boolean a = c1.getCellType()==cell.CELL_TYPE_BLANK  ;
						
					 if (a==false) 
					 {
						   
					// Since cell have types and need o use different methods
					if (c1.getCellType()==(c2.getCellType())) {
						if (c1.getCellType() == cell.CELL_TYPE_STRING) {
							String v1 = c1.getStringCellValue();
							String v2 = c2.getStringCellValue();
							res = v1.equalsIgnoreCase(v2);
							System.out.println("Its matched : "+ v1 + " === "+ v2);
						}
						if (c1.getCellType() == cell.CELL_TYPE_NUMERIC) 
						{
							// If cell type is numeric, we need to check if data is of Date type
							if (DateUtil.isCellDateFormatted(c1) | DateUtil.isCellDateFormatted(c2)) 
							{
								// Need to use DataFormatter to get data in given style otherwise it will come as time stamp
								DataFormatter df = new DataFormatter();
								String v1 = df.formatCellValue(c1);
								String v2 = df.formatCellValue(c2);
								res = v1.equalsIgnoreCase(v2);
								System.out.println("Its matched : "+ v1 + " === "+ v2);
							} 
							else
							{
								double v1 = c1.getNumericCellValue();
								double v2 = c2.getNumericCellValue();
								res = v1==v2;
								System.out.println("Its matched : "+ v1 + " === "+ v2);
							}
						}
						if (c1.getCellType() == cell.CELL_TYPE_BOOLEAN) {
							boolean v1 = c1.getBooleanCellValue();
							boolean v2 = c2.getBooleanCellValue();
							res = v1==v2;
							System.out.println("Its matched : "+ v1 + " === "+ v2);
						}
					} else
					{
						// If cell types are not same, exit comparison  
						Assert.fail("Non matching cell type.");
						res= false;
					}
					}
				}
			}
		}
		System.out.println("Hurray! Both work books have same data.");
		return res;
	}
   
   
	public boolean verifyDataInCvs(String file1,String file2) throws IOException
	{
	    ArrayList al1=new ArrayList();
	    ArrayList al2=new ArrayList();
	    //ArrayList al3=new ArrayList();

	    BufferedReader CSVFile1 = new BufferedReader(new FileReader(file1));
	    String dataRow1 = CSVFile1.readLine();
	    while (dataRow1 != null)
	    {
	        String[] dataArray1 = dataRow1.split(",");
	        for (String item1:dataArray1)
	        { 
	           al1.add(item1);
	        }

	        dataRow1 = CSVFile1.readLine(); // Read next line of data.
	    }

	     CSVFile1.close();

	    BufferedReader CSVFile2 = new BufferedReader(new FileReader(file2));
	    String dataRow2 = CSVFile2.readLine();
	    while (dataRow2 != null)
	    {
	        String[] dataArray2 = dataRow2.split(",");
	        for (String item2:dataArray2)
	        { 
	           al2.add(item2);

	        }
	        dataRow2 = CSVFile2.readLine(); // Read next line of data.
	    }
	     CSVFile2.close();

	     System.out.println("al1 : "+al1);
	 	System.out.println("al2 : "+al2);
	    
		     
	if(al1.equals(al2))
	{
		return true;
	}
	else
	{
		return false;
	}
}
	
	
	
   
}



