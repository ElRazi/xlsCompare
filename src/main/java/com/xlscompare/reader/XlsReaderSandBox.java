package com.xlscompare.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*
for .xls  - Workbook wb = new HSSFWorkbook();
for .xlsx - Workbook wb = new XSSFWorkbook();
 */

public class XlsReaderSandBox {
	String filePath;
	
	public XlsReaderSandBox(String filePath)
	{
		this.filePath = filePath;
	}
	
	public void read()
	{
		Workbook wb = null;
		try(InputStream inp = new FileInputStream(filePath);) {
		
		wb = WorkbookFactory.create(inp);
	    
//	    if (cell == null)
//	        cell = row.createCell(3);
//	    cell.setCellType(Cell.CELL_TYPE_STRING);
//	    cell.setCellValue("a test");

//	    // Write the output to a file
//	    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//	    wb.write(fileOut);
//	    fileOut.close();
	    
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
		
		if( wb != null )
		{
			Sheet sheet = wb.getSheetAt(0);
		    Row row = sheet.getRow(11);
		    Cell cell = row.getCell(0);
			System.out.println(cell.getStringCellValue());
		}
	}
	
	public void write()
	{
		Workbook wb = null;

		File file = new File(filePath);
		if(!file.exists())
		{
			return;
		}
		
		try(InputStream inp = new FileInputStream(filePath);) {
		
        wb = WorkbookFactory.create(inp);
	    Sheet sheet = wb.getSheetAt(0);
	    Row row = sheet.createRow(10);
	    Cell cell = row.createCell(2);
	    cell.setCellValue("Use \n test cell row Creation");
	    
	    //to enable newlines you need set a cell styles with wrap=true
	    CellStyle cs = wb.createCellStyle();
	    cs.setWrapText(true);
	    cell.setCellStyle(cs);
	    
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}

		if( wb != null )
		{
			try(FileOutputStream fileOut = new FileOutputStream(filePath);) {
				wb.write(fileOut);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		}
		
	}
	
}
