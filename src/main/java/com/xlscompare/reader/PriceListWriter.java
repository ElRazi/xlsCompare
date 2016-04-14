package com.xlscompare.reader;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class PriceListWriter {
	private String filePath;
	private Workbook wb;
	private CellStyle cs;

	public static void createXLSFile(File file)
	{
		try (FileOutputStream fileOut = new FileOutputStream(file);){
			Workbook wb = new HSSFWorkbook();
			wb.createSheet("output");
			wb.write(fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public PriceListWriter(String filePath)
	{
		this.filePath = filePath;
	}
	
	public Workbook getWorkbook()
	{
		if( wb == null )
		{
			wb = readWorkbookFromFile();
		}

		return wb;
	}
	
	protected synchronized Workbook readWorkbookFromFile()
	{
		File file = new File(filePath);
		if(!file.exists())
		{
			createXLSFile(file);
		}

		XlsReader xlsReader = new XlsReader(filePath); 
		return xlsReader.read();
	}
	
	public void writeRawToTheEnd( Map<Integer, Object> cellValues, int sheetNum)
	{
		Workbook wb = getWorkbook();
		
		Sheet sheet = wb.getSheetAt(sheetNum);

		int rowNum = sheet.getPhysicalNumberOfRows();
	    Row row = sheet.createRow(rowNum);

	    for(Integer cellNum : cellValues.keySet())
	    {
	    	Object value = cellValues.get(cellNum);
	    	writeCell(row, cellNum, value);
	    }
	}
	
	public void writeCell(Row row, int cellNum, Object cellValue)
	{
		Cell cell = row.createCell(cellNum);
		
		if(cellValue instanceof String)
		{
			cell.setCellValue((String)cellValue);
		}
		if(cellValue instanceof Double)
		{
			cell.setCellValue((Double)cellValue);
		}
		if(cellValue instanceof Date)
		{
			cell.setCellValue((Date)cellValue);
		}

	    //to enable newlines you need set a cell styles with wrap=true
	    cell.setCellStyle(getCellStyle());
	}
	
	private CellStyle getCellStyle()
	{
		if(cs == null)
		{
			cs = getWorkbook().createCellStyle();
			cs.setWrapText(true);
		}
		
		return cs;
	}
	
	public void save()
	{
		XlsWriter writer = new XlsWriter(wb, filePath);
		writer.write();
	}
}
