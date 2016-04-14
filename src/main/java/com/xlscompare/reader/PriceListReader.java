package com.xlscompare.reader;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class PriceListReader {
	private String filePath;
	private Workbook wb;
	
	public PriceListReader(String filePath)
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
		XlsReader xlsReader = new XlsReader(filePath); 
		return xlsReader.read();
	}
	
	public void processSupplierPriceRaws( RowProcessor rowProcessor, int sheetNum )
	{
		wb = getWorkbook();
		if( wb == null )
		{
			System.out.println("Error opening workbook");
			return;
		}
		
		System.out.println("Processing: " + filePath);
		processSheet(wb.getSheetAt(sheetNum), rowProcessor);
	}
	
	private void processSheet( Sheet sheet, RowProcessor rowProcessor )
	{
		if( sheet == null || rowProcessor == null )
		{
			System.out.println("No sheet or row processor");
			return;
		}
		
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		for(int i = firstRowNum; i<lastRowNum; i++)
		{
			rowProcessor.process(sheet.getRow(i));
		}
	}
}