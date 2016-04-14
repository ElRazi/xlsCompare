package com.xlscompare.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class XlsReader {

	private String filePath;

	public XlsReader(String filePath)
	{
		this.filePath = filePath;
	}
	
	public Workbook read()
	{
		Workbook wb = null;
		try(InputStream inp = new FileInputStream(filePath);) {
		
		wb = WorkbookFactory.create(inp);
	    
		} catch (IOException | InvalidFormatException e) {
			e.printStackTrace();
		}

		return wb;
	}
}
