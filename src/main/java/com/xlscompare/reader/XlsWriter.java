package com.xlscompare.reader;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class XlsWriter {
	protected String filePath;
	protected Workbook wb;

	public XlsWriter(Workbook wb,  String filePath)
	{
		this.wb = wb;
		this.filePath = filePath;
	}
	
	public void write()
	{
		if( wb != null )
		{
			try(FileOutputStream fileOut = new FileOutputStream(filePath);) {
				wb.write(fileOut);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
