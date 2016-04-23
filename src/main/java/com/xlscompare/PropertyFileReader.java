package com.xlscompare;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class PropertyFileReader {
	
	public static final String SOURCE_FILE = "supplierFile";
	public static final String DEST_FILE = "baseFile";
	public static final String SOURCE_COLUMN = "supplierNameColumn";
	public static final String DEST_COLUMN = "baseNameColumn";
	public static final String SOURCE_SHEET = "supplierSheet";
	public static final String DEST_SHEET = "baseSheet";
	public static final String SOURCE_CELLS = "supplierCells";
	public static final String DEST_CELLS = "baseCells";
	public static final String ADD_SOURCE_CELLS = "addSupplierCells";
	public static final String ADD_DEST_CELLS = "addBaseCells";
	public static final String OUTPUTFILE = "outputFile";
	
	public static Properties readProperties(String propertyFile)
	{
		Properties prop = new Properties();
		
		try (FileInputStream input = new FileInputStream(propertyFile);)
		{
			prop.load(input);
		} catch (IOException e) {
			e.printStackTrace();
		}

		return prop;
	}
}
