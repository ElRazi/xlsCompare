package com.xlscompare;

import java.util.Arrays;
import java.util.List;
import java.util.Properties;

import com.xlscompare.reader.PriceListComparator;

/**
 * Hello world!
 *
 */
public class XlsComparator 
{
    public static void main( String[] args )
    {
//        System.out.println( "Hello World!" );
//
//        XlsReaderSandBox reader = new XlsReaderSandBox("D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\Драбини-WERK.xls");
//        reader.read();
        
//        XlsReader reader = new XlsReader("D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\прайс-Way-09.03.16.xlsx");
//        reader.read();
        
//
//        XlsReader reader = new XlsReader("D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\export-products-16-03-16_16-57-12это-база-с-сайта.xlsx");
//        reader.write();
        
//        String sourceFile = "D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\test00\\source.xlsx";
//    	String destFile = "D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\test00\\dest.xlsx";
//    	int sourceColumn=1;
//    	int destColumn=0;
//    	int sourceSheet=0;
//    	int destSheet=0;
//    	int[] sourceCells = new int[]{1,2,3};
//    	int[] destCells = new int[]{1,3,2};

    	String propertiesFile = "input.properties";
    	if(args.length > 0)
    		propertiesFile = args[1];
    	
    	System.out.println(propertiesFile);
//    	Properties readProperties = PropertyFileReader.readProperties("D:\\Projects\\Java\\kirushus\\xlscompare\\test-data\\test03\\input.properties");
    	Properties readProperties = PropertyFileReader.readProperties(propertiesFile);
    	
    	String sourceFile = readProperties.getProperty(PropertyFileReader.SOURCE_FILE);
    	String destFile = readProperties.getProperty(PropertyFileReader.DEST_FILE);
    	String outputFile = readProperties.getProperty(PropertyFileReader.OUTPUTFILE);
    	int sourceColumn = Integer.parseInt(readProperties.getProperty(PropertyFileReader.SOURCE_COLUMN));
    	int destColumn = Integer.parseInt(readProperties.getProperty(PropertyFileReader.DEST_COLUMN));
    	int sourceSheet = Integer.parseInt(readProperties.getProperty(PropertyFileReader.SOURCE_SHEET));
    	int destSheet = Integer.parseInt(readProperties.getProperty(PropertyFileReader.DEST_SHEET));
    	int[] sourceCells = csvToIntArray(readProperties.getProperty(PropertyFileReader.SOURCE_CELLS));
    	int[] destCells = csvToIntArray(readProperties.getProperty(PropertyFileReader.DEST_CELLS));
    	
    	System.out.println(sourceFile);
    	System.out.println(destFile);
    	System.out.println(outputFile);
    	System.out.println(sourceColumn);
    	System.out.println(destColumn);
    	System.out.println(sourceSheet);
    	System.out.println(destSheet);
    	System.out.println(Arrays.toString(sourceCells));
    	System.out.println(Arrays.toString(destCells));

        PriceListComparator plc = new PriceListComparator(sourceFile, destFile, sourceSheet, destSheet, sourceColumn, destColumn, sourceCells, destCells);
//        plc.compareAndLog(outputFile);
        plc.findAndApplyDifferences(outputFile);
        
        System.out.println("Done");
    }
    
    private static int[] csvToIntArray(String str)
    {
    	String [] items = str.split(",");
    	List<String> container = Arrays.asList(items);
    	
    	int[] array = new int[container.size()];
    	
    	int i = 0;
    	for(String s: container)
    	{
    		array[i++] = Integer.parseInt(s);
    	}
    	
    	return array;
    }
}
