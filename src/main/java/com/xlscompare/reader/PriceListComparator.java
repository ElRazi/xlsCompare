package com.xlscompare.reader;

import java.io.File;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class PriceListComparator 
{
	String sourceFile;
	String destFile;
	int sourceNameColumn;
	int destNameColumn;
	int sourceSheet;
	int destSheet;
	int[] sourceColumns;
	int[] destColumns;
	
	public PriceListComparator(String sourceFile, String destFile, int sourceSheet, int destSheet, int sourceNameColumn, int destNameColumn, int[] sourceColumns, int[] destColumns)
	{
		this.sourceFile = sourceFile;
		this.destFile = destFile;
		this.sourceNameColumn = sourceNameColumn;
		this.destNameColumn = destNameColumn;
		this.sourceSheet = sourceSheet;
		this.destSheet = destSheet;
		this.sourceColumns = sourceColumns;
		this.destColumns = destColumns;
		checkValues();
	}

	private void checkValues()
	{
		boolean somethingIsWrong = false;
		File file = new File(sourceFile);
		if(!file.exists())
		{
			somethingIsWrong = true;
			System.out.println(file.getAbsolutePath() + " does not exist");
		}
		
		file = new File(destFile);
		if(!file.exists())
		{
			somethingIsWrong = true;
			System.out.println(file.getAbsolutePath() + " does not exist");
		}
		
		if(sourceColumns.length != destColumns.length)
		{
			somethingIsWrong = true;
			System.out.println("column lengths should be equal");
		}
		
		
		if(somethingIsWrong)
			throw new IllegalArgumentException();
	}
	
	public void compareAndLog(String outputFile)
	{
		Set<Object> destNames = readNames(destFile, destNameColumn);
		List<Row> uniqueRows = getUniqueRows(destNames, sourceFile, sourceNameColumn);
		writeRowsToFile(uniqueRows, outputFile);
	}

	private void writeRowsToFile(List<Row> rows, String outputFile)
	{
		PriceListWriter writer = new PriceListWriter(outputFile);
		
		for(Row row : rows)
		{
			Map<Integer, Object> cellValues = new HashMap<>();
			for(int i = 0; i < sourceColumns.length; i++)
			{
				Object cellValue = readCellValue(row, sourceColumns[i]);

				cellValues.put(destColumns[i], cellValue);
			}
			writer.writeRawToTheEnd(cellValues, 0);
		}
		
		writer.save();
	}
	
	private List<Row> getUniqueRows(Set<Object> destNames, String sourceFile, int sourceColumn)
	{
		List<Row> uniqueRows = new LinkedList<>();
		PriceListReader destReader = new PriceListReader(sourceFile);
		destReader.processSupplierPriceRaws(new RowProcessor() {
			@Override
			public void process(Row row) {
				Object value = readCellValue(row, sourceColumn);
				if(value != null && !destNames.contains(value))
				{
					uniqueRows.add(row);
//					System.out.println(value);
				}
			}
		}, destSheet);
		
		return uniqueRows;
	}

	private Set<Object> readNames(String file, int nameColumn)
	{
		final Set<Object> names = new HashSet<>();
		PriceListReader sourceReader = new PriceListReader(file);
		sourceReader.processSupplierPriceRaws(new RowProcessor() {
					@Override
					public void process(Row row) {
						Object stringCellValue = readCellValue(row, nameColumn);
						if (stringCellValue != null && stringCellValue != null) {
							names.add(stringCellValue);
						}
					}
				}, sourceSheet);

		return names;
	}

	private Object readCellValue(Row row, int columnNum)
	{
		Object cellValue = null;
		Cell cell = row.getCell(columnNum);

		if(cell == null)
		{
			System.out.println("Error reading cell: " + columnNum + " row: " + row.getRowNum());
			return null;
		}
		
		if( cell.getCellType() == Cell.CELL_TYPE_STRING)
		{
			cellValue = cell != null ? cell.getStringCellValue() : null;
		}
		else if( cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
		{
			cellValue = cell.getNumericCellValue();
		}
		else if( cell.getCellType() == Cell.CELL_TYPE_FORMULA)
		{
			cellValue = cell.getCellFormula();
		}
		else if( cell.getCellType() == Cell.CELL_TYPE_BOOLEAN)
		{
			cellValue = cell.getBooleanCellValue();
		}
		
		return cellValue;
	}
}
