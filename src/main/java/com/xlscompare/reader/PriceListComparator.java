package com.xlscompare.reader;

import java.io.File;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

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

	public void findAndApplyDifferences(String outputFile)
	{
		PriceListReader supplierFileReader =  new PriceListReader(sourceFile);

		DifferentialSupplierRowProcessor differentialSupplierRowProcessor = new DifferentialSupplierRowProcessor(destFile, sourceNameColumn, destNameColumn, destSheet, sourceColumns, destColumns);
		supplierFileReader.processRaws(differentialSupplierRowProcessor, sourceSheet);
		differentialSupplierRowProcessor.saveProcessedWorkbook(outputFile);
	}

	public static class DifferentialSupplierRowProcessor implements RowProcessor
	{
		PriceListReader baseFileReader;
		int sourceKeyColumn;
		int destKeyColumn;
		int destSheet;
		int[] sourceColumns;
		int[] destColumns;
		Map<Object, Row> rowIndex;
		
		
		public DifferentialSupplierRowProcessor(String destFile, int sourceKeyColumn, int destKeyColumn, int destSheet, int[] sourceColumns, int[] destColumns)
		{
			this.sourceKeyColumn = sourceKeyColumn;
			this.destKeyColumn = destKeyColumn;
			this.destSheet = destSheet;
			this.sourceColumns = sourceColumns;
			this.destColumns = destColumns;
			indexBaseFile(destFile);
		}

		public void indexBaseFile(String destFile)
		{
			baseFileReader = new PriceListReader(destFile);
			rowIndex = new HashMap<>();
			baseFileReader.processRaws(new RowProcessor() {
				@Override
				public void process(Row row) {
					if(row == null)
					{
						return;
					}

					Object destKeyValue = Util.readCellValue(row, destKeyColumn);
					if(destKeyValue != null)
					{
						rowIndex.put(destKeyValue, row);
					}
						
				}
			}, destSheet);
		}
		
		@Override
		public void process(Row row) 
		{
			Object cellValue = Util.readCellValue(row, sourceKeyColumn);

			if( cellValue != null )
			{
//				RowFinder rowFinder = new RowFinder(cellValue, row, destKeyColumn);
//				baseFileReader.processRaws(rowFinder, destSheet);
//				List<Row> foundRows = rowFinder.getFoundRows();

				Row destRow = rowIndex.get(cellValue);
				if( destRow != null )
				{
					compareAndSetDifferences(row, destRow, sourceColumns, destColumns);
				}
			}
		}

		public void compareAndSetDifferences(Row sourceRow, Row destRow, int[] sourceColumns, int[] destColumns)
		{
			for(int i = 0; i<sourceColumns.length; i++)
			{
				Object sourceValue = Util.readCellValue(sourceRow, sourceColumns[i]);
				Object destValue = Util.readCellValue(destRow, destColumns[i]);
				
				if(sourceValue != null && ! sourceValue.equals(destValue))
				{
					Util.writeCellValue(destRow, destColumns[i], sourceValue);
					System.out.println("Change in: [" + destRow.getRowNum() + ", " + destColumns[i] + "] values: " + destValue + " -> " + sourceValue);
				}
			}
		}

		public static class RowFinder implements RowProcessor
		{
			Object keyValue;
			Row sourceRow;
			int destKeyColumn;
			List<Row> foundRows = new LinkedList<>();


			public RowFinder(Object keyValue, Row sourceRow, int destKeyColumn)
			{
				this.keyValue = keyValue;
				this.sourceRow = sourceRow;
				this.destKeyColumn = destKeyColumn;
			}

			@Override
			public void process(Row row) 
			{
				Object cellValue = Util.readCellValue(row, destKeyColumn);

				if(cellValue != null && cellValue.equals(keyValue))
				{
					foundRows.add(row);
				}
			}

			public List<Row> getFoundRows() {
				return foundRows;
			}
		}
		
		public void saveProcessedWorkbook(String outputFile)
		{
			Workbook workbook = baseFileReader.getWorkbook();
			PriceListWriter writer = new PriceListWriter(outputFile);
			writer.setWb(workbook);
			writer.save();
		}
	}
	
	private void writeRowsToFile(List<Row> rows, String outputFile)
	{
		PriceListWriter writer = new PriceListWriter(outputFile);
		
		for(Row row : rows)
		{
			Map<Integer, Object> cellValues = new HashMap<>();
			for(int i = 0; i < sourceColumns.length; i++)
			{
				Object cellValue = Util.readCellValue(row, sourceColumns[i]);

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
		destReader.processRaws(new RowProcessor() {
			@Override
			public void process(Row row) {
				Object value = Util.readCellValue(row, sourceColumn);
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
		sourceReader.processRaws(new RowProcessor() {
					@Override
					public void process(Row row) {
						Object stringCellValue = Util.readCellValue(row, nameColumn);
						if (stringCellValue != null && stringCellValue != null) {
							names.add(stringCellValue);
						}
					}
				}, sourceSheet);

		return names;
	}
}
