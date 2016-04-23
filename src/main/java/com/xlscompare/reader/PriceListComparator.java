package com.xlscompare.reader;

import java.io.File;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

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
	int[] addSourceColumns;
	int[] addDestColumns;

	public PriceListComparator(String sourceFile, String destFile, int sourceSheet, int destSheet, int sourceNameColumn, int destNameColumn, int[] sourceColumns, int[] destColumns, int[] addSourceColumns, int[] addDestColumns)
	{
		this.sourceFile = sourceFile;
		this.destFile = destFile;
		this.sourceNameColumn = sourceNameColumn;
		this.destNameColumn = destNameColumn;
		this.sourceSheet = sourceSheet;
		this.destSheet = destSheet;
		this.sourceColumns = sourceColumns;
		this.destColumns = destColumns;
		this.addSourceColumns = addSourceColumns;
		this.addDestColumns = addDestColumns;
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
		
		if(addSourceColumns.length != addDestColumns.length)
		{
			somethingIsWrong = true;
			System.out.println(" 'add' column lengths should be equal");
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

		//Differences in values
		DifferentialSupplierRowProcessor differentialSupplierRowProcessor = new DifferentialSupplierRowProcessor(destFile, sourceNameColumn, destNameColumn, destSheet, sourceColumns, destColumns, addSourceColumns, addDestColumns);
		supplierFileReader.processRaws(differentialSupplierRowProcessor, sourceSheet);
		
		differentialSupplierRowProcessor.indexSource(supplierFileReader, sourceSheet);
		differentialSupplierRowProcessor.doRemove();
		differentialSupplierRowProcessor.doAdd();
		
		differentialSupplierRowProcessor.saveProcessedWorkbook(outputFile);
	}

	public static class DifferentialSupplierRowProcessor implements RowProcessor, SheetProcessor
	{
		PriceListReader baseFileReader;
		int sourceKeyColumn;
		int destKeyColumn;
		int destSheet;
		int[] sourceColumns;
		int[] destColumns;
		int[] addSourceColumns;
		int[] addDestColumns;
		Map<Object, Row> rowIndexBase;
		Map<Object, Row> rowIndexSource;
		
		public DifferentialSupplierRowProcessor(String destFile, int sourceKeyColumn, int destKeyColumn, int destSheet, int[] sourceColumns, int[] destColumns, int[] addSourceColumns, int[] addDestColumns)
		{
			this.sourceKeyColumn = sourceKeyColumn;
			this.destKeyColumn = destKeyColumn;
			this.destSheet = destSheet;
			this.sourceColumns = sourceColumns;
			this.destColumns = destColumns;
			this.addSourceColumns = addSourceColumns;
			this.addDestColumns = addDestColumns;

			baseFileReader = new PriceListReader(destFile);
			rowIndexBase = indexFile(baseFileReader, destSheet, destKeyColumn);
		}

		public Map<Object, Row> indexFile(PriceListReader reader, int sheet, int keyColumn)
		{
			final Map<Object, Row> rowIndex = new HashMap<>(); 
			reader.processRaws(new RowProcessor() {
				@Override
				public void process(Row row) {
					if(row == null)
					{
						return;
					}

					Object destKeyValue = Util.readCellValue(row, keyColumn);
					if(destKeyValue != null)
					{
						rowIndex.put(destKeyValue, row);
					}
						
				}
			}, sheet);
			
			return rowIndex;
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

				Row destRow = rowIndexBase.get(cellValue);
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

		@Override
		public void process(Sheet sheet) {
			// TODO Auto-generated method stub
			
		}
		
		public void indexSource(PriceListReader supplierFileReader, int sourceSheet)
		{
			rowIndexSource = indexFile(supplierFileReader, sourceSheet, sourceKeyColumn);
		}
		
		public void doAdd()
		{
			final Set<Row> rowsToAdd = new HashSet<>();
			for (Object sourceKey : rowIndexSource.keySet()) {
				if(rowIndexBase.get(sourceKey) == null) {
					rowsToAdd.add(rowIndexSource.get(sourceKey));
				}
			}

			if(!rowsToAdd.isEmpty()) {
				baseFileReader.processSheets(new SheetProcessor() {
					@Override
					public void process(Sheet sheet) {
						for( Row row : rowsToAdd )
						{
							System.out.println("added row: " + row.getRowNum() );
							Row newRow = sheet.createRow(sheet.getLastRowNum()+1);
							for(int i = 0; i < addSourceColumns.length; i++) {
								Object cellValue = Util.readCellValue(row, addSourceColumns[i]);
								newRow.createCell(addDestColumns[i]);
								Util.writeCellValue(newRow, addDestColumns[i], cellValue);
							}
						}
					}
				}, destSheet);
			}
		}
		
		public void doRemove()
		{
			final Set<Integer> rowsToRemove = new HashSet<>();
			for (Object baseKey : rowIndexBase.keySet()) {
				Row sourceRow = rowIndexSource.get(baseKey);
				if(sourceRow == null) {
					rowsToRemove.add(rowIndexBase.get(baseKey).getRowNum());
				}
			}
			
			if(!rowsToRemove.isEmpty())
			{
				baseFileReader.processSheets(new SheetProcessor() {
					@Override
					public void process(Sheet sheet) {
						for (int r = sheet.getLastRowNum(); r >= 0; r--) {
							if (rowsToRemove.contains(r)) {
								Row row = sheet.getRow(r);
								System.out.println("remove row: " + r);
								if (r == sheet.getLastRowNum()) {
									sheet.removeRow(row);
								} else {
									sheet.removeRow(row);
									sheet.shiftRows(r+1, sheet.getLastRowNum(), -1);
								}
							}
						}
					}
				}, destSheet);
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
