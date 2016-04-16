package com.xlscompare.reader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Util {
	public static Object readCellValue(Row row, int columnNum)
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
	
	public static void writeCellValue(Row row, int columnNum, Object cellValue)
	{
		Cell cell = row.getCell(columnNum);

		if(cell == null)
		{
			System.out.println("Error writing cell: " + columnNum + " row: " + row.getRowNum());
			return;
		}
		
		if(cellValue instanceof String)
		{
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue((String)cellValue);
		}
		if(cellValue instanceof Double)
		{
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Double)cellValue);
		}
		if(cellValue instanceof Boolean)
		{
			cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
			cell.setCellValue((Boolean)cellValue);
		}
	}
}
