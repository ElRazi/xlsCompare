package com.xlscompare.reader;

import org.apache.poi.ss.usermodel.Row;

public interface RowProcessor {
	public void process(Row row);
}
