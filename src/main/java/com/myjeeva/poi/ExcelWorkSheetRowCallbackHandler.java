/*
   Copyright 2013 https://github.com/DouglasCAyers

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
 */

package com.myjeeva.poi;

import java.util.LinkedHashMap;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;

/**
 * <p>
 * Excel Worksheet Handler for XML SAX parsing (.xlsx document model) <a
 * href="http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api"
 * >http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api</a>
 * </p>
 * 
 * <p>
 * Inspired by Jeevanandam Madanagopal <a
 * href="https://github.com/jeevatkm/generic-repo/tree/master/excelReader"
 * >https://github.com/jeevatkm/generic-repo/tree/master/excelReader</a>
 * </p>
 * 
 * <p>
 * <strong>Usage:</strong> Provide a {@link ExcelRowContentCallback} callback
 * that will be provided a map representing a row of data from the file. The
 * keys will be the column headers and values the row data. Your callback class
 * encapsulates any business logic for processing the string data into dates,
 * numbers, etc to allow full customization of the parsing and processing logic.
 * </p>
 * 
 * @author https://github.com/DouglasCAyers
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam M.</a>
 */
public class ExcelWorkSheetRowCallbackHandler implements SheetContentsHandler {

	private static final Log LOG = LogFactory
			.getLog(ExcelWorkSheetRowCallbackHandler.class);

	private static final int HEADER_ROW = 0;

	// once an entire row of data has been read, pass map to this callback for
	// processing
	private ExcelRowContentCallback rowCallback;

	// LinkedHashMaps are used so iteration order is predictable over insertion
	// order
	private LinkedHashMap<String, String> currentRowMap; // map of column
															// headers => row
															// values (eg, 'A'
															// => 'White Shirts'
															// )
	private LinkedHashMap<String, String> columnHeaders; // map of column
															// references =>
															// column headers
															// (eg, 'A' =>
															// 'Product Title' )
	private int currentRow;

	public ExcelWorkSheetRowCallbackHandler(
			ExcelRowContentCallback rowCallbackHandler) {
		this.rowCallback = rowCallbackHandler;
	}

	/**
	 * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler#startRow(int)
	 */
	@Override
	public void startRow(int rowNum) {

		this.currentRow = rowNum;

		if (this.currentRow == HEADER_ROW) {
			this.columnHeaders = new LinkedHashMap<String, String>();
		} else {
			this.currentRowMap = new LinkedHashMap<String, String>();

			// Add column header as key into current row map so that each entry
			// will exist.
			// This ensures each column header will be in the "currentRowMap"
			// when passed to the user callback.
			// Remember, the 'column headers map key' is the actual cell column
			// reference, it's value is the file column header value.
			// In the 'cell' method below, this empty string will be overwritten
			// with the file row value (if has one, else remains empty).
			for (String columnHeader : this.columnHeaders.values()) {
				this.currentRowMap.put(columnHeader, "");
			}

		}

	}

	/**
	 * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler#cell(java.lang.String,
	 *      java.lang.String)
	 */
	@Override
	public void cell(String cellReference, String formattedValue) {

		// Note, POI will not invoke this method if the cell
		// is blank or if it detects there's no more data in the row.
		// So don't count on this being invoked the same number of times each
		// row.
		// That's another reason why in above code we ensure each column header
		// is in the 'currentRowMap'.

		if (this.currentRow == HEADER_ROW) {
			this.columnHeaders.put(getColumnReference(cellReference),
					formattedValue);
		} else {
			String columnHeader = this.columnHeaders
					.get(getColumnReference(cellReference));
			this.currentRowMap.put(columnHeader, formattedValue);
		}

	}

	/**
	 * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler#endRow()
	 */
	@Override
	public void endRow() {

		if (this.currentRow > HEADER_ROW) {
			try {
				LOG.debug("rowNum=" + currentRow + ", map=" + currentRowMap);

				this.rowCallback.processRow(currentRow, currentRowMap);
			} catch (Exception e) {
				throw new RuntimeException("Error invoking callback", e);
			}
		}

	}

	/**
	 * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler#headerFooter(java.lang.String,
	 *      boolean, java.lang.String)
	 */
	@Override
	public void headerFooter(String text, boolean isHeader, String tagName) {
		// Not Used
	}

	/**
	 * Returns the alphabetic column reference from this cell reference.
	 * Example: Given 'A12' returns 'A' or given 'BA205' returns 'BA'
	 */
	private String getColumnReference(String cellReference) {

		if (StringUtils.isBlank(cellReference)) {
			return "";
		}

		return cellReference.split("[0-9]*$")[0];
	}
}
