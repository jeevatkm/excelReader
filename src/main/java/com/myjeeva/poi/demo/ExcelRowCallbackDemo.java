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

package com.myjeeva.poi.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;

import com.myjeeva.poi.ExcelReader;
import com.myjeeva.poi.ExcelRowContentCallback;
import com.myjeeva.poi.ExcelSheetCallback;
import com.myjeeva.poi.ExcelWorkSheetRowCallbackHandler;

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
 */
public class ExcelRowCallbackDemo {
	private static final Log LOG = LogFactory
			.getLog(ExcelRowCallbackDemo.class);

	public static void main(String[] args) throws Exception {

		String SAMPLE_PERSON_DATA_FILE_PATH = "src/main/resources/Sample-Person-Data.xlsx";

		File file = new File(SAMPLE_PERSON_DATA_FILE_PATH);
		InputStream inputStream = new FileInputStream(file);

		// The package open is instantaneous, as it should be.
		OPCPackage pkg = null;
		try {
			ExcelWorkSheetRowCallbackHandler sheetRowCallbackHandler = new ExcelWorkSheetRowCallbackHandler(
					new ExcelRowContentCallback() {

						@Override
						public void processRow(int rowNum,
								Map<String, String> map) {

							// Do any custom row processing here, such as save
							// to database
							// Convert map values, as necessary, to dates or
							// parse as currency, etc
							System.out.println("rowNum=" + rowNum + ", map="
									+ map);

						}

					});

			pkg = OPCPackage.open(inputStream);
			ExcelReader excelReader = new ExcelReader(pkg,
					sheetRowCallbackHandler, new ExcelSheetCallback() {
						private int sheetNumber = 0;
						@Override
						public void startSheet(int sheetNum) {
							this.sheetNumber = sheetNum;

							System.out
									.println("Started processing sheet number="
											+ sheetNumber);
						}

						@Override
						public void endSheet() {
							System.out
									.println("Processing completed for sheet number="
											+ sheetNumber);
						}
					});
			excelReader.process();

		} catch (RuntimeException are) {
			LOG.error(are.getMessage(), are.getCause());
		} catch (InvalidFormatException ife) {
			LOG.error(ife.getMessage(), ife.getCause());
		} catch (IOException ioe) {
			LOG.error(ioe.getMessage(), ioe.getCause());
		} finally {
			IOUtils.closeQuietly(inputStream);
			try {
				if (null != pkg) {
					pkg.close();
				}
			} catch (IOException e) {
				// just ignore IO exception
			}
		}
	}
}
