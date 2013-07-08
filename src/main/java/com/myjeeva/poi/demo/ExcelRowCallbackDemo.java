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
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.util.IOUtils;

import com.myjeeva.poi.ExcelXSSFRowCallbackHandler;
import com.myjeeva.poi.ExcelXSSFRowCallbackHandler.ExcelRowContentCallback;

/**
 * Excel Worksheet Handler for XML SAX parsing (.xlsx document model)
 * http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
 *
 * Inspired by Jeevanandam Madanagopal
 * https://github.com/jeevatkm/generic-repo/tree/master/excelReader
 *
 * Usage: Provide a {@link ExcelRowContentCallback} callback that will be provided a map
 * representing a row of data from the file. The keys will be the column headers and values the row data.
 * Your callback class encapsulates any business logic for processing the string data into dates, numbers, etc
 * to allow full customization of the parsing and processing logic.
 *
 * @author https://github.com/DouglasCAyers
 */
public class ExcelRowCallbackDemo {

	public static void main( String[] args ) throws Exception {

		String SAMPLE_PERSON_DATA_FILE_PATH = "src/main/resources/Sample-Person-Data.xlsx";

		File file = new File( SAMPLE_PERSON_DATA_FILE_PATH );
		InputStream fileInputStream = new FileInputStream( file );

		try {

			// Can create handler passing in various arguments for the excel source:
			//	1) String (file path)
			//	2) File
			//	3) InputStream
			//	4) OPCPackage
			ExcelXSSFRowCallbackHandler handler = new ExcelXSSFRowCallbackHandler( SAMPLE_PERSON_DATA_FILE_PATH, new ExcelRowContentCallback() {

				@Override
				public void process( Map<String, String> map, int rowNum ) {

					// Do any custom row processing here, such as save to database
					// Convert map values, as necessary, to dates or parse as currency, etc
					System.out.println( "rowNum=" + rowNum + ", map=" + map );

				}

			});

			// This method returns nothing, as all processing will occur in your callback defined above
			handler.parse();

		} finally {

			IOUtils.closeQuietly( fileInputStream );

		}

	}

}
