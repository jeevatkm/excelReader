/**
 * The MIT License
 *
 * Copyright (c) Jeevanandam M. (jeeva@myjeeva.com) 
 * Copyright 2013 https://github.com/DouglasCAyers
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
 * associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute,
 * sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
 * NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
 * DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 * 
 */

package com.myjeeva.poi;

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
 * Inspired by Jeevanandam M. <a
 * href="https://github.com/jeevatkm/generic-repo/tree/master/excelReader"
 * >https://github.com/jeevatkm/generic-repo/tree/master/excelReader</a>
 * </p>
 * 
 * <p>
 * <strong>Usage:</strong> Provide a {@link ExcelRowContentCallback} callback that will be provided
 * a map representing a row of data from the file. The keys will be the column headers and values
 * the row data. Your callback class encapsulates any business logic for processing the string data
 * into dates, numbers, etc to allow full customization of the parsing and processing logic.
 * </p>
 * 
 * @author https://github.com/DouglasCAyers
 */
public class ExcelWorkSheetRowCallbackHandlerTest {
  private static final Log LOG = LogFactory.getLog(ExcelWorkSheetRowCallbackHandlerTest.class);

  public static void main(String[] args) throws Exception {

    String SAMPLE_PERSON_DATA_FILE_PATH = "src/test/resources/Sample-Person-Data.xlsx";

    File file = new File(SAMPLE_PERSON_DATA_FILE_PATH);
    InputStream inputStream = new FileInputStream(file);

    // The package open is instantaneous, as it should be.
    OPCPackage pkg = null;
    try {
      ExcelWorkSheetRowCallbackHandler sheetRowCallbackHandler =
          new ExcelWorkSheetRowCallbackHandler(new ExcelRowContentCallback() {

            @Override
            public void processRow(int rowNum, Map<String, String> map) {

              // Do any custom row processing here, such as save
              // to database
              // Convert map values, as necessary, to dates or
              // parse as currency, etc
              System.out.println("rowNum=" + rowNum + ", map=" + map);

            }

          });

      pkg = OPCPackage.open(inputStream);

      ExcelSheetCallback sheetCallback = new ExcelSheetCallback() {
        private int sheetNumber = 0;

        @Override
        public void startSheet(int sheetNum, String sheetName) {
          this.sheetNumber = sheetNum;
          System.out.println("Started processing sheet number=" + sheetNumber
              + " and Sheet Name is '" + sheetName + "'");
        }

        @Override
        public void endSheet() {
          System.out.println("Processing completed for sheet number=" + sheetNumber);
        }
      };

      System.out.println("Constructor: pkg, sheetRowCallbackHandler, sheetCallback");
      ExcelReader example1 = new ExcelReader(pkg, sheetRowCallbackHandler, sheetCallback);
      example1.process();

      System.out.println("\nConstructor: filePath, sheetRowCallbackHandler, sheetCallback");
      ExcelReader example2 =
          new ExcelReader(SAMPLE_PERSON_DATA_FILE_PATH, sheetRowCallbackHandler, sheetCallback);
      example2.process();

      System.out.println("\nConstructor: file, sheetRowCallbackHandler, sheetCallback");
      ExcelReader example3 = new ExcelReader(file, sheetRowCallbackHandler, null);
      example3.process();

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
