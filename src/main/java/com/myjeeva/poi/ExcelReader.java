/**
 * The MIT License
 *
 * Copyright (c) 2012 www.myjeeva.com
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE. 
 * 
 */
package com.myjeeva.poi;

/**
 * Generic Excel File(XLSX) Reading using Apache POI 
 * 
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam Madanagopal</a> 
 */
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class ExcelReader {

	private static final Log LOG = LogFactory.getLog(ExcelReader.class);

	private static final int READ_ALL = -1;

	private OPCPackage xlsxPackage;
	private SheetContentsHandler sheetContentsHandler;

	/**
	 * Constructor: Microsoft Excel File Reader (XLSX)
	 * 
	 * @param pkg
	 *            a {@link OPCPackage} object - The package to process XLSX
	 * @param sheetContentsHandler
	 *            a {@link SheetContentsHandler} object - WorkSheet contents
	 *            handler
	 */
	public ExcelReader(OPCPackage pkg, SheetContentsHandler sheetContentsHandler) {
		this.xlsxPackage = pkg;
		this.sheetContentsHandler = sheetContentsHandler;
	}

	/**
	 * Processing all the WorkSheet from XLSX Workbook.
	 * 
	 *  <br><br><strong>For Example:</strong><br>
	 *  <code>ExcelReader excelReader = new ExcelReader(pkg, workSheetHandler);
	 *  excelReader.process();</code>
	 * @throws RuntimeException
	 */
	public void process() throws RuntimeException {
		read(READ_ALL);
	}

	/**
	 * Processing of particular WorkSheet (zero based) from XLSX Workbook.
	 * 
	 * <br><br><strong>For Example:</strong><br>
	 * <code>ExcelReader excelReader = new ExcelReader(pkg, workSheetHandler);
	 * excelReader.process(2);</code>
	 * 
	 * @param sheetNumber
	 *            a int object
	 * @throws RuntimeException
	 */
	public void process(int sheetNumber) throws RuntimeException {
		read(sheetNumber);
	}

	private void read(int sheetNumber) throws RuntimeException {
		ReadOnlySharedStringsTable strings;
		try {
			strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
			XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
			StylesTable styles = xssfReader.getStylesTable();
			Iterator<InputStream> sheets = xssfReader.getSheetsData();
			for (int sheet = 0; sheets.hasNext(); sheet++) {
				InputStream stream = sheets.next();
				if ((READ_ALL == sheetNumber) || (sheet == sheetNumber)) {
					readSheet(styles, strings, stream);
				}
				IOUtils.closeQuietly(stream);
			}
		} catch (IOException ioe) {
			LOG.error(ioe.getMessage(), ioe.getCause());
		} catch (SAXException se) {
			LOG.error(se.getMessage(), se.getCause());
		} catch (OpenXML4JException oxe) {
			LOG.error(oxe.getMessage(), oxe.getCause());
		} catch (ParserConfigurationException pce) {
			LOG.error(pce.getMessage(), pce.getCause());
		}
	}

	/**
	 * Parses the content of one sheet using the specified styles and
	 * shared-strings tables.
	 * 
	 * @param styles
	 *            a {@link StylesTable} object
	 * @param sharedStringsTable
	 *            a {@link ReadOnlySharedStringsTable} object
	 * @param sheetInputStream
	 *            a {@link InputStream} object
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 */
	private void readSheet(StylesTable styles,
			ReadOnlySharedStringsTable sharedStringsTable,
			InputStream sheetInputStream) throws IOException,
			ParserConfigurationException, SAXException {
		InputSource sheetSource = new InputSource(sheetInputStream);		
		SAXParserFactory saxFactory = SAXParserFactory.newInstance();
		
		SAXParser saxParser = saxFactory.newSAXParser();
		XMLReader sheetParser = saxParser.getXMLReader();
		
		ContentHandler handler = new XSSFSheetXMLHandler(styles,
				sharedStringsTable, sheetContentsHandler, true);
		sheetParser.setContentHandler(handler);
		
		sheetParser.parse(sheetSource);
	}
}
