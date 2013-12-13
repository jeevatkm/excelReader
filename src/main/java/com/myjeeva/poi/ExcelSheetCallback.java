/**
 * 
 */
package com.myjeeva.poi;

/**
 * Callback for notifying sheet processing
 * 
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam M.</a>
 * 
 * @since v1.2
 */
public interface ExcelSheetCallback {

	void startSheet(int sheetNum);

	void endSheet();
	
}
