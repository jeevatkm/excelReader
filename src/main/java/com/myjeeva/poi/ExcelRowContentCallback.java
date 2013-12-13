/**
 * 
 */
package com.myjeeva.poi;

import java.util.Map;

/**
 * Callback for processing a single row from excel file. Map keys are same
 * as first row header columns.
 * 
 * @since v1.1
 */
public interface ExcelRowContentCallback {

	void processRow(int rowNum, Map<String, String> map) throws Exception;
	
}
