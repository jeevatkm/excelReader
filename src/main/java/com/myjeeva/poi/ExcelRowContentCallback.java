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

import java.util.Map;

/**
 * Callback for processing a single row from excel file. Map keys are same as first row header
 * columns.
 * 
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam M.</a>
 * 
 * @since v1.1
 */
public interface ExcelRowContentCallback {

  void processRow(int rowNum, Map<String, String> map) throws Exception;

}
