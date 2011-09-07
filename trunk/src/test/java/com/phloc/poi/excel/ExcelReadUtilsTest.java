/**
 * Copyright (C) 2006-2011 phloc systems
 * http://www.phloc.com
 * office[at]phloc[dot]com
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.phloc.poi.excel;

import static org.junit.Assert.assertEquals;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

/**
 * Test class for class {@link ExcelReadUtils}.
 * 
 * @author philip
 */
public final class ExcelReadUtilsTest
{
  @Test
  public void testGetCellValueObject ()
  {
    for (final EExcelVersion eVersion : EExcelVersion.values ())
    {
      final Workbook aWB = eVersion.createWorkbook ();
      final Sheet aSheet = aWB.createSheet ();
      final Row aRow = aSheet.createRow (0);
      final Cell aCell = aRow.createCell (0);

      // boolean
      aCell.setCellValue (true);
      assertEquals (Boolean.TRUE, ExcelReadUtils.getCellValueObject (aCell));

      // int
      aCell.setCellValue (4711);
      assertEquals (Integer.valueOf (4711), ExcelReadUtils.getCellValueObject (aCell));

      // long
      aCell.setCellValue (Long.MAX_VALUE);
      assertEquals (Long.valueOf (Long.MAX_VALUE), ExcelReadUtils.getCellValueObject (aCell));

      // double
      aCell.setCellValue (3.14159);
      assertEquals (Double.valueOf (3.14159), ExcelReadUtils.getCellValueObject (aCell));

      // String
      aCell.setCellValue ("Anyhow");
      assertEquals ("Anyhow", ExcelReadUtils.getCellValueObject (aCell));

      // Rich text string
      final Font aFont = aWB.createFont ();
      aFont.setItalic (true);
      final RichTextString aRTS = eVersion.createRichText ("Anyhow");
      aRTS.applyFont (1, 3, aFont);
      aCell.setCellValue (aRTS);
      assertEquals ("Anyhow", ExcelReadUtils.getCellValueObject (aCell));
    }
  }
}
