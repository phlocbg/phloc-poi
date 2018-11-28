/**
 * Copyright (C) 2006-2015 phloc systems
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

import static org.junit.Assert.fail;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Assert;
import org.junit.Test;

import com.phloc.commons.string.StringHelper;

/**
 * Test class for class {@link WorkbookCreationHelper}.
 *
 * @author Boris Gregorcic
 */
public final class WorkbookCreationHelperTest
{
  private static final char CONTENT_BIT = 'x';
  
  @Test
  public void testAddCellFormulaUndefined ()
  {
    final WorkbookCreationHelper aWCH = new WorkbookCreationHelper (EExcelVersion.XLSX);
    aWCH.createNewSheet ();
    aWCH.addRow ();
    try
    {
      aWCH.addCellFormula (")invalid("); //$NON-NLS-1$
      fail ();
    }
    catch (final FormulaParseException ex)
    {
      // expected
    }
  }

  @Test
  public void testAddCellFormulaSyntax ()
  {
    final WorkbookCreationHelper aWCH = new WorkbookCreationHelper (EExcelVersion.XLSX);
    aWCH.createNewSheet ();
    aWCH.addRow ();
    aWCH.addCellFormula ("ABC(A1)"); //$NON-NLS-1$
  }
  
  private static final String getSmallContent ()
  {
    return StringHelper.getRepeated (CONTENT_BIT, 100);
  }

  private static final String getMaxContent ()
  {
    return StringHelper.getRepeated (CONTENT_BIT, SpreadsheetVersion.EXCEL97.getMaxTextLength ());
  }

  @Test
  public void testCellOverflowString ()
  {
    final WorkbookCreationHelper aWBC = new WorkbookCreationHelper (EExcelVersion.XLSX);
    aWBC.createNewSheet ();
    aWBC.addRow ();
    Cell aCell;
    aCell = aWBC.addCell (getSmallContent ());
    Assert.assertEquals (getSmallContent (), aCell.getStringCellValue ());
    aCell = aWBC.addCell (getMaxContent ());
    Assert.assertEquals (getMaxContent (), aCell.getStringCellValue ());
    aCell = aWBC.addCell (getMaxContent () + "NARF");
    Assert.assertEquals (getMaxContent (), aCell.getStringCellValue ());
  }

  @Test
  public void testCellOverflowRichText ()
  {
    final WorkbookCreationHelper aWBC = new WorkbookCreationHelper (EExcelVersion.XLSX);
    aWBC.createNewSheet ();
    aWBC.addRow ();
    Cell aCell;
    
    aCell = aWBC.addCell (createRichText (getSmallContent (), aWBC));
    Assert.assertEquals (createRichText (getSmallContent (), aWBC).getString (),
                         aCell.getRichStringCellValue ().getString ());
    aCell = aWBC.addCell (createRichText (getMaxContent (), aWBC));
    Assert.assertEquals (createRichText (getMaxContent (), aWBC).getString (),
                         aCell.getRichStringCellValue ().getString ());
    aCell = aWBC.addCell (createRichText (getMaxContent () + "NARF", aWBC));
    Assert.assertEquals (createRichText (getMaxContent (), aWBC).getString (),
                         aCell.getRichStringCellValue ().getString ());
  }

  private static XSSFRichTextString createRichText (final String sString,
                                                    final WorkbookCreationHelper aWBC)
  {
    final XSSFRichTextString aRT = new XSSFRichTextString (sString);
    final Font aFont = aWBC.createFont ();
    aFont.setItalic (true);
    aRT.applyFont (1, 3, aFont);
    return aRT;
  }
}
