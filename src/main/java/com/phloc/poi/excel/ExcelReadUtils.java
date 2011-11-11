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

import java.util.Date;

import javax.annotation.Nullable;
import javax.annotation.concurrent.Immutable;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * Misc Excel read helper methods.
 * 
 * @author philip
 */
@Immutable
public final class ExcelReadUtils
{
  private ExcelReadUtils ()
  {}

  /**
   * Return the best matching Java object underlying the passed cell.<br>
   * Note: Date values cannot be determined automatically!
   * 
   * @param aCell
   *        The cell to be queried. May be <code>null</code>.
   * @return <code>null</code> if the cell is <code>null</code> or if it is of
   *         type blank.
   */
  @Nullable
  public static Object getCellValueObject (@Nullable final Cell aCell)
  {
    if (aCell == null)
      return null;

    final int nCellType = aCell.getCellType ();
    switch (nCellType)
    {
      case Cell.CELL_TYPE_NUMERIC:
      {
        final double dValue = aCell.getNumericCellValue ();
        if (dValue == (int) dValue)
        {
          // It's not a real double value, it's an int value
          return Integer.valueOf ((int) dValue);
        }
        if (dValue == (long) dValue)
        {
          // It's not a real double value, it's an int value
          return Long.valueOf ((long) dValue);
        }
        // It's a real floating point number
        // Using the BigDecimal (double) constructor leads to very weird
        // values compared to "Double.toString" which works!
        return Double.valueOf (dValue);
      }
      case Cell.CELL_TYPE_STRING:
        return aCell.getStringCellValue ();
      case Cell.CELL_TYPE_FORMULA:
        return aCell.getStringCellValue ();
      case Cell.CELL_TYPE_BLANK:
        return null;
      case Cell.CELL_TYPE_BOOLEAN:
        return Boolean.valueOf (aCell.getBooleanCellValue ());
      default:
        throw new IllegalArgumentException ("The cell type " + nCellType + " is unsupported!");
    }
  }

  @Nullable
  public static String getCellValueString (@Nullable final Cell aCell)
  {
    final Object aObject = getCellValueObject (aCell);
    return aObject == null ? null : aObject.toString ();
  }

  @Nullable
  public static Number getCellValueNumber (@Nullable final Cell aCell)
  {
    return (Number) getCellValueObject (aCell);
  }

  @Nullable
  public static Date getCellValueJavaDate (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getDateCellValue ();
  }

  @Nullable
  public static RichTextString getCellValueRichText (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getRichStringCellValue ();
  }

  @Nullable
  public static String getCellFormula (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getCellFormula ();
  }

  @Nullable
  public static Hyperlink getHyperlink (@Nullable final Cell aCell)
  {
    return aCell == null ? null : aCell.getHyperlink ();
  }
}
