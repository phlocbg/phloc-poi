/**
 * Copyright (C) 2006-2013 phloc systems
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

import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

import javax.annotation.Nonnegative;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.WillClose;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.phloc.commons.io.streams.StreamUtils;
import com.phloc.commons.state.ESuccess;
import com.phloc.poi.excel.style.ExcelStyle;
import com.phloc.poi.excel.style.ExcelStyleCache;

/**
 * A utility class for creating very simple Excel workbooks.
 * 
 * @author philip
 */
public final class WorkbookCreationHelper
{
  private static final Logger s_aLogger = LoggerFactory.getLogger (WorkbookCreationHelper.class);

  private final Workbook m_aWB;
  private final CreationHelper m_aCreationHelper;
  private final ExcelStyleCache m_aStyleCache = new ExcelStyleCache ();
  private Sheet m_aLastSheet;
  private int m_nLastSheetRowIndex = 0;
  private Row m_aLastRow;
  private int m_nLastRowCellIndex = 0;
  private Cell m_aLastCell;
  private int m_nMaxCellIndex = 0;
  private int m_nCreatedCellStyles = 0;

  public WorkbookCreationHelper (@Nonnull final EExcelVersion eVersion)
  {
    this (eVersion.createWorkbook ());
  }

  public WorkbookCreationHelper (@Nonnull final Workbook aWB)
  {
    if (aWB == null)
      throw new NullPointerException ("workbook");
    m_aWB = aWB;
    m_aCreationHelper = aWB.getCreationHelper ();
  }

  @Nonnull
  public Workbook getWorkbook ()
  {
    return m_aWB;
  }

  /**
   * @return A new sheet with a default name
   */
  @Nonnull
  public Sheet createNewSheet ()
  {
    return createNewSheet (null);
  }

  /**
   * Create a new sheet with an optional name
   * 
   * @param sName
   *        The name to be used. May be <code>null</code>.
   * @return The created workbook sheet
   */
  @Nonnull
  public Sheet createNewSheet (@Nullable final String sName)
  {
    m_aLastSheet = sName == null ? m_aWB.createSheet () : m_aWB.createSheet (sName);
    m_nLastSheetRowIndex = 0;
    m_aLastRow = null;
    m_nLastRowCellIndex = 0;
    m_aLastCell = null;
    m_nMaxCellIndex = 0;
    return m_aLastSheet;
  }

  /**
   * @return A new row in the current sheet.
   */
  @Nonnull
  public Row addRow ()
  {
    m_aLastRow = m_aLastSheet.createRow (m_nLastSheetRowIndex++);
    m_nLastRowCellIndex = 0;
    m_aLastCell = null;
    return m_aLastRow;
  }

  /**
   * @return The number of rows in the current sheet, 0-based.
   */
  @Nonnegative
  public int getRowCount ()
  {
    return m_nLastSheetRowIndex;
  }

  /**
   * @return A new cell in the current row of the current sheet
   */
  @Nonnull
  public Cell addCell ()
  {
    m_aLastCell = m_aLastRow.createCell (m_nLastRowCellIndex++);

    // Check for the maximum cell index in this sheet
    if (m_nLastRowCellIndex > m_nMaxCellIndex)
      m_nMaxCellIndex = m_nLastRowCellIndex;
    return m_aLastCell;
  }

  /**
   * @param bValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final boolean bValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_BOOLEAN);
    aCell.setCellValue (bValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final Calendar aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final Date aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param dValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final double dValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (dValue);
    return aCell;
  }

  /**
   * @param aValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final RichTextString aValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_STRING);
    aCell.setCellValue (aValue);
    return aCell;
  }

  /**
   * @param sValue
   *        The value to be set.
   * @return A new cell in the current row of the current sheet with the passed
   *         value
   */
  @Nonnull
  public Cell addCell (final String sValue)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_STRING);
    aCell.setCellValue (sValue);
    return aCell;
  }

  /**
   * @param sFormula
   *        The formula to be set. May be <code>null</code> to set no formula.
   * @return A new cell in the current row of the current sheet with the passed
   *         formula
   */
  @Nonnull
  public Cell addCellFormula (@Nullable final String sFormula)
  {
    final Cell aCell = addCell ();
    aCell.setCellType (Cell.CELL_TYPE_FORMULA);
    aCell.setCellFormula (sFormula);
    return aCell;
  }

  /**
   * Set the cell style of the last added cell
   * 
   * @param aExcelStyle
   *        The style to be set.
   */
  public void addCellStyle (@Nonnull final ExcelStyle aExcelStyle)
  {
    if (aExcelStyle == null)
      throw new NullPointerException ("excelStyle");
    if (m_aLastCell == null)
      throw new IllegalStateException ("No cell present for current row!");

    CellStyle aCellStyle = m_aStyleCache.getCellStyle (aExcelStyle);
    if (aCellStyle == null)
    {
      aCellStyle = m_aWB.createCellStyle ();
      aExcelStyle.fillCellStyle (m_aWB, aCellStyle, m_aCreationHelper);
      m_aStyleCache.addCellStyle (aExcelStyle, aCellStyle);
      m_nCreatedCellStyles++;
    }
    m_aLastCell.setCellStyle (aCellStyle);
  }

  /**
   * @return The number of cells in the current row in the current sheet,
   *         0-based
   */
  @Nonnegative
  public int getCellCountInRow ()
  {
    return m_nLastRowCellIndex;
  }

  /**
   * @return The maximum number of cells in a single row in the current sheet,
   *         0-based.
   */
  @Nonnegative
  public int getMaximumCellCountInRowInSheet ()
  {
    return m_nMaxCellIndex;
  }

  /**
   * Auto size all columns to be matching width in the current sheet
   */
  public void autoSizeAllColumns ()
  {
    // auto-adjust all columns (except description and image description)
    for (short nCol = 0; nCol < m_nMaxCellIndex; ++nCol)
      try
      {
        m_aLastSheet.autoSizeColumn (nCol);
      }
      catch (final IllegalArgumentException ex)
      {
        // Happens if a column is too large
        s_aLogger.warn ("Failed to resize column " + nCol + ": column too wide!");
      }
  }

  /**
   * Add an auto filter on all columns in the current sheet.
   */
  public void autoFilterAllColumns ()
  {
    // Set auto filter on all columns
    // Always on the first row (param1, param2)
    // From first column to last column (param3, param4)
    m_aLastSheet.setAutoFilter (new CellRangeAddress (0, 0, 0, m_nMaxCellIndex - 1));
  }

  /**
   * Write the current workbook to an output stream.
   * 
   * @param aOS
   *        The output stream to write to. May not be <code>null</code>. Is
   *        automatically closed independent of the success state.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess write (@Nonnull @WillClose final OutputStream aOS)
  {
    if (aOS == null)
      throw new NullPointerException ("outputStream");

    try
    {
      if (m_nCreatedCellStyles > 0 && s_aLogger.isDebugEnabled ())
        s_aLogger.debug ("Writing Excel workbook with " + m_nCreatedCellStyles + " different cell styles");

      m_aWB.write (aOS);
      return ESuccess.SUCCESS;
    }
    catch (final IOException ex)
    {
      if (!StreamUtils.isKnownEOFException (ex))
        s_aLogger.error ("Failed to write Excel workbook to output stream " + aOS, ex);
      return ESuccess.FAILURE;
    }
    finally
    {
      StreamUtils.close (aOS);
    }
  }
}
