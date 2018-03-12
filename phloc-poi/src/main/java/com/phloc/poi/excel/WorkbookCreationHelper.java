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

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Calendar;
import java.util.Date;
import java.util.Set;

import javax.annotation.Nonnegative;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.WillClose;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.joda.time.DateTime;
import org.joda.time.LocalDate;
import org.joda.time.LocalDateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.phloc.commons.ValueEnforcer;
import com.phloc.commons.collections.ContainerHelper;
import com.phloc.commons.io.file.FileUtils;
import com.phloc.commons.io.streams.StreamUtils;
import com.phloc.commons.state.ESuccess;
import com.phloc.datetime.CPDT;
import com.phloc.poi.excel.style.ExcelStyle;
import com.phloc.poi.excel.style.ExcelStyleCache;

/**
 * A utility class for creating very simple Excel workbooks.
 * 
 * @author Philip Helger
 */
public final class WorkbookCreationHelper
{
  /** The BigInteger for the largest possible long value */
  private static final BigInteger BIGINT_MAX_LONG = BigInteger.valueOf (Long.MAX_VALUE);
  
  /** The BigInteger for the smallest possible long value */
  private static final BigInteger BIGINT_MIN_LONG = BigInteger.valueOf (Long.MIN_VALUE);
  
  private static final Logger LOG = LoggerFactory.getLogger (WorkbookCreationHelper.class);
  
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
    this.m_aWB = ValueEnforcer.notNull (aWB, "Workbook"); //$NON-NLS-1$
    this.m_aCreationHelper = aWB.getCreationHelper ();
  }
  
  @Nonnull
  public Workbook getWorkbook ()
  {
    return this.m_aWB;
  }
  
  /**
   * Create a new font in the passed workbook.
   * 
   * @return The created font.
   */
  @Nonnull
  public Font createFont ()
  {
    return this.m_aWB.createFont ();
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
   *          The name to be used. May be <code>null</code>.
   * @return The created workbook sheet
   */
  @Nonnull
  public Sheet createNewSheet (@Nullable final String sName)
  {
    this.m_aLastSheet = sName == null ? this.m_aWB.createSheet () : this.m_aWB.createSheet (sName);
    this.m_nLastSheetRowIndex = 0;
    this.m_aLastRow = null;
    this.m_nLastRowCellIndex = 0;
    this.m_aLastCell = null;
    this.m_nMaxCellIndex = 0;
    return this.m_aLastSheet;
  }
  
  public void goTo (final Row aRow)
  {
    this.m_aLastRow = aRow;
    this.m_nLastSheetRowIndex = this.m_aLastRow.getRowNum () + 1;
  }
  
  public void fillDenormatizationValues (final Set <Integer> aColIndexes)
  {
    final Row aMainRow = this.m_aLastRow;
    // cache last row idx?
    for (int i = aMainRow.getRowNum (); i < this.m_aLastSheet.getLastRowNum (); i++)
    {
      final Row aCurRow = nextRow ();
      if (!ContainerHelper.isEmpty (aColIndexes))
      {
        for (final Integer aColIdx : aColIndexes)
        {
          final Cell aSourceCell = aMainRow.getCell (aColIdx.intValue ());
          copyToRow (aSourceCell, aCurRow);
        }
      }
    }
    goToLastRow ();
  }
  
  public Cell copyToRow (final Cell aSource, final Row aRow)
  {
    final Cell aCopy = aRow.createCell (aSource.getColumnIndex ());
    apply (aSource, aCopy);
    return aCopy;
  }
  
  @SuppressWarnings ("static-method")
  public void apply (final Cell aSource, final Cell aTarget)
  {
    // not cloning the style in order to not create uncached instances (style overflow!)
    
    // CellStyle aStyle = this.m_aWB.createCellStyle ();
    // aStyle.cloneStyleFrom (aSource.getCellStyle ());
    
    final CellStyle aStyle = aSource.getCellStyle ();
    if (aStyle != null)
    {
      aTarget.setCellStyle (aStyle);
    }
    // If there is a cell comment, copy
    if (aSource.getCellComment () != null)
    {
      aTarget.setCellComment (aSource.getCellComment ());
    }
    
    // If there is a cell hyperlink, copy
    if (aSource.getHyperlink () != null)
    {
      aTarget.setHyperlink (aSource.getHyperlink ());
    }
    
    // Set the cell data type
    aTarget.setCellType (aSource.getCellType ());
    
    // Set the cell data value
    switch (aSource.getCellType ())
    {
      case Cell.CELL_TYPE_BLANK:
        aTarget.setCellValue (aSource.getStringCellValue ());
        break;
      case Cell.CELL_TYPE_BOOLEAN:
        aTarget.setCellValue (aSource.getBooleanCellValue ());
        break;
      case Cell.CELL_TYPE_ERROR:
        aTarget.setCellErrorValue (aSource.getErrorCellValue ());
        break;
      case Cell.CELL_TYPE_FORMULA:
        aTarget.setCellFormula (aSource.getCellFormula ());
        break;
      case Cell.CELL_TYPE_NUMERIC:
        aTarget.setCellValue (aSource.getNumericCellValue ());
        break;
      case Cell.CELL_TYPE_STRING:
        aTarget.setCellValue (aSource.getRichStringCellValue ());
        break;
    }
  }
  
  public void goTo (final Cell aCell)
  {
    goTo (aCell.getRow ());
    this.m_aLastCell = aCell;
  }
  
  /**
   * @return Creates a new row in this sheet if not yet existing after the current one, or returns
   *         the existing next row
   */
  @Nonnull
  public Row nextRow ()
  {
    final Row aNext = this.m_aLastSheet.getRow (this.m_nLastSheetRowIndex);
    if (aNext == null)
    {
      return addRow ();
    }
    this.m_aLastRow = aNext;
    this.m_nLastSheetRowIndex++;
    goToLastCell ();
    return this.m_aLastRow;
  }
  
  /**
   * @return Creates a new row in this sheet if not yet existing after the current one, or returns
   *         the existing next row
   */
  @Nullable
  public Row goToLastRow ()
  {
    this.m_aLastRow = this.m_aLastSheet.getRow (this.m_aLastSheet.getLastRowNum ());
    this.m_nLastSheetRowIndex = this.m_aLastRow == null ? 0 : this.m_aLastRow.getRowNum () + 1;
    goToLastCell ();
    return this.m_aLastRow;
  }
  
  /**
   * @return Creates a new row in this sheet if not yet existing after the current one, or returns
   *         the existing next row
   */
  @Nullable
  public Cell goToLastCell ()
  {
    if (this.m_aLastRow == null)
    {
      this.m_nLastRowCellIndex = 0;
      this.m_aLastCell = null;
    }
    else
    {
      // getLastCellNum returns index + 1 or -1
      this.m_nLastRowCellIndex = this.m_aLastRow.getLastCellNum () - 1;
      if (this.m_nLastRowCellIndex < 0)
      {
        this.m_nLastRowCellIndex = 0;
      }
      this.m_aLastCell = this.m_aLastRow.getCell (this.m_nLastRowCellIndex);
    }
    return this.m_aLastCell;
  }
  
  /**
   * @return A new row in the current sheet.
   */
  @Nonnull
  public Row addRow ()
  {
    this.m_aLastRow = this.m_aLastSheet.createRow (this.m_nLastSheetRowIndex++);
    this.m_nLastRowCellIndex = 0;
    this.m_aLastCell = null;
    return this.m_aLastRow;
  }
  
  /**
   * @return The number of rows in the current sheet, 0-based.
   */
  @Nonnegative
  public int getRowCount ()
  {
    return this.m_nLastSheetRowIndex;
  }
  
  /**
   * @return A new cell in the current row of the current sheet
   */
  @Nonnull
  public Cell addCell ()
  {
    return addCell ((ExcelStyle) null);
  }
  
  /**
   * @return A new cell in the current row of the current sheet
   */
  @Nonnull
  public Cell addCell (@Nullable final ExcelStyle aStyle)
  {
    this.m_aLastCell = this.m_aLastRow.createCell (this.m_nLastRowCellIndex++);
    
    // Check for the maximum cell index in this sheet
    if (this.m_nLastRowCellIndex > this.m_nMaxCellIndex)
    {
      this.m_nMaxCellIndex = this.m_nLastRowCellIndex;
    }
    if (aStyle != null)
    {
      addCellStyle (aStyle);
    }
    return this.m_aLastCell;
  }
  
  /**
   * @param bValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final boolean bValue)
  {
    return addCell (bValue, null);
  }
  
  @Nonnull
  public Cell addCell (final boolean bValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_BOOLEAN);
    aCell.setCellValue (bValue);
    return aCell;
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final Calendar aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (final Calendar aValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (aValue);
    return aCell;
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final Date aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (final Date aValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (aValue);
    return aCell;
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (@Nonnull final LocalDate aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (@Nonnull final LocalDate aValue, @Nullable final ExcelStyle aStyle)
  {
    return addCell (aValue.toDateTime (CPDT.NULL_LOCAL_TIME), aStyle);
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (@Nonnull final LocalDateTime aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (@Nonnull final LocalDateTime aValue, @Nullable final ExcelStyle aStyle)
  {
    return addCell (aValue.toDateTime (), aStyle);
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (@Nonnull final DateTime aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (@Nonnull final DateTime aValue, @Nullable final ExcelStyle aStyle)
  {
    return addCell (aValue.toDate (), aStyle);
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (@Nonnull final BigInteger aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (@Nonnull final BigInteger aValue, @Nullable final ExcelStyle aStyle)
  {
    if (aValue.compareTo (BIGINT_MIN_LONG) >= 0 && aValue.compareTo (BIGINT_MAX_LONG) <= 0)
    {
      return addCell (aValue.longValue (), aStyle);
    }
    
    // Too large - add as string
    return addCell (aValue.toString (), aStyle);
  }
  
  /**
   * @param dValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final double dValue)
  {
    return addCell (dValue, null);
  }
  
  @Nonnull
  public Cell addCell (final double dValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_NUMERIC);
    aCell.setCellValue (dValue);
    return aCell;
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (@Nonnull final BigDecimal aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (@Nonnull final BigDecimal aValue, @Nullable final ExcelStyle aStyle)
  {
    try
    {
      return addCell (aValue.doubleValue (), aStyle);
    }
    catch (final NumberFormatException ex)
    {
      // Add as string if too large for a double
      return addCell (aValue.toString (), aStyle);
    }
  }
  
  /**
   * @param aValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final RichTextString aValue)
  {
    return addCell (aValue, null);
  }
  
  @Nonnull
  public Cell addCell (final RichTextString aValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_STRING);
    aCell.setCellValue (aValue);
    return aCell;
  }
  
  /**
   * @param sValue
   *          The value to be set.
   * @return A new cell in the current row of the current sheet with the passed value
   */
  @Nonnull
  public Cell addCell (final String sValue)
  {
    return addCell (sValue, null);
  }
  
  @Nonnull
  public Cell addCell (final String sValue, @Nullable final ExcelStyle aStyle)
  {
    final Cell aCell = addCell (aStyle);
    aCell.setCellType (Cell.CELL_TYPE_STRING);
    aCell.setCellValue (sValue);
    return aCell;
  }
  
  /**
   * @param sFormula
   *          The formula to be set. May be <code>null</code> to set no formula.
   * @return A new cell in the current row of the current sheet with the passed formula
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
   *          The style to be set.
   */
  public void addCellStyle (@Nonnull final ExcelStyle aExcelStyle)
  {
    ValueEnforcer.notNull (aExcelStyle, "ExcelStyle"); //$NON-NLS-1$
    if (this.m_aLastCell == null)
    {
      throw new IllegalStateException ("No cell present for current row!"); //$NON-NLS-1$
    }
    this.m_aLastCell.setCellStyle (getOrCreateCellStyle (aExcelStyle));
  }
  
  public CellStyle getOrCreateCellStyle (@Nonnull final ExcelStyle aExcelStyle)
  {
    ValueEnforcer.notNull (aExcelStyle, "ExcelStyle"); //$NON-NLS-1$
    CellStyle aCellStyle = this.m_aStyleCache.getCellStyle (aExcelStyle);
    if (aCellStyle == null)
    {
      aCellStyle = this.m_aWB.createCellStyle ();
      aExcelStyle.fillCellStyle (this.m_aWB, aCellStyle, this.m_aCreationHelper);
      this.m_aStyleCache.addCellStyle (aExcelStyle, aCellStyle);
      this.m_nCreatedCellStyles++;
    }
    return aCellStyle;
  }
  
  /**
   * @return The number of cells in the current row in the current sheet, 0-based
   */
  @Nonnegative
  public int getCellCountInRow ()
  {
    return this.m_nLastRowCellIndex;
  }
  
  public void setLastRowCellIndex (@Nonnegative final int nIndex)
  {
    this.m_nLastRowCellIndex = nIndex;
  }
  
  /**
   * @return The maximum number of cells in a single row in the current sheet, 0-based.
   */
  @Nonnegative
  public int getMaximumCellCountInRowInSheet ()
  {
    return this.m_nMaxCellIndex;
  }
  
  /**
   * Auto size all columns to be matching width in the current sheet
   */
  public void autoSizeAllColumns ()
  {
    // auto-adjust all columns (except description and image description)
    for (short nCol = 0; nCol < this.m_nMaxCellIndex; ++nCol)
    {
      try
      {
        this.m_aLastSheet.autoSizeColumn (nCol);
      }
      catch (final IllegalArgumentException ex)
      {
        // Happens if a column is too large
        LOG.warn ("Failed to resize column " + nCol + ": column too wide!"); //$NON-NLS-1$ //$NON-NLS-2$
      }
    }
  }
  
  /**
   * Add an auto filter on the first row on all columns in the current sheet.
   */
  public void autoFilterAllColumns ()
  {
    autoFilterAllColumns (0);
  }
  
  /**
   * @param nRowIndex
   *          The 0-based index of the row, where to set the filter. Add an auto filter on all
   *          columns in the current sheet.
   */
  public void autoFilterAllColumns (@Nonnegative final int nRowIndex)
  {
    // Set auto filter on all columns
    // Use the specified row (param1, param2)
    // From first column to last column (param3, param4)
    this.m_aLastSheet.setAutoFilter (new CellRangeAddress (nRowIndex,
                                                           nRowIndex,
                                                           0,
                                                           this.m_nMaxCellIndex - 1));
  }
  
  /**
   * Write the current workbook to a file
   * 
   * @param sFilename
   *          The file to write to. May not be <code>null</code>.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess write (@Nonnull final String sFilename)
  {
    return write (FileUtils.getOutputStream (sFilename));
  }
  
  /**
   * Write the current workbook to a file
   * 
   * @param aFile
   *          The file to write to. May not be <code>null</code>.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess write (@Nonnull final File aFile)
  {
    return write (FileUtils.getOutputStream (aFile));
  }
  
  /**
   * Write the current workbook to an output stream.
   * 
   * @param aOS
   *          The output stream to write to. May not be <code>null</code>. Is automatically closed
   *          independent of the success state.
   * @return {@link ESuccess}
   */
  @Nonnull
  public ESuccess write (@Nonnull @WillClose final OutputStream aOS)
  {
    ValueEnforcer.notNull (aOS, "OutputStream"); //$NON-NLS-1$
    
    try
    {
      if (this.m_nCreatedCellStyles > 0 && LOG.isDebugEnabled ())
      {
        LOG.debug ("Writing Excel workbook with " + //$NON-NLS-1$
                   this.m_nCreatedCellStyles +
                   " different cell styles"); //$NON-NLS-1$
      }
      
      this.m_aWB.write (aOS);
      return ESuccess.SUCCESS;
    }
    catch (final IOException ex)
    {
      if (!StreamUtils.isKnownEOFException (ex))
      {
        LOG.error ("Failed to write Excel workbook to output stream " + aOS, ex); //$NON-NLS-1$
      }
      return ESuccess.FAILURE;
    }
    finally
    {
      StreamUtils.close (aOS);
    }
  }
}
