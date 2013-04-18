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
import java.io.InputStream;

import javax.annotation.CheckForSigned;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.WillClose;

import org.apache.poi.POIXMLException;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.phloc.commons.CGlobal;
import com.phloc.commons.mime.CMimeType;
import com.phloc.commons.mime.IMimeType;
import com.phloc.poi.POISetup;

/**
 * Encapsulates the main differences between the different excel versions.
 * 
 * @author philip
 */
public enum EExcelVersion
{
  XLS
  {
    @Override
    @Nonnull
    public HSSFWorkbook createWorkbook ()
    {
      return new HSSFWorkbook ();
    }

    @Override
    @Nullable
    public HSSFWorkbook readWorkbook (@Nonnull @WillClose final InputStream aIS)
    {
      try
      {
        // Closes the input stream internally
        return new HSSFWorkbook (aIS);
      }
      catch (final OfficeXmlFileException ex)
      {
        // No XLS
        return null;
      }
      catch (final IOException ex)
      {
        return null;
      }
    }

    @Override
    @Nonnull
    public HSSFRichTextString createRichText (final String sValue)
    {
      return new HSSFRichTextString (sValue);
    }

    @Override
    @Nonnull
    public String getFileExtension ()
    {
      return ".xls";
    }

    @Override
    @Nonnull
    public IMimeType getMimeType ()
    {
      return CMimeType.APPLICATION_MS_EXCEL;
    }

    @Override
    public boolean hasRowLimitPerSheet ()
    {
      return true;
    }

    @Override
    public int getRowLimitPerSheet ()
    {
      // Max row limit is 65536. The index is therefore 0up to (limit-1)
      return 65536;
    }
  },
  XLSX
  {
    @Override
    @Nonnull
    public XSSFWorkbook createWorkbook ()
    {
      return new XSSFWorkbook ();
    }

    @Override
    @Nullable
    public XSSFWorkbook readWorkbook (@Nonnull @WillClose final InputStream aIS)
    {
      try
      {
        // Closes the input stream internally
        return new XSSFWorkbook (aIS);
      }
      catch (final POIXMLException ex)
      {
        // No XLSX
        return null;
      }
      catch (final IOException ex)
      {
        return null;
      }
    }

    @Override
    @Nonnull
    public XSSFRichTextString createRichText (final String sValue)
    {
      return new XSSFRichTextString (sValue);
    }

    @Override
    @Nonnull
    public String getFileExtension ()
    {
      return ".xlsx";
    }

    @Override
    @Nonnull
    public IMimeType getMimeType ()
    {
      return CMimeType.APPLICATION_MS_EXCEL_2007;
    }

    @Override
    public boolean hasRowLimitPerSheet ()
    {
      return false;
    }

    @Override
    public int getRowLimitPerSheet ()
    {
      return CGlobal.ILLEGAL_UINT;
    }
  };

  static
  {
    POISetup.initOnDemand ();
  }

  /**
   * @return A newly created workbook of this version
   */
  @Nonnull
  public abstract Workbook createWorkbook ();

  /**
   * Open an existing work book for reading.
   * 
   * @param aIS
   *        The input stream to read from. May not be <code>null</code>.
   * @return <code>null</code> in case the workbook cannot be opened.
   */
  @Nullable
  public abstract Workbook readWorkbook (@Nonnull InputStream aIS);

  @Nonnull
  public abstract RichTextString createRichText (String sValue);

  /**
   * @return The file extension incl. the leading dot.
   */
  @Nonnull
  public abstract String getFileExtension ();

  /**
   * @return The MIME type for this excel version.
   */
  @Nonnull
  public abstract IMimeType getMimeType ();

  /**
   * @return <code>true</code> if this Excel version has a row limit inside a
   *         sheet
   */
  public abstract boolean hasRowLimitPerSheet ();

  /**
   * @return the maximum number of rows per sheet (incl.) or
   *         {@link CGlobal#ILLEGAL_UINT} if no limit exists
   */
  @CheckForSigned
  public abstract int getRowLimitPerSheet ();
}
