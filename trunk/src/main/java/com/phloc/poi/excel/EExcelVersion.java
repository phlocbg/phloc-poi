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

import javax.annotation.Nonnull;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.phloc.commons.mime.CMimeType;
import com.phloc.commons.mime.IMimeType;

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
    public Workbook createWorkbook ()
    {
      return new HSSFWorkbook ();
    }

    @Override
    @Nonnull
    public RichTextString createRichText (final String sValue)
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
  },
  XLSX
  {
    @Override
    @Nonnull
    public Workbook createWorkbook ()
    {
      return new XSSFWorkbook ();
    }

    @Override
    @Nonnull
    public RichTextString createRichText (final String sValue)
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
  };

  @Nonnull
  public abstract Workbook createWorkbook ();

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
}
