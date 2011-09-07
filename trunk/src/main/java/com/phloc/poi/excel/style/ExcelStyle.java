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
package com.phloc.poi.excel.style;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;

import com.phloc.commons.ICloneable;
import com.phloc.commons.string.ToStringGenerator;
import com.phloc.commons.compare.CompareUtils;
import com.phloc.commons.hash.HashCodeGenerator;

public final class ExcelStyle implements ICloneable <ExcelStyle>
{
  private EExcelAlignment m_eAlign;
  private EExcelVerticalAlignment m_eVAlign;
  private boolean m_bWrapText = false;
  private String m_sDataFormat;
  private IndexedColors m_eFillBackgroundColor;
  private IndexedColors m_eFillForegroundColor;
  private EExcelPattern m_eFillPattern;
  private EExcelBorder m_eBorderTop;
  private EExcelBorder m_eBorderRight;
  private EExcelBorder m_eBorderBottom;
  private EExcelBorder m_eBorderLeft;

  public ExcelStyle ()
  {}

  public ExcelStyle (@Nonnull final ExcelStyle rhs)
  {
    m_bWrapText = rhs.m_bWrapText;
    m_sDataFormat = rhs.m_sDataFormat;
    m_eFillBackgroundColor = rhs.m_eFillBackgroundColor;
    m_eFillForegroundColor = rhs.m_eFillForegroundColor;
    m_eFillPattern = rhs.m_eFillPattern;
    m_eBorderTop = rhs.m_eBorderTop;
    m_eBorderRight = rhs.m_eBorderRight;
    m_eBorderBottom = rhs.m_eBorderBottom;
    m_eBorderLeft = rhs.m_eBorderLeft;
  }

  @Nonnull
  public ExcelStyle setDataFormat (@Nullable final EExcelAlignment eAlign)
  {
    m_eAlign = eAlign;
    return this;
  }

  @Nonnull
  public ExcelStyle setDataFormat (@Nullable final EExcelVerticalAlignment eVAlign)
  {
    m_eVAlign = eVAlign;
    return this;
  }

  @Nonnull
  public ExcelStyle setWrapText (final boolean bWrapText)
  {
    m_bWrapText = bWrapText;
    return this;
  }

  @Nonnull
  public ExcelStyle setDataFormat (@Nullable final String sDataFormat)
  {
    m_sDataFormat = sDataFormat;
    return this;
  }

  @Nonnull
  public ExcelStyle setFillBackgroundColor (@Nullable final IndexedColors eColor)
  {
    m_eFillBackgroundColor = eColor;
    return this;
  }

  @Nonnull
  public ExcelStyle setFillForegroundColor (@Nullable final IndexedColors eColor)
  {
    m_eFillForegroundColor = eColor;
    return this;
  }

  @Nonnull
  public ExcelStyle setFillPattern (@Nullable final EExcelPattern ePattern)
  {
    m_eFillPattern = ePattern;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorderTop (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderTop = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorderRight (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderRight = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorderBottom (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderBottom = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorderLeft (@Nullable final EExcelBorder eBorder)
  {
    m_eBorderLeft = eBorder;
    return this;
  }

  @Nonnull
  public ExcelStyle setBorder (@Nullable final EExcelBorder eBorder)
  {
    return setBorderTop (eBorder).setBorderRight (eBorder).setBorderBottom (eBorder).setBorderLeft (eBorder);
  }

  @Nonnull
  public ExcelStyle getClone ()
  {
    return new ExcelStyle (this);
  }

  public void fillCellStyle (@Nonnull final CellStyle aCS, @Nonnull final CreationHelper aCreationHelper)
  {
    if (m_eAlign != null)
      aCS.setAlignment (m_eAlign.getValue ());
    if (m_eVAlign != null)
      aCS.setVerticalAlignment (m_eVAlign.getValue ());
    aCS.setWrapText (m_bWrapText);
    if (m_sDataFormat != null)
      aCS.setDataFormat (aCreationHelper.createDataFormat ().getFormat (m_sDataFormat));
    if (m_eFillBackgroundColor != null)
      aCS.setFillBackgroundColor (m_eFillBackgroundColor.getIndex ());
    if (m_eFillForegroundColor != null)
      aCS.setFillForegroundColor (m_eFillForegroundColor.getIndex ());
    if (m_eFillPattern != null)
      aCS.setFillPattern (m_eFillPattern.getValue ());
    if (m_eBorderTop != null)
      aCS.setBorderTop (m_eBorderTop.getValue ());
    if (m_eBorderRight != null)
      aCS.setBorderRight (m_eBorderRight.getValue ());
    if (m_eBorderBottom != null)
      aCS.setBorderBottom (m_eBorderBottom.getValue ());
    if (m_eBorderLeft != null)
      aCS.setBorderLeft (m_eBorderLeft.getValue ());
  }

  @Override
  public boolean equals (final Object o)
  {
    if (o == this)
      return true;
    if (!(o instanceof ExcelStyle))
      return false;
    final ExcelStyle rhs = (ExcelStyle) o;
    return CompareUtils.nullSafeEquals (m_eAlign, rhs.m_eAlign) &&
           CompareUtils.nullSafeEquals (m_eVAlign, rhs.m_eVAlign) &&
           m_bWrapText == rhs.m_bWrapText &&
           CompareUtils.nullSafeEquals (m_sDataFormat, rhs.m_sDataFormat) &&
           CompareUtils.nullSafeEquals (m_eFillBackgroundColor, rhs.m_eFillBackgroundColor) &&
           CompareUtils.nullSafeEquals (m_eFillForegroundColor, rhs.m_eFillForegroundColor) &&
           CompareUtils.nullSafeEquals (m_eFillPattern, rhs.m_eFillPattern) &&
           CompareUtils.nullSafeEquals (m_eBorderTop, rhs.m_eBorderTop) &&
           CompareUtils.nullSafeEquals (m_eBorderRight, rhs.m_eBorderRight) &&
           CompareUtils.nullSafeEquals (m_eBorderBottom, rhs.m_eBorderBottom) &&
           CompareUtils.nullSafeEquals (m_eBorderLeft, rhs.m_eBorderLeft);
  }

  @Override
  public int hashCode ()
  {
    return new HashCodeGenerator (this).append (m_eAlign)
                                       .append (m_eVAlign)
                                       .append (m_bWrapText)
                                       .append (m_sDataFormat)
                                       .append (m_eFillBackgroundColor)
                                       .append (m_eFillForegroundColor)
                                       .append (m_eFillPattern)
                                       .append (m_eBorderTop)
                                       .append (m_eBorderRight)
                                       .append (m_eBorderBottom)
                                       .append (m_eBorderLeft)
                                       .getHashCode ();
  }

  @Override
  public String toString ()
  {
    return new ToStringGenerator (this).appendIfNotNull ("align", m_eAlign)
                                       .appendIfNotNull ("verticalAlign", m_eVAlign)
                                       .append ("wrapText", m_bWrapText)
                                       .appendIfNotNull ("dataFormat", m_sDataFormat)
                                       .appendIfNotNull ("fillBackgroundColor", m_eFillBackgroundColor)
                                       .appendIfNotNull ("fillForegroundColor", m_eFillForegroundColor)
                                       .appendIfNotNull ("fillPattern", m_eFillPattern)
                                       .appendIfNotNull ("borderTop", m_eBorderTop)
                                       .appendIfNotNull ("borderRight", m_eBorderRight)
                                       .appendIfNotNull ("borderBottom", m_eBorderBottom)
                                       .appendIfNotNull ("borderLeft", m_eBorderLeft)
                                       .toString ();
  }
}
