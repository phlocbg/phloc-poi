/**
 * Copyright (C) 2006-2014 phloc systems
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

import java.util.HashMap;
import java.util.Map;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.CellStyle;

import com.phloc.commons.ValueEnforcer;
import com.phloc.commons.string.ToStringGenerator;

/**
 * A caching class that maps {@link ExcelStyle} objects to {@link CellStyle}.
 *
 * @author Philip Helger
 */
public final class ExcelStyleCache
{
  private final Map <ExcelStyle, CellStyle> m_aMap = new HashMap <ExcelStyle, CellStyle> ();

  public ExcelStyleCache ()
  {}

  @Nullable
  public CellStyle getCellStyle (@Nullable final ExcelStyle aExcelStyle)
  {
    return m_aMap.get (aExcelStyle);
  }

  public void addCellStyle (@Nonnull final ExcelStyle aExcelStyle, @Nonnull final CellStyle aCellStyle)
  {
    ValueEnforcer.notNull (aExcelStyle, "ExcelStyle");
    ValueEnforcer.notNull (aCellStyle, "CellStyle");

    m_aMap.put (aExcelStyle, aCellStyle);
  }

  @Override
  public String toString ()
  {
    return new ToStringGenerator (this).append ("map", m_aMap).toString ();
  }
}
