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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel vertical alignment enum.
 * 
 * @author philip
 */
public enum EExcelVerticalAlignment
{
  VERTICAL_TOP (CellStyle.VERTICAL_TOP),
  VERTICAL_CENTER (CellStyle.VERTICAL_CENTER),
  VERTICAL_BOTTOM (CellStyle.VERTICAL_BOTTOM),
  VERTICAL_JUSTIFY (CellStyle.VERTICAL_JUSTIFY);

  private short m_nValue;

  private EExcelVerticalAlignment (final short nValue)
  {
    m_nValue = nValue;
  }

  public short getValue ()
  {
    return m_nValue;
  }
}
