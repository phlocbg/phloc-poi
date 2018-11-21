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
package com.phloc.poi.excel.style;

import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel vertical alignment enum.
 * 
 * @author Boris Gregorcic
 */
public enum EExcelVerticalAlignment
{
  VERTICAL_TOP (VerticalAlignment.TOP),
  VERTICAL_CENTER (VerticalAlignment.CENTER),
  VERTICAL_BOTTOM (VerticalAlignment.BOTTOM),
  VERTICAL_JUSTIFY (VerticalAlignment.JUSTIFY);

  private final VerticalAlignment m_eValue;

  private EExcelVerticalAlignment (final VerticalAlignment eValue)
  {
    m_eValue = eValue;
  }

  public VerticalAlignment getValue ()
  {
    return m_eValue;
  }
}
