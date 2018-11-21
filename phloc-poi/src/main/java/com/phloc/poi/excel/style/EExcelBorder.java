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

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Excel border styles.
 * 
 * @author Boris Gregorcic
 */
public enum EExcelBorder
{
  BORDER_NONE (BorderStyle.NONE),
  BORDER_THIN (BorderStyle.THIN),
  BORDER_MEDIUM (BorderStyle.MEDIUM),
  BORDER_DASHED (BorderStyle.DASHED),
  BORDER_HAIR (BorderStyle.HAIR),
  BORDER_THICK (BorderStyle.THICK),
  BORDER_DOUBLE (BorderStyle.DOUBLE),
  BORDER_DOTTED (BorderStyle.DOTTED),
  BORDER_MEDIUM_DASHED (BorderStyle.MEDIUM_DASHED),
  BORDER_DASH_DOT (BorderStyle.DASH_DOT),
  BORDER_MEDIUM_DASH_DOT (BorderStyle.MEDIUM_DASH_DOT),
  BORDER_DASH_DOT_DOT (BorderStyle.DASH_DOT_DOT),
  BORDER_MEDIUM_DASH_DOT_DOT (BorderStyle.MEDIUM_DASH_DOT_DOT),
  BORDER_SLANTED_DASH_DOT (BorderStyle.SLANTED_DASH_DOT);

  private final BorderStyle m_eValue;

  private EExcelBorder (final BorderStyle eValue)
  {
    m_eValue = eValue;
  }

  public BorderStyle getValue ()
  {
    return m_eValue;
  }
}
