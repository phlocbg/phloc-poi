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

import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel horizontal alignment enum.
 * 
 * @author Boris Gregorcic
 */
public enum EExcelAlignment
{
  ALIGN_GENERAL (HorizontalAlignment.GENERAL),
  ALIGN_LEFT (HorizontalAlignment.LEFT),
  ALIGN_CENTER (HorizontalAlignment.CENTER),
  ALIGN_RIGHT (HorizontalAlignment.RIGHT),
  ALIGN_FILL (HorizontalAlignment.FILL),
  ALIGN_JUSTIFY (HorizontalAlignment.JUSTIFY),
  ALIGN_CENTER_SELECTION (HorizontalAlignment.CENTER_SELECTION);

  private final HorizontalAlignment m_eValue;

  private EExcelAlignment (final HorizontalAlignment eValue)
  {
	  m_eValue = eValue;
  }

  public HorizontalAlignment getValue ()
  {
    return m_eValue;
  }
}
