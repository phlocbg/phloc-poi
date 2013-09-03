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
package com.phloc.poi.excel.style;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

import com.phloc.commons.mock.PhlocTestUtils;

/**
 * Test class for class {@link ExcelStyle}.
 * 
 * @author Philip Helger
 */
public final class ExcelStyleTest
{
  @Test
  public void testBasic ()
  {
    final ExcelStyle e = new ExcelStyle ();
    PhlocTestUtils.testDefaultImplementationWithEqualContentObject (e, new ExcelStyle ());
    e.setBorder (EExcelBorder.BORDER_DASH_DOT);
    PhlocTestUtils.testDefaultImplementationWithDifferentContentObject (e, new ExcelStyle ());
    PhlocTestUtils.testDefaultImplementationWithEqualContentObject (e,
                                                                    new ExcelStyle ().setBorder (EExcelBorder.BORDER_DASH_DOT));
  }

  @Test
  public void testAlign ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getAlign ());
    PhlocTestUtils.testGetClone (e);
    for (final EExcelAlignment eAlign : EExcelAlignment.values ())
    {
      assertSame (e, e.setAlign (eAlign));
      assertEquals (eAlign, e.getAlign ());
      PhlocTestUtils.testGetClone (e);
    }
  }

  @Test
  public void testVerticalAlign ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getVerticalAlign ());
    PhlocTestUtils.testGetClone (e);
    for (final EExcelVerticalAlignment eAlign : EExcelVerticalAlignment.values ())
    {
      assertSame (e, e.setVerticalAlign (eAlign));
      assertEquals (eAlign, e.getVerticalAlign ());
      PhlocTestUtils.testGetClone (e);
    }
  }

  @Test
  public void testWrapText ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertTrue (e.isWrapText () == ExcelStyle.DEFAULT_WRAP_TEXT);
    PhlocTestUtils.testGetClone (e);
    assertSame (e, e.setWrapText (true));
    assertTrue (e.isWrapText ());
    PhlocTestUtils.testGetClone (e);
    assertSame (e, e.setWrapText (false));
    assertFalse (e.isWrapText ());
    PhlocTestUtils.testGetClone (e);
  }

  @Test
  public void testDataFormat ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getDataFormat ());
    PhlocTestUtils.testGetClone (e);
    assertSame (e, e.setDataFormat ("abc"));
    assertEquals ("abc", e.getDataFormat ());
    PhlocTestUtils.testGetClone (e);
    assertSame (e, e.setDataFormat (null));
    assertNull (e.getDataFormat ());
    PhlocTestUtils.testGetClone (e);
  }

  @Test
  public void testFillBackgroundColor ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillBackgroundColor ());
    PhlocTestUtils.testGetClone (e);
    for (final IndexedColors eColor : IndexedColors.values ())
    {
      assertSame (e, e.setFillBackgroundColor (eColor));
      assertEquals (eColor, e.getFillBackgroundColor ());
      PhlocTestUtils.testGetClone (e);
    }
  }

  @Test
  public void testFillForegroundColor ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillForegroundColor ());
    PhlocTestUtils.testGetClone (e);
    for (final IndexedColors eColor : IndexedColors.values ())
    {
      assertSame (e, e.setFillForegroundColor (eColor));
      assertEquals (eColor, e.getFillForegroundColor ());
      PhlocTestUtils.testGetClone (e);
    }
  }

  @Test
  public void testFillPattern ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getFillPattern ());
    PhlocTestUtils.testGetClone (e);
    for (final EExcelPattern ePattern : EExcelPattern.values ())
    {
      assertSame (e, e.setFillPattern (ePattern));
      assertEquals (ePattern, e.getFillPattern ());
      PhlocTestUtils.testGetClone (e);
    }
  }

  @Test
  public void testBorder ()
  {
    final ExcelStyle e = new ExcelStyle ();
    assertNull (e.getBorderTop ());
    assertNull (e.getBorderRight ());
    assertNull (e.getBorderBottom ());
    assertNull (e.getBorderLeft ());
    PhlocTestUtils.testGetClone (e);
    for (final EExcelBorder eBorder : EExcelBorder.values ())
    {
      assertSame (e, e.setBorder (eBorder));
      assertEquals (eBorder, e.getBorderTop ());
      assertEquals (eBorder, e.getBorderRight ());
      assertEquals (eBorder, e.getBorderBottom ());
      assertEquals (eBorder, e.getBorderLeft ());
      PhlocTestUtils.testGetClone (e);
    }
  }
}
