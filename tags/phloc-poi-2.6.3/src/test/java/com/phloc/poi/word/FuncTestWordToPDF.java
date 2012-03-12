/**
 * Copyright (C) 2006-2012 phloc systems
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
package com.phloc.poi.word;

import java.io.File;

import org.apache.poi.hwpf.converter.WordToFoConverter;
import org.junit.Test;

public class FuncTestWordToPDF
{
  @Test
  public void testToPDF ()
  {
    final File aSrcFile = new File ("src/test/resources/word/test1.doc");
    final File aFOFile = new File (aSrcFile.getParentFile (), "test1.fo");
    WordToFoConverter.main (new String [] { aSrcFile.getAbsolutePath (), aFOFile.getAbsolutePath () });
  }
}
