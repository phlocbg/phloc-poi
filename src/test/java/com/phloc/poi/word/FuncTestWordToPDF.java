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
