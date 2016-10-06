package com.phloc.poi.excel;

import java.io.File;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.Workbook;

public interface IWorkbookWrapper
{  
  /**
   * @return The workbook object representing the Excel
   */
  @Nonnull
  Workbook getWorkbook ();
  
  /**
   * @return A file object that denotes an absolute target location for serializing, may be
   *         <code>null</code>
   */
  @Nullable
  File getFile ();
  
  /**
   * @return A file name or relative path that should be used for further processing, may be <code>null</code>
   */
  @Nullable
  String getFileName ();  
}
