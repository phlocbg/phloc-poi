package com.phloc.poi.excel;

import java.io.File;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookWrapper implements IWorkbookWrapper
{
  private Workbook m_aWorkbook;
  private String m_sFileName = null;
  private File m_aFile = null;
  
  public WorkbookWrapper (@Nonnull Workbook aWorkbook)
  {
    if (aWorkbook == null)
    {
      throw new NullPointerException ("aWorkbook"); //$NON-NLS-1$
    }
    this.m_aWorkbook = aWorkbook;
  }
  
  @Override
  @Nonnull
  public Workbook getWorkbook ()
  {
    return this.m_aWorkbook;
  }
  
  @Override
  @Nullable
  public File getFile ()
  {
    return this.m_aFile;
  }
  
  @Nullable
  public IWorkbookWrapper setFile (@Nullable File aFile)
  {
    m_aFile = aFile;
    return this;
  }
  
  @Override
  @Nullable
  public String getFileName ()
  {
    return this.m_sFileName;
  }
  
  @Nullable
  public IWorkbookWrapper setFileName (@Nullable String sFileName)
  {
    m_sFileName = sFileName;
    return this;
  }
}
