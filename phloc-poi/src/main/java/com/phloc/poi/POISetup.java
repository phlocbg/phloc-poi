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
package com.phloc.poi;

import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

import javax.annotation.Nonnegative;
import javax.annotation.concurrent.ThreadSafe;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.phloc.commons.SystemProperties;
import com.phloc.commons.ValueEnforcer;

/**
 * This class can be used to initialize POI to work best with the phloc stack.
 *
 * @author Boris Gregorcic
 */
@ThreadSafe
public final class POISetup
{
  public static final String SYS_PROP_POI_LOGGER = "org.apache.poi.util.POILogger"; //$NON-NLS-1$
  private static final Logger LOG = LoggerFactory.getLogger (POISetup.class);
  public static final int DEFAULT_WINDOW_SIZE = 100;
  private static final AtomicBoolean s_aInited = new AtomicBoolean (false);
  private static AtomicInteger WINDOW_SIZE = new AtomicInteger (DEFAULT_WINDOW_SIZE);
  private static AtomicBoolean COMPRESSED_TEMP_FILES = new AtomicBoolean (true);
  
  private POISetup ()
  {}
  
  public static void enableCustomLogger (final boolean bEnable)
  {
    if (bEnable)
      SystemProperties.setPropertyValue (SYS_PROP_POI_LOGGER, POISLF4JLogger.class.getName ());
    else
      SystemProperties.removePropertyValue (SYS_PROP_POI_LOGGER);
  }
  
  public static void initOnDemand ()
  {
    if (s_aInited.compareAndSet (false, true))
    {
      enableCustomLogger (true);
    }
  }
  
  /**
   * Sets the window size for POI SXSSF streaming model
   * 
   * @param nWindowSize
   *          the window size to use (default: {@value #DEFAULT_WINDOW_SIZE})
   */
  public static void setWindowSize (@Nonnegative final int nWindowSize)
  {
    ValueEnforcer.isGT0 (nWindowSize, "window size"); //$NON-NLS-1$
    LOG.info ("Setting POI window size to {}", String.valueOf (nWindowSize)); //$NON-NLS-1$
    WINDOW_SIZE.set (nWindowSize);
  }
  
  /**
   * @return The window size for POI SXSSF streaming model
   */
  public static int getWindowSize ()
  {
    return WINDOW_SIZE.get ();
  }
  
  public static void setCompressTempFiles (final boolean bCompress)
  {
    COMPRESSED_TEMP_FILES.set (bCompress);
  }
  
  public static boolean isCompressTempFiles ()
  {
    return COMPRESSED_TEMP_FILES.get ();
  }
}
