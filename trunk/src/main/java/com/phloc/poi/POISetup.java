package com.phloc.poi;

import java.util.concurrent.atomic.AtomicBoolean;

import org.apache.poi.util.POISLF4JLogger;

import com.phloc.commons.SystemProperties;

/**
 * This class can be used to initialize POI to work best with the phloc stack.
 * 
 * @author philip
 */
public final class POISetup
{
  private static final AtomicBoolean s_aInited = new AtomicBoolean (false);

  private POISetup ()
  {}

  public static void initOnDemand ()
  {
    if (s_aInited.compareAndSet (false, true))
    {
      SystemProperties.setPropertyValue ("org.apache.poi.util.POILogger", POISLF4JLogger.class.getName ());
    }
  }
}
