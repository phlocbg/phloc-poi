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

import javax.annotation.concurrent.ThreadSafe;

import com.phloc.commons.SystemProperties;

/**
 * This class can be used to initialize POI to work best with the phloc stack.
 *
 * @author Philip Helger
 */
@ThreadSafe
public final class POISetup
{
  public static final String SYS_PROP_POI_LOGGER = "org.apache.poi.util.POILogger";
  private static final AtomicBoolean s_aInited = new AtomicBoolean (false);

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
      enableCustomLogger (true);
  }
}
