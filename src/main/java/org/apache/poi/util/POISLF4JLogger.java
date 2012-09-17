package org.apache.poi.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Set the system property "org.apache.poi.util.POILogger" to this class, to use
 * it. Must reside in this package, because the super class POILogger only has a
 * package-private ctor.
 * 
 * @author philip
 */
public class POISLF4JLogger extends POILogger
{
  private static final String BLA = "An exception occurred";
  private Logger log = null;

  public POISLF4JLogger ()
  {}

  @Override
  public void initialize (final String cat)
  {
    this.log = LoggerFactory.getLogger (cat);
  }

  /**
   * Log a message
   * 
   * @param level
   *        One of DEBUG, INFO, WARN, ERROR, FATAL
   * @param obj1
   *        The object to log.
   */
  @Override
  public void log (final int level, final Object obj1)
  {
    if (level == FATAL || level == ERROR)
    {
      if (log.isErrorEnabled ())
        log.error ("{}", obj1);
    }
    else
      if (level == WARN)
      {
        if (log.isWarnEnabled ())
          log.warn ("{}", obj1);
      }
      else
        if (level == INFO)
        {
          if (log.isInfoEnabled ())
            log.info ("{}", obj1);
        }
        else
          if (level == DEBUG)
          {
            if (log.isDebugEnabled ())
              log.debug ("{}", obj1);
          }
          else
          {
            if (log.isTraceEnabled ())
              log.trace ("{}", obj1);
          }
  }

  /**
   * Log a message
   * 
   * @param level
   *        One of DEBUG, INFO, WARN, ERROR, FATAL
   * @param obj1
   *        The object to log. This is converted to a string.
   * @param exception
   *        An exception to be logged
   */
  @Override
  public void log (final int level, final Object obj1, final Throwable exception)
  {
    if (level == FATAL || level == ERROR)
    {
      if (log.isErrorEnabled ())
      {
        if (obj1 != null)
          log.error ("{}", obj1, exception);
        else
          log.error (BLA, exception);
      }
    }
    else
      if (level == WARN)
      {
        if (log.isWarnEnabled ())
        {
          if (obj1 != null)
            log.warn ("{}", obj1, exception);
          else
            log.warn (BLA, exception);
        }
      }
      else
        if (level == INFO)
        {
          if (log.isInfoEnabled ())
          {
            if (obj1 != null)
              log.info ("{}", obj1, exception);
            else
              log.info (BLA, exception);
          }
        }
        else
          if (level == DEBUG)
          {
            if (log.isDebugEnabled ())
            {
              if (obj1 != null)
                log.debug ("{}", obj1, exception);
              else
                log.debug (BLA, exception);
            }
          }
          else
          {
            if (log.isTraceEnabled ())
            {
              if (obj1 != null)
                log.trace ("{}", obj1, exception);
              else
                log.trace (BLA, exception);
            }
          }
  }

  /**
   * Check if a logger is enabled to log at the specified level
   * 
   * @param level
   *        One of DEBUG, INFO, WARN, ERROR, FATAL
   */
  @Override
  public boolean check (final int level)
  {
    if (level == FATAL || level == ERROR)
      return log.isErrorEnabled ();
    if (level == WARN)
      return log.isWarnEnabled ();
    if (level == INFO)
      return log.isInfoEnabled ();
    if (level == DEBUG)
      return log.isDebugEnabled ();
    return log.isTraceEnabled ();
  }
}
