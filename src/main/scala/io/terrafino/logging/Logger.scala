package io.terrafino.logging

import org.apache.log4j.{ Logger => ApacheLogger }

trait Logger {
  val loggerName = this.getClass.getName
  lazy val logger = ApacheLogger.getLogger(loggerName)
}