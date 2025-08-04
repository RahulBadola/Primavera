@echo off

@echo Running Database Configuration...

SETLOCAL

set INTG_HOME=.

SET PRIMAVERA_OPTS=-Dprimavera.bootstrap.home="%INTG_HOME%"

"%JAVA_HOME%\bin\java" -classpath "lib/intgserver.jar;lib/ojdbc6.jar;lib/sqljdbc.jar;lib/log4j.jar;lib/mail.jar;lib/spring.jar;lib/commons-logging.jar;" -Dadmin.showIntegration="y" -Dadmin.workingDir="%INTG_HOME%" -Dadmin.logDir="%INTG_HOME%/PrimaveraLogs" -Dadmin.console="y" -Dadmin.mode=install -Dadmin.precompile="n" %PRIMAVERA_OPTS% com.primavera.admintool.AdminApp

echo Done.
ENDLOCAL

