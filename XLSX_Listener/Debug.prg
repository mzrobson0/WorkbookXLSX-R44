LOCAL lcPath, lcTable, loListener, lcAcctCode, ldBegDate, ldEndDate, lcBusinessName, lcBusinessAddr, lcReportName


*******************************************************************************************
*-*	Open listener and set properties

loListener = NEWOBJECT("xlsx_listener", "..\vfpxworkbookxlsx.vcx")
loListener.CodePage = 1252             && Default value
loListener.DebugLoadReport()