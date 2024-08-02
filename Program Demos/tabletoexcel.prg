#INCLUDE "..\VFPxWorkbookXLSX.h"
LOCAL loExcel, lcTable, lcExcel, lnTime
lcTable = GETFILE("DBF")
IF FILE(lcTable)
	lcExcel = FORCEEXT(lcTable, "xlsx")
	loExcel = NEWOBJECT("VFPxWorkbookXLSX", "e:\my work\foxpro\projects\workbookxlsx\VFPxWorkbookXLSX.vcx")
	lnTime  = SECONDS()
	loExcel.SaveTableToWorkbookEx(lcTable, lcExcel, .NULL., .T.)
	?SECONDS()-lnTime
ENDIF