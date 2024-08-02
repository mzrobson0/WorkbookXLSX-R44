*PUBLIC loExcel   && to keep it from being destroyed and closing the cursors
LOCAL lcFile, loText
lcFile = GETFILE("xlsx", "Workbook", "Load", 0, "Select Workbook to load into Class")
IF !EMPTY(lcFile)
	loExcel = NEWOBJECT("VFPxWorkbookXLSX", "E:\My Work\FoxPro\Projects\WorkbookXLSX\VFPxWorkbookXLSX.vcx")
*	lnBegSec = SECONDS()
*	? "Processing file: " + JUSTFNAME(lcFile)
	lnWB = loExcel.OpenXlsxWorkbook(lcFile)
*	loExcel.SaveWorkbook(lnWb)
*	lnEndSec = SECONDS()
*	IF (lnEndSec - lnBegSec) > 120
*		? "Time to load: " + TRANSFORM((lnEndSec - lnBegSec)/60) + " minutes"
*	ELSE
*		? "Time to load: " + TRANSFORM(lnEndSec - lnBegSec) + " seconds"
*	ENDIF
*
*	lnBegSec = SECONDS()
	loExcel.SaveWorkbookAs(lnWB, ADDBS(JUSTPATH(lcFile))+JUSTSTEM(lcFile)+"-2.xlsx")
*	lnEndSec = SECONDS()
*	IF (lnEndSec - lnBegSec) > 120
*		? "Time to save: " + TRANSFORM((lnEndSec - lnBegSec)/60) + " minutes"
*	ELSE
*		? "Time to save: " + TRANSFORM(lnEndSec - lnBegSec) + " seconds"
*	ENDIF
ENDIF
