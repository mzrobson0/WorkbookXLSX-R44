*PUBLIC loExcel   && to keep it from being destroyed and closing the cursors
LOCAL lnTime, lnWb, lnSh, lnRow, lnCol

IF !PEMSTATUS(_SCREEN, "System", 5) .OR. VARTYPE(_SCREEN.System) <> "O"
	DO ..\System.app
ENDIF

loExcel = NEWOBJECT("VFPxWorkbookXLSX", "..\VFPxWorkbookXLSX.vcx")

loExcel.Demo()

*loExcel.DeleteAllWorkbooks()

*loExcel.SaveTableToWorkbook(HOME(2) + "Data\customer", "DemoCustomer.xlsx", .T., .T.)
*loExcel.SaveTableToWorkbook(HOME(2) + "Data\employee", "DemoEmployee.xlsx", .T., .T.)
*loExcel.SaveTableToWorkbook(HOME(2) + "Data\products", "DemoProducts.xlsx", .T., .T.)
*loExcel.DeleteAllWorkbooks()

*SET DEBUGOUT TO "TimeTest.txt"
*lnWb = loExcel.CreateWorkbook("TimeTest1.xlsx")
*IF lnWb > 0
*	DEBUGOUT "Writing First Time Test Workbook"
* 	lnSh = loExcel.AddSheet(lnWb, "Test Sheet 1")
* 	lnTime = SECONDS()
*	FOR lnRow=1 TO 10000
*		FOR lnCol=1 TO 5
*			loExcel.SetCellValue(lnWb, lnSh, lnRow, lnCol, SYS(2015))
*		ENDFOR
*		FOR lnCol=6 TO 10
*			loExcel.SetCellValue(lnWb, lnSh, lnRow, lnCol, lnRow*lnCol)
*		ENDFOR
*	ENDFOR
*	DEBUGOUT "SetCellValue:", SECONDS()-lnTime
* 	loExcel.SaveWorkbook(lnWb)
*	DEBUGOUT "SaveWorkbook:", SECONDS()-lnTime
*ENDIF
*SET DEBUGOUT TO