#INCLUDE "..\vfpxworkbookxlsx.h"
*PUBLIC goExcel   && to keep it from being destroyed and closing the cursors
LOCAL lcFile, lnWb, loSheets, lnSh, lnRow, lnCol, lcOutPath, lnSec, lcValue
lcFile = GETFILE("xlsx", "Workbook", "Load", 0, "Select Workbook to load into Class")
IF !EMPTY(lcFile)
	lcOutPath = ADDBS(JUSTPATH(lcFile))
	goExcel   = NEWOBJECT("VFPxWorkbookXLSX", "..\VFPxWorkbookXLSX.vcx", "", 1250)   && Set codepage 1250
*	goExcel.Debug = .T.

	lnSec = SECONDS()
*	lnWb = goExcel.OpenXlsxWorkbookSheet(lcFile, 2)
	lnWb = goExcel.OpenXlsxWorkbook(lcFile)
	? "Workbook Open: " + TRANSFORM(SECONDS() - lnSec)

*	FOR lnRow=3 TO 6
*		FOR lnCol=1 TO 6
*			goExcel.SetCellValue(lnWB, 1, lnRow, lnCol, lnRow*lnCol)
*		ENDFOR
*	ENDFOR

*	goExcel.SetCellValue(lnWb, 1, 5, 3, DATE())

*	lnRows=goExcel.GetLastRowNumber(lnWb,1)
*	goExcel.SetCellAlignmentRange(lnWb, 1, 1, 9, lnRows,10, CELL_HORIZ_ALIGN_RIGHT, CELL_VERT_ALIGN_CENTER)
*	goExcel.SetCellnumberformatRange(lnWb, 1, 1, 9, lnRows, 10, CELL_FORMAT_CURRENCY_RED_PAREN)
*	goExcel.SetColumnwidth(lnWb, 1, 1, 12)
*	goExcel.SetColumnwidth(lnWb, 1, 4, 30)
*	goExcel.SetColumnwidth(lnWb, 1, 5, 15)
*	goExcel.SetColumnwidth(lnWb, 1, 7, 12)
*	goExcel.SetColumnwidth(lnWb, 1, 8, 12)
*	goExcel.SetColumnwidth(lnWb, 1, 9, 12)

*	?goExcel.GetCellValue(lnWB, 1, 57, 2)
*	?goExcel.GetCellValue(lnWB, 1, 58, 2)
*	SET DEBUGOUT TO lcOutPath + "DebugExcelRead.txt"
*	loSheets = goExcel.GetWorkbookSheets(lnWb)
*	FOR lnSh=1 TO loSheets.Count
*		DEBUGOUT "Sheet Index: ", loSheets.List[lnSh, 1]      && Displays sheet index (which may not be the same as lnSh)
*		DEBUGOUT "Sheet Name:  ", loSheets.List[lnSh, 2]      && Displays sheet name
*		FOR lnRow=1 TO goExcel.GetLastRowNumber(lnWb, loSheets.List[lnSh, 1])
*			DEBUGOUT "Row: ", lnRow
*			FOR lnCol=1 TO goExcel.GetMaxColumnNumber(lnWb, loSheets.List[lnSh, 1])
*				IF goExcel.IsCellFormula(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
*					lcValue = goExcel.GetCellFormula(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
*					DEBUGOUT "Column: ", lnCol, "  Formula: ", goExcel.GetCellFormula(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
*				ENDIF
*				DEBUGOUT "Column: ", lnCol, "  Value: ", goExcel.GetCellValue(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
*				lcValue = goExcel.GetCellValue(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
*			ENDFOR
*		ENDFOR
*	ENDFOR
*	SET DEBUGOUT TO
*	lcImageFile = ADDBS(SYS(5) + SYS(2003)) + "vfpxbanner.gif"                         && Must provide full path
*	goExcel.AddImage(lnWB, 1, lcImageFile, IMAGE_ANCHOR_TYPE_ONE,, 16, 0, 3, 0)

*	goExcel.InsertRow(lnWB, 1,  8, INSERT_BEFORE, 3)
*	goExcel.InsertRow(lnWB, 1, 17, INSERT_BEFORE)
*	goExcel.InsertColumn(lnWB, 1,  2, INSERT_RIGHT)
*	goExcel.DeleteRow(lnWB, 1, 16)
*	goExcel.DeleteColumn(lnWB, 1, 5)
*	goExcel.InsertCell(lnWB, 1,  10, 2, INSERT_LEFT)

*	for lnRow=1 to 3
*		goExcel.SetRowHeight(lnWB, 1, lnRow, 30)
*	endfor
	goExcel.SetRowHeightRange(lnWB, 1, 1, 3, 30)

	lnSec = SECONDS()
	goExcel.SaveWorkbookAs(lnWb, lcOutPath + JUSTSTEM(lcFile) + "_2.xlsx")
	? "Workbook Save: " + TRANSFORM(SECONDS() - lnSec)
ENDIF
