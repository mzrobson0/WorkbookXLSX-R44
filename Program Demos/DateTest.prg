*- DateTest.prg
*-
LOCAL loXL, lnWb, lnSh, lnStyle, lcFormat, lnFormat, lnRow

#INCLUDE "..\vfpxworkbookxlsx.h"

loXL = NEWOBJECT("VFPxWorkbookXLSX", "..\VFPxWorkbookXLSX.vcx")
lnWb = loXL.CreateWorkbook('DateTest.xlsx')
lnSh = loXL.AddSheet(lnWb, 'Sheet1')

loXL.SetColumnWidth(lnWb, lnSh, 1, 15)
loXL.SetColumnWidth(lnWb, lnSh, 2, 40)
loXL.SetColumnWidth(lnWb, lnSh, 3, 75)

loXL.SetCellValue(lnWb, lnSh, 1, 1, 'Value')
loXL.SetCellValue(lnWb, lnSh, 1, 2, 'Format')
loXL.SetCellValue(lnWb, lnSh, 1, 3, 'Notes')

lnRow = 2
loXL.SetCellValue(lnWb, lnSh, 2, 1, 44015.338889)

lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = '###0.0"kg";[Red]-###0.0"kg";"";General'
lnFormat = loXL.AddNumericFormat(lnWb, lcFormat)
loXL.AddStyleNumericFormat(lnWb, lnStyle, lnFormat)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, 'A custom numeric format. Works perfectly.')


lnRow = 3
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'dd mmm yy hh:mm'
lnFormat = loXL.AddCustomDateTimeFormat(lnWb, lcFormat)
loXL.AddStyleNumericFormat(lnWb, lnStyle, lnFormat)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, 'A custom datetime format. Works when used manually in excel, but not in code.')

lnRow = 4
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'dd mmm yy'
lnFormat = loXL.AddCustomDateTimeFormat(lnWb, lcFormat)
loXL.AddStyleNumericFormat(lnWb, lnStyle, lnFormat)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, 'A custom date format. Works when used manually in excel, but not in code.')

lnRow = 5
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'CELL_FORMAT_DATETIME_DDMMMYYYY_TT24'
set step on
loXL.AddStyleNumericFormat(lnWb, lnStyle, CELL_FORMAT_DATETIME_DDMMMYYYY_TT24)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, "Not sure what's happened to the TT24 part of this, or why MMM doesn't show Jul.")

lnRow = 6
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'CELL_FORMAT_DATETIME_MDYYHMM'
loXL.AddStyleNumericFormat(lnWb, lnStyle, CELL_FORMAT_DATETIME_MDYYHMM)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, "Correct, except that YY implies 20 not 2020.")

lnRow = 7
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'CELL_FORMAT_DATE_DMMMYY'
loXL.AddStyleNumericFormat(lnWb, lnStyle, CELL_FORMAT_DATE_DMMMYY)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, "I can live with this, though I don't like the dashes.")

lnRow = 8
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
lnStyle = loXL.CreateFormatStyle(lnWb)
lcFormat = 'CELL_FORMAT_DATETIME_MMMDDYYYY_TT24'
loXL.AddStyleNumericFormat(lnWb, lnStyle, CELL_FORMAT_DATETIME_MMMDDYYYY_TT24)
loXL.SetCellStyleRange(lnWb, lnSh, lnRow, 1, lnRow, 1, lnStyle)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, lcFormat)
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, "Well, that didn't work!")

lnRow = 9
lnStyle = loXL.CreateFormatStyle(lnWb)
loXL.SetCellValue(lnWb, lnSh, lnRow, 1, 44015.338889)
loXL.SetCellValue(lnWb, lnSh, lnRow, 2, "CELL_FORMAT_DATE_DMMMYYHMM")
loXL.SetCellValue(lnWb, lnSh, lnRow, 3, "This would do, but it doesn't exist.")

loXL.SaveWorkbook(lnWb)
