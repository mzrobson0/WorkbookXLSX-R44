LOCAL lcPath, lcTable, loListener, lcAcctCode, ldBegDate, ldEndDate, lcBusinessName, lcBusinessAddr, lcReportName

*******************************************************************************************
*-*	Open tables

IF !USED('journal')
	lcPath  = ADDBS(JUSTPATH(SYS(16, 2)))
	lcTable = lcPath + "Test Data\journal.dbf"
	USE (lcTable) IN 0 ALIAS journal SHARED

	lcTable = lcPath + "Test Data\accounts.dbf"
	USE (lcTable) IN 0 ALIAS accounts SHARED

	lcTable = lcPath + "Test Data\svcordrhdr.dbf"
	USE (lcTable) IN 0 ALIAS svcordrhdr SHARED

	lcTable = lcPath + "Test Data\svcordritems.dbf"
	USE (lcTable) IN 0 ALIAS svcordritems SHARED

	lcTable = lcPath + "Test Data\svctypes.dbf"
	USE (lcTable) IN 0 ALIAS svctypes SHARED

	lcTable = lcPath + "Test Data\statuses.dbf"
	USE (lcTable) IN 0 ALIAS statuses SHARED
ENDIF

*******************************************************************************************
*-*	Set report variables

ldBegDate  = DATE(2000, 1, 1)
ldEndDate  = DATE(2022, 12, 31)
lcAcctCode = "6"

lcBusinessName = "Test Name"
lcBusinessAddr = "123 Main St" + CHR(13)+CHR(10) + "Somewhere, GA  30817"
lcReportName   = "Account Totals - All Years"

*******************************************************************************************
*-*	Open listener and set properties

loListener = NEWOBJECT("xlsx_listener", "..\vfpxworkbookxlsx.vcx")
loListener.CodePage         = 1252             && Default value
loListener.SheetName        = "Test Output"    && Will be changed to use Directives in the group header band
loListener.FreezePanes      = .T.              && Default value is .T.
loListener.ShowGridLines    = .T.              && Default value
loListener.DefRowHeight     = 1666.6670        && Default value and maximum value; will be set back to default if equal to 0.0000
loListener.IgnoreCellErrors = .T.              && Default value is .T.; sets the cell properties to ignore numeric value formatted as text

loListener.PageMarginUnitOfMeas = "Inches"     && Default value; allowed values: Inches, Millimeters, Centimeters (case insensitive)
loListener.PageMarginsLeft   = 0.50            && Default value; in inches
loListener.PageMarginsRight  = 0.50            && Default value; in inches
loListener.PageMarginsTop    = 0.50            && Default value; in inches
loListener.PageMarginsBottom = 0.50            && Default value; in inches
loListener.PageMarginsHeader = 0.30            && Default value; in inches
loListener.PageMarginsFooter = 0.30            && Default value; in inches

loListener.DebugMode = .F.                     && Default value.F.; if set to .T., then it will write out dbf files for the report object content in order to debug the report export to Excel
                                               && Note, the database table will contain the report actual contents in the field cellvalue; if this is sensitive information please set the value(s)
                                               && to some other non-sensitive value before sending the debug table and report FRX files to me for debugging purposes

*******************************************************************************************
*-*	Test for Simple Reports
IF .T.
	SELECT jrn.accountid, act.code, act.aname, act.descriptn, SUM(jrn.amount) AS amount ;
		FROM journal AS jrn ;
		LEFT JOIN accounts AS act ON act.id = jrn.accountid ;
		WHERE act.code = lcAcctCode ;
			AND jrn.transdate >= ldBegDate ;
			AND jrn.transdate <= ldEndDate ;
		ORDER BY act.code ;
		GROUP BY 1, 2, 3, 4 ;
		INTO CURSOR c_accttotals

	SELECT c_accttotals
	loListener.XlsxFileName  = "SimpleReportTitlePgHdr.xlsx"
	REPORT FORM SimpleReportTitlePgHdr.frx OBJECT loListener
*	REPORT FORM SimpleReportTitlePgHdr.frx PREVIEW NOCONSOLE
ENDIF

*******************************************************************************************
*-*	Test for Single Group Bands
IF .T.
	SELECT ord.id, ord.memberid, ord.orderdate, ord.mbrname, sta.sname, ord.statusid ;
		FROM svcordrhdr AS ord ;
		LEFT JOIN statuses AS sta ON sta.id = ord.statusid ;
		ORDER BY ord.memberid ;
		INTO CURSOR c_ordershdr

	SELECT c_ordershdr
	loListener.XlsxFileName = "SingleGroupPgBrk.xlsx"
	REPORT FORM SingleGroupPgBrk.frx OBJECT loListener
*	REPORT FORM SingleGroupPgBrk.frx PREVIEW NOCONSOLE

	loListener.XlsxFileName = "SingleGroupPgBrkTitle.xlsx"
	REPORT FORM SingleGroupPgBrkTitle.frx OBJECT loListener
*	REPORT FORM SingleGroupPgBrkTitle.frx PREVIEW NOCONSOLE

	loListener.XlsxFileName = "SingleGroupNoPgBrk.xlsx"
	REPORT FORM SingleGroupNoPgBrk.frx OBJECT loListener
*	REPORT FORM SingleGroupNoPgBrk.frx PREVIEW NOCONSOLE
ENDIF

*******************************************************************************************
*-*	Test for Multi-line detail band Reports
IF .T.
	SELECT jrn.accountid, act.code, act.aname, act.descriptn, SUM(jrn.amount) AS amount ;
		FROM journal AS jrn ;
		LEFT JOIN accounts AS act ON act.id = jrn.accountid ;
		WHERE act.code = lcAcctCode ;
			AND jrn.transdate >= ldBegDate ;
			AND jrn.transdate <= ldEndDate ;
		ORDER BY act.code ;
		GROUP BY 1, 2, 3, 4 ;
		INTO CURSOR c_accttotals

	SELECT c_accttotals
	loListener.XlsxFileName  = "DtlMultiLine.xlsx"
	REPORT FORM DtlMultiLine.frx OBJECT loListener
*	REPORT FORM DtlMultiLine.frx PREVIEW NOCONSOLE
ENDIF

*******************************************************************************************
*-*	Test for Detail Header and Footer bands
IF .T.
	SELECT jrn.accountid, act.code, act.aname, act.descriptn, SUM(jrn.amount) AS amount ;
		FROM journal AS jrn ;
		LEFT JOIN accounts AS act ON act.id = jrn.accountid ;
		WHERE act.code = lcAcctCode ;
			AND jrn.transdate >= ldBegDate ;
			AND jrn.transdate <= ldEndDate ;
		ORDER BY act.code ;
		GROUP BY 1, 2, 3, 4 ;
		INTO CURSOR c_accttotals

	SELECT c_accttotals
	loListener.XlsxFileName  = "DetailHdrFtr.xlsx"
	REPORT FORM DetailHdrFtr.frx OBJECT loListener
*	REPORT FORM DetailHdrFtr.frx PREVIEW NOCONSOLE
ENDIF

*******************************************************************************************
*-*	Test for Multi-Detail Bands
IF .T.
	SELECT ord.id, ord.memberid, ord.orderdate, ord.mbrname, sta.sname, ord.statusid ;
		FROM svcordrhdr AS ord ;
		LEFT JOIN statuses AS sta ON sta.id = ord.statusid ;
		WHERE INLIST(ord.id, "00005", "00006", "0000E") ;
		INTO CURSOR c_ordershdr

	SELECT itm.id, itm.svcorderid, itm.descriptn, itm.amount, itm.rate, itm.discount, itm.scheddate, itm.datescope, typ.type, typ.sname AS typename, sta.sname AS statusname, ;
		   0000 AS hours, 0000 AS quantity, 00000000.00 AS linetotal, SPACE(30) AS infodate ;
		FROM svcordritems AS itm ;
		LEFT JOIN svctypes AS typ ON typ.id = itm.svctypeid ;
		LEFT JOIN statuses AS sta ON sta.id = itm.statusid ;
		INTO CURSOR c_printitems READWRITE

	SELECT c_printitems
	INDEX ON svcorderid TAG svcorderid
	SCAN
		DO CASE
			CASE c_printitems.type = "T"
				lnQty   = 0
				lnHours = c_printitems.amount
				lnTotal = c_printitems.amount * c_printitems.rate * (100 - c_printitems.discount) / 100

			CASE c_printitems.type = "Q"
				lnQty   = c_printitems.amount
				lnHours = 0
				lnTotal = c_printitems.amount * c_printitems.rate * (100 - c_printitems.discount) / 100
		ENDCASE
		DO CASE
			CASE c_printitems.datescope = 1
				lcSchDate = DTOC(c_printitems.scheddate)

			CASE c_printitems.datescope = 2
				lcSchDate = "Week of " + DTOC(c_printitems.scheddate)

			CASE c_printitems.datescope = 3
				lcSchDate = "Month of " + CMONTH(c_printitems.scheddate) + " " + TRANSFORM(YEAR(c_printitems.scheddate))
		ENDCASE
		REPLACE c_printitems.quantity  WITH lnQty, ;
				c_printitems.hours     WITH lnHours, ;
				c_printitems.linetotal WITH lnTotal, ;
				c_printitems.infodate  WITH lcSchDate IN c_printitems
		IF ISNULL(c_printitems.statusname) .OR. EMPTY(c_printitems.statusname)
			REPLACE c_printitems.statusname WITH "Not Found" IN c_printitems
		ENDIF
	ENDSCAN

	SELECT c_ordershdr
	SET RELATION TO id INTO c_printitems
	loListener.XlsxFileName  = "MultiDetalBands.xlsx"
	REPORT FORM MultiDetalBands.frx OBJECT loListener
*	REPORT FORM MultiDetalBands.frx PREVIEW NOCONSOLE
	SET RELATION TO 
ENDIF

USE IN SELECT('c_ordershdr')
USE IN SELECT('c_printitems')
USE IN SELECT('statuses')
USE IN SELECT('svctypes')
USE IN SELECT('svcordrhdr')
USE IN SELECT('svcordritems')
USE IN SELECT('c_accttotals')
USE IN SELECT('journal')
USE IN SELECT('accounts')
