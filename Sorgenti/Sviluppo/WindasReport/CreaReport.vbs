CreateReport()

Sub CreateReport()

	Dim objrptdll
	
    Set objrptdll = CreateObject("WindasReport.CReport")

	objrptdll.DbType = "@DBTYPE" 
	objrptdll.DbDatabase = "@DATABASE"
	objrptdll.DbUser = "@USER"
	objrptdll.DbPassword = "@PASSWORD"
	objrptdll.DbServer = "@SERVER"
	objrptdll.DbVersion = "@AUTHENTICATION"
	call objrptdll.SetStartDateFromString("@STARTDATE")
	call objrptdll.SetEndDateFromString("@ENDDATE")
	objrptdll.Stations = "@STATIONS"
	objrptdll.Param = "@MEASURES"
	objrptdll.Table = "@DATASOURCE"
	objrptdll.ModelFileName = "@FILENAME"
	
	Call objrptdll.CreateReport
	
	Set objrptdll = Nothing

End sub