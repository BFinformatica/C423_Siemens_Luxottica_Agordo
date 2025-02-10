CreateReport()

Sub CreateReport()

	Dim objrptdll
	
    	Set objrptdll = CreateObject("WindasReport.CReport")

	objrptdll.DbType = "SQL" 
	objrptdll.DbDatabase = "bfdata"
	objrptdll.DbUser = "bf"
	objrptdll.DbPassword = "Bfinfo9876"
	objrptdll.DbServer = "localhost\SQLEXPRESS"
	objrptdll.DbVersion = ""
	call objrptdll.SetStartDateFromString("23/10/2024 00:00")
	call objrptdll.SetEndDateFromString("23/10/2024 00:00")
	objrptdll.Stations = "CEMS1"
	objrptdll.Param = "CO,O2,SO2,PFUMI,TFUMI,QFUMI,O2U,THC,TLINEA,NOX,PLV,H2O"
	objrptdll.Table = "WDS_ELAB"
	objrptdll.ModelFileName = "c:\Windas\bfinfo\Wtf_ModelloOrario.xls"
	
	'objrptdll.DbType = "SQL" 
	'objrptdll.DbDatabase = "bfdata"
	'objrptdll.DbUser = "bf"
	'objrptdll.DbPassword = "Bfinfo9876"
	'objrptdll.DbServer = "localhost\WINCC"
	'objrptdll.DbVersion = ""
	'call objrptdll.SetStartDateFromString("10/06/2015")
	'call objrptdll.SetEndDateFromString("10/06/2015")
	'objrptdll.Stations = "WINCC"
	'objrptdll.Param = "CO_L1,THC_L1,T_Fumi_L1"
	'objrptdll.Table = "wds_elab"
	'objrptdll.ModelFileName = "C:\Documents and Settings\Administrator\Desktop\Unicalce\Report VBS Italiano\Wtf_ModelloOrario.xls"
	
	Call objrptdll.CreateReport
	
	Set objrptdll = Nothing

End sub