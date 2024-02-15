'#######################################################################################
'//###    TESTCASE NAME   :  TeamCenterLogin
'//###
'//###    DESCRIPTION     :   		Launching Team Center
'//###
'//###    HISTORY        		 :   		AUTHOR              		DATE        				VERSION
'//###
'//###    CREATED BY      :   			Aniket Gadage			  15-Jan-2024		   	UFT-ONE
'//###
'//###    REVIWED BY      :									 15-Jan-2024
'//###
'//###
'//###    Run on Tc Build  :
'//######################################################################################
'//######################################################################################

Option Explicit
'*********************************************************************************
'Variable Declaration
'*********************************************************************************
Dim bReturn

'*********************************************************************************
'Set the UFT Test Requirement
'*********************************************************************************
Call fn_setup_TestCaseinit

'*********************************************************************************
'Test Case Start Time
'*********************************************************************************
Call fn_UpdateLogFile(Cstr(Date) + " ***************************** QTP Action1 - Start **********************************"+Cstr(Time) )

'*********************************************************************************
'Launching TeamCenter Application
'*********************************************************************************
Call fn_Launch_TC()

'*********************************************************************************
'Login TeamCenter Application
'*********************************************************************************
bReturn=fn_Login_TeamCenter(DataTable.Value("UserName","TeamCenterLogin"),DataTable.Value("Password","TeamCenterLogin"))
 @@ hightlight id_;_455550900_;_script infofile_;_ZIP::ssf1.xml_;_
If bReturn Then
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Pass | Successfully Login To  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
Else
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Fail | UnSuccessfully Login To  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
End If

'*********************************************************************************
'Folder Creation In TeamCenter Application
'*********************************************************************************
bReturn=fn_CreateFolder(DataTable.Value("Folder_Name","TeamCenterLogin"),DataTable.Value("Folder_Desc","TeamCenterLogin")) @@ hightlight id_;_455550900_;_script infofile_;_ZIP::ssf1.xml_;_

If bReturn Then
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Pass | Successfull Folder Creation  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
Else
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Fail | UnSuccessfull Folder Creation  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
End If
'*********************************************************************************
'Logout TeamCenter Application
'*********************************************************************************
bReturn=fn_Logout_TeamCenter()

If bReturn Then
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Pass | Successfully Logout From  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
Else
	Call Fn_UpdateLogFile("[" + Cstr(time) + "] - Action - Fail | UnSuccessfully Logout From  TeamCenter for User [" + DataTable.Value("UserName","TeamCenterLogin") + "]")
End If

'*********************************************************************************
'Test Case End Time
'*********************************************************************************
Call fn_UpdateLogFile(Cstr(Date) + " ***************************** QTP Action1 - End **********************************"+Cstr(Time) )
