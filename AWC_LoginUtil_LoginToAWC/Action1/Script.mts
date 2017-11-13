'! @Name AWC_LoginUtil_LoginToAWC
'! @Details Action word to Login to AWC client
'! @InputParam1. sAutomationID : Automation ID  
'! @InputParam2. sLoginDetails : Application login details
'! @InputParam3. bSetGroupRole : Group and Role setting flag
'! @InputParam4. sNavigateOption : Navigate from home folder option
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 03-Feb-2016
'! @Version 1.0
'! @Example  LoadAndRunAction "AWC_LoginUtil\AWC_LoginUtil_LoginToAWC","AWC_LoginUtil_LoginToAWC",OneIteration,"TestUser2EngineeringOckAlEngineer",True,True,"Changes"

Option Explicit

'Declaring variables
Dim sAutomationID,bRelaunch,bSetGroupRole,sNavigateOption
Dim bReturn
Dim iCounter
Dim aLoginDetails
Dim sLoginDetails,sUserSettingsAutomationID
Dim objAWCSignIn,objAWCDefaultPage

GBL_CURRENT_EXECUTABLE_APP="AWC"

sAutomationID=Parameter("sAutomationID")
bRelaunch=Parameter("bRelaunch")
bSetGroupRole=Parameter("bSetGroupRole")
sNavigateOption=Parameter("sNavigateOption")

'Creating object of [ AWC Sign In ] page
Set objAWCSignIn = Fn_FSOUtil_XMLFileOperations("getobject","AWC_LoginUtil_OR", "wbpge_AWCSignIn","")
Set objAWCDefaultPage = Fn_FSOUtil_XMLFileOperations("getobject","AWC_LoginUtil_OR", "wbpge_AWCDefaultPage","")

'getting login credential information
If sAutomationID<>"" Then
	sUserSettingsAutomationID=sAutomationID
	sLoginDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sAutomationID)
End If

If Cbool(bRelaunch)Then
	LoadAndRunAction "AWC_LoginUtil\AWC_LoginUtil_KillProcess","AWC_LoginUtil_KillProcess",OneIteration,""
End If

aLoginDetails  =Split(sLoginDetails,"~",-1,1)
		
'Navigating to AWC application login page
LoadAndRunAction "AWC_LoginUtil\AWC_LoginUtil_NavigateToAWCApplication","AWC_LoginUtil_NavigateToAWCApplication",OneIteration

'Captures function execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","AWC_LoginUtil_LoginToAWC","Login","","")

If Fn_WEB_UI_WebObject_Operations("AWC_LoginUtil_LoginToAWC","Exist",objAWCSignIn.WebEdit("wbedt_UserName"),"","","") =True  Then			
	'Entering User Name
	If Fn_Web_UI_WebEdit_Operations("AWC_LoginUtil_LoginToAWC", "Set", objAWCSignIn.WebEdit("wbedt_UserName"), "", aLoginDetails(0))=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter User name [ " & Cstr(aLoginDetails(0)) & " ] while login to AWC application","","","","","")
		Call Fn_ExitTest()
	End If
	'Entering Password
	If Fn_Web_UI_WebEdit_Operations("AWC_LoginUtil_LoginToAWC", "SetSecure", objAWCSignIn.WebEdit("wbedt_Password"), "", aLoginDetails(1))=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to enter Password [ " & Cstr(aLoginDetails(1)) & " ] while login to AWC application","","","","","")
		Call Fn_ExitTest()
	End If
	'Clicking on [ Login ] button
	If Fn_Web_UI_WebButton_Operations("AWC_LoginUtil_LoginToAWC", "Click", objAWCSignIn.WebButton("wbbtn_SignIn"), "","","","")=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ SignIn ] button while login to AWC application","","","","","")
		Call Fn_ExitTest()
	End If
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to AWC application as login page not exist","","","","","")
	Call Fn_ExitTest()
End IF

bReturn=False
For iCounter = 1 To 30
	wait 0,500
	'Checking existance of Login Page
	If Fn_WEB_UI_WebObject_Operations("AWC_LoginUtil_LoginToAWC", "Exist",objAWCSignIn.WebEdit("wbedt_UserName"),"","","") Then
		bReturn=True
		Exit For
	End If
Next
If bReturn=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to login to AWC application with user [ " & Cstr(sLoginDetails) & " ]","","","","","")
	Call Fn_ExitTest()
End If

Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully login to AWC application with user [ " & Cstr(sLoginDetails) & " ]","","","","","")
'Captures function execution End time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","AWC_LoginUtil_LoginToAWC","Login","","")

'Setting group and role
If Cbool(bSetGroupRole) Then
	LoadAndRunAction "AWC_Common\AWC_UserSettingOperations","AWC_UserSettingOperations",OneIteration,"SetGroupRole",sUserSettingsAutomationID,"",""
End If
Call Fn_AWC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Navigating from home page
If sNavigateOption<>"" Then
	LoadAndRunAction "AWC_Common\AWC_Common_MainPageNavigationOperations","AWC_Common_MainPageNavigationOperations",OneIteration,"Navigate",sNavigateOption
End If
Call Fn_AWC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to AWC login operation fail due to Err.Number [ " & Cstr(Err.Number) & " ] and Err.Description [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

Set objAWCSignIn=Nothing
Set objAWCDefaultPage = Nothing
	
Function Fn_ExitTest()
	'Releasing all required objects
	Set objAWCSignIn=Nothing
	Set objAWCDefaultPage = Nothing	
	ExitTest
End Function
