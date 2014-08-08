;*******************************************************************
;Description: Session Affinity
;
;Purpose: Creates a Java Project and publish in cloud with staging target
;Environment and Overwrite previous deployment
;
;Date: 10 Jun 2014 , Modified on 12 June 2014
;Author: Ganesh
;Company: Brillio
;*********************************************************************

;********************************************
;Include Standard Library
;*******************************************
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiTreeView.au3>
#include <File.au3>
;******************************************

;***************************************************************
;Initialize AutoIT Key delay
;****************************************************************
AutoItSetOption ( "SendKeyDelay", 400)

;******************************************************************
;Reading test data from xls
;To do - move helper function
;******************************************************************
;Open xls
Local $sFilePath1 = @ScriptDir & "\" & "TestData.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)
Dim $oExcel1 = _ExcelBookNew(0)

;Local $sFilePath2 = @ScriptDir & "\" & "Result.xlsx"  ;This file should already exist in the mentioned path
;Local $oExcel1 = _ExcelBookOpen($sFilePath2,0,False)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

 ; Reading xls data into variables
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 10, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 10, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 10, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 10, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 10, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 10, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 14, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 14, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 10, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 10, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 10, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 10, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 10, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 10, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 10, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 10, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 10, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 10, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 10, 19)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 10, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 10, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 10, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 10, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 10, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 10, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 10, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 10, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 10, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 10, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 10, 31)
_ExcelBookClose($oExcel,0)
;*******************************************************************************
;*******************************************************************************






;to do - Pre validation steps

;Opening instance of Eclipse
Local $pro = ProcessExists("eclipse.exe")
If $pro > 0 Then
 Delete()
Else
 OpenEclipse()
 Delete()
EndIf
;OpenEclipse()
;Deleting the existing project
;Delete()


;Creating Java Project
CreateJavaProject()

;Creating JSP file and insert code
CreateJSPFile()

;CreateAzurePackage
CreateAzurePackage()

;Publish to Cloud
PublishToCloud()

;Wait for 10 min RDP screen
;Sleep(720000)

For $i = 7 to 1 Step - 1
   Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:msctls_progress32]")
   Local $syslk = ControlCommand($wnd, "", $wnd1,"IsVisible", "")
If $syslk = 1 Then
   ;Check RDP and Open excel
   CheckRDPConnection()
   Sleep(10000)
   ;Check for published key word in Azure activity log and update excel
   ValidateTextAndUpdateExcel()
   Sleep(5000)
   ExitLoop
Else
   Sleep(120000)
EndIf
   Next

;Check RDP and Open excel
;CheckRDPConnection()

;Check for published key word in Azure activity log and update excel
;ValidateTextAndUpdateExcel()

;to do - Post validation steps





;***************************************************************
;Helper Functions
;***************************************************************
Func wincheck($fun ,$ctrl)
 	  Local $act = WinActive($ctrl)
	  if $act = 0 Then
		 Local $lFile = FileOpen(@ScriptDir & "\" & "Error.log", 1)
 		 Local $wrt = _FileWriteLog(@ScriptDir & "\" & "Error.log", "Error Opening:" & $ctrl, 1)
		 MsgBox("","",$wrt)
		 FileClose($lFile)
		 MsgBox($MB_OK,"Error","Error status is recorded in Error.log")
	  EndIf
 EndFunc
;***************************************************************
;Function to Open instance of Eclipse
;***************************************************************
Func OpenEclipse()
Run($testCaseEclipseExePath)
WinWaitActive("Workspace Launcher")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseWorkSpacePath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create Java Project
;***************************************************************
Func CreateJavaProject()
Send("!fnd")
WinWaitActive("[Title:New Dynamic Web Project]")

Local $funame, $cntrlname
$cntrlname = "[Title:New Dynamic Web Project]"
$funame = "CreateJavaProject"
wincheck($funame,$cntrlname)

AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
;Send("{TAB 10}")
;Send("{Enter}")
Send("!f")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create JSP file and insert code
;***************************************************************
Func CreateJSPFile()
sleep(3000)
Send("{APPSKEY}")
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)

Send("^a")
Send("{Backspace}")
ClipPut($testCaseJspText)
Send("^v")
Send("^+s")

;create newsession.jsp
MouseClick("primary",105, 395, 1)
Send("{APPSKEY}")
Sleep(1000)
Send("n")
Send("{down 14}")
Send("{enter}")
Send("newsession.jsp")
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
#comments-start
if $JunoOrKep = "Juno" Then
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Else
   Local $temp = "Web - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
   EndIf
   #comments-end
Sleep(2000)
Send("^a")
Send("{Backspace}")
ClipPut($testcaseNewSessionJSPText)
Send("^v")
AutoItSetOption ( "SendKeyDelay", 400)
Send("^+s")
;MouseClick("primary", 74, 114, 1)

EndFunc
;******************************************************************

;***************************************************************
;Function to create Azure project
;***************************************************************

Func CreateAzurePackage()
;WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}}")
;Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
;Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
;Local $cnt = _GUICtrlTreeView_GetCount($hToolBar)
;WinActivate($hToolBar)
;;ControlClick($hWnd,"","left",1)
;_GUICtrlTreeView_Expand($hToolBar,0,False)
;Send("{RIGHT}")
;Send("{LEFT}")
;Send("{RIGHT}")
;Send("{DOWN}")
;Send("{UP}")
;Send("{ENTER}")

Send("{APPSKEY}")
Sleep(1000)
#comments-start
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
#comments-end
Send("e")
Send("{Left}")
Send("{UP}")
;Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")


; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "[Title:New Azure Deployment Project]"
$funame = "CreateAzurePackage"
wincheck($funame,$cntrlname)

AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")
;JDK configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckJdk,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","Check", "")
   EndIf
AutoItSetOption ( "SendKeyDelay", 100)
Send("{TAB}")
Send("+")
Send("{End}")
Send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("!N")

;Server Configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckLocalServer,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","Check", "")
   EndIf
Send("{TAB}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
Send($testCaseServerPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 2}")

 for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next
Send("!N")
Send("!N")
ControlCommand("New Azure Deployment Project","","[CLASSNN:Button16]","Check", "")
Send("!F")
EndFunc
;******************************************************************

;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud()
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse")
Send("{Up}")
Send("{APPSKEY}")
Sleep(1000)
#comments-start
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
#comments-end
Send("e")
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")

WinWaitActive("Publish Wizard")
Sleep(3000)
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd
WinActive("Publish Wizard")
Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:1]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseSubscription)

 Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:2]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseStorageAccount)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:3]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServiceName)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetOS)

Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:5]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetEnvironment)
#comments-start
Send("{TAB}")

for $count = $testCaseSubscription to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseStorageAccount to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseServiceName to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseTargetOS to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
for $count = $testCaseTargetEnvironment to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
#comments-end
Local $cmp = StringCompare($testCaseCheckOverwrite,"UnCheck")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
   Else
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseUserName)
Send("{TAB}")
Send($testCasePassword)
Send("{TAB 2}")
Send($testCasePassword)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB}")
ControlCommand("Publish Wizard","","[CLASSNN:Button5]","Check", "")
Send("{TAB}")
Send("{Enter}")
EndFunc
;***************************************************************************

;*****************************************************************
;Function to check the status of RDP and Open Excel
;****************************************************************
Func CheckRDPConnection()
Local $tempTime = _Date_Time_GetLocalTime()
Local $timeDateStamp = _Date_Time_SystemTimeToDateTimeStr($tempTime)
Local $RDPWindow = ControlCommand("Remote Desktop Connection","","[CLASSNN:Button1]","IsVisible", "")
;MsgBox("","",$RDPWindow,3)
#comments-start
Dim $oExcel = _ExcelBookNew(0)

If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "parameter is not a number")
    Exit
EndIf
#comments-end

_ExcelWriteCell($oExcel1, "Date And Time", 1, 1)
_ExcelWriteCell($oExcel1, "Scenario7" , 1, 4)
_ExcelWriteCell($oExcel1, $timeDateStamp , 2, 1)
_ExcelWriteCell($oExcel1, "RDPConnectionStatus", 1, 2)
_ExcelWriteCell($oExcel1, "Test Result" , 1, 3)

if $RDPWindow = 1 Then
_ExcelWriteCell($oExcel1, "Yes", 2, 2)
Send("{TAB 4}")
Send("{Enter}")
Else
_ExcelWriteCell($oExcel1, "No", 2, 2)
EndIf

;Local $flag = _ExcelBookSaveAs($oExcel, @ScriptDir & "\Result" & @ScriptName, "xls",0,1)
;If $flag <> 1 Then MsgBox($MB_SYSTEMMODAL, "Not Successful", "File Not Saved!", 5)
;_ExcelBookClose($oExcel, 1, 0)
EndFunc
;***************************************************************************************

;**************************************************************************
;Function to check publish key word in Azure activity log and update excel
;**************************************************************************
Func ValidateTextAndUpdateExcel()
   #comments-start
MouseClick("primary",565, 632, 1)
Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)
#comments-end
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink]")
 ControlClick($wnd,"",$wnd1,"left")

;Check in webpage and update excel
Send("{TAB}")
Send("{Enter}")
Sleep(8000)
Send("{F6}")
Send("^c")
Local $url = ClipGet();
Local $temp = $url & $testCaseProjectName
Local $oIE = _IECreate($temp,0,1,1,1)
_IELoadWait($oIE)
Local $readHTML = _IEBodyReadText($oIE)
Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)
Sleep(10000)
_IEQuit($oIE)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

if $iCmp = 1 Then
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")
_ExcelWriteCell($oExcel1, "Test Passed" , 2, 3)
Else
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Failed")
_ExcelWriteCell($oExcel1, "Test Failed" , 2, 3)
EndIf

Local $flag = _ExcelBookSaveAs($oExcel1, @ScriptDir & "\" & $testCaseName & "Result", "xls",0,1)
;Local $flag = _ExcelBookSave($oExcel1,0)
If Not @error Then MsgBox($MB_SYSTEMMODAL, "Success", "File was Saved!", 3)
_ExcelBookClose($oExcel1, 1, 0)
EndFunc
;*******************************************************************************

Func Delete()

Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
WinActivate($hToolBar)
MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}}")
for $i = 6 to 1 Step - 1
   Local $chk = _GUICtrlTreeView_GetCount($hToolBar)

if $chk = 0 Then
	ExitLoop
 Else
	MouseClick("primary",119, 490, 1)
   Send("^+{NUMPADDIV}}")
   Send("{RIGHT}")
   Send("{DOWN}")
   Send("{UP}")
   Send("{DELETE}")
   Send("{SPACE}")
   Send("{ENTER}")

   EndIf
Next
EndFunc
;*************************************