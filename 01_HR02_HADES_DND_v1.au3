Global $oMyError = ObjEvent("AutoIt.Error","COMErrHandler")
#AutoIt3Wrapper_UseX64=Y
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <String.au3>

Const $ERROR_SAP_NOT_YET_AVAL = -2147483638
Const $SAP_LOCAL_LANDSCAPE_PATH = @AppDataDir & "\SAP\Common\SAPUILandscape.xml"
Const $SYS_ADMINS = "tomas.ac@volvo.com;tomas.chudik@volvo.com"
Const $SAP_FNOTFOUND_SHELL = 1  ; Work item ID not found in the shell
Const $SAP_FNOTFOUND_REPORT = 2 ; Work item ID found in the shell but not in the report
Const $SAP_FCORUPTED = 3        ; Work item ID found in the shell, found in the report but corrupted
Const $SP_SITE_URL = "https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it"
Const $SP_FOLDER_NAME = "HR02_HADES_ARCHIVE_QA"
Global $oError
Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
Local $oMAIL = ObjCreate("CDO.Message")
Local $oSYSINFO = ObjCreate("ADSystemInfo")
Local $oFSO = ObjCreate("Scripting.FileSystemObject")
Local $oUSER = ObjGet("LDAP://" & $oSYSINFO.UserName)
Local $dictReport = ObjCreate("Scripting.Dictionary")
Local $dictFilesMovedFromShare = ObjCreate("Scripting.Dictionary")
Local $dictFilesUploaded = ObjCreate("Scripting.Dictionary")
Local $dictFilesFailedToUpload = ObjCreate("Scripting.Dictionary")
Local $nSAPLogonPID = Null 							; SAPLogon processId
Local $nFilesMoved = 0								; Number of files uploaded to the SAP
Local $nSessionsCount = 0
Local $nDuplicateFiles = 0							; Files uploaded to the SAP more than once due to delay in explorer
Local $sSAPSystemDescription = ""   				; Description string using to open connection to a SAP system
Local $sSAPSystemName = "FQ2"       				; Command line argument. For now variable initialized manually
Local $sSourceDirPath = "C:\!AUTO\HR02_HADES_DND_SOURCE"	; Command line argument. For now variable initialized manually
Local $sProcesseDirPath = "C:\!AUTO\HR02_HADES_DND_PROCESSED" ; Command line argument. For now variable initialized manually
Local $sNetworkSharePath = "\\siljun003\SCANNER\HR02_HADES_DND" ; Command line argument. For now variable initialized manually
Local $sOawdFolder = "HR02 Invoice Handler"				; Command line argument. For now variable initialized manually
Local $sFileMoved
Local $oFOLDER
Local $oGUI = Null		   							; SAPGUI object
Local $oSAP = Null		   							; SAP scripting engine
Local $oCON = Null         							; SAP connection object - upload files
Local $oSES = Null         							; SAP session object
Local $oSES2 = Null									; SAP session object
Local $oTBAR = Null									; SAP transaction window/field
Local $oWND0 = Null									; SAP wnd[0] - The one where we start after connection is established
Local $oWND1 = Null									; SAP wnd[1] - The one with drag and drop area
Local $hwndDnD										; SAP drag and drop window handle
Local $hwndSAPSystem								; SAP window where we execute transactions
Local $hwndExp										; Windows explorer window handle. Our source directory
Local $aFilesToRename								; Array holding files that should be renamed after moving from network share subfolder
Local $nPIDexp										; PID of the win explorer process started by the script
Local $aDnDSize			            				; Array containing dimensions and coordinates of the SAP Drag'n'Drop window
Local $aExpSize										; Array containing dimensions and coordinates of the Windows Explorer window
Local $aWindows										; Array holding window handles
Local $aStringParts									; Array holding substrings returned by StringSplit()
Local $hTimer = TimerInit()

;************************ Main *******************************
;*************************************************************
If $CmdLine[0] = 0 Then
	Exit(301) ; No cli params passed
EndIf

$sSAPSystemName = $CmdLine[1]
If Not StringRight($sNetworkSharePath,"\") Then $sNetworkSharePath = $sNetworkSharePath & "\"
If Not StringRight($sProcesseDirPath,"\") Then $sProcesseDirPath= $sProcesseDirPath & "\"
If Not StringRight($sSourceDirPath,"\") Then $sSourceDirPath = $sSourceDirPath & "\"
If Not $oFSO.FolderExists($sProcesseDirPath) Then $oFSO.CreateFolder($sProcesseDirPath)
If Not $oFSO.FolderExists($sSourceDirPath) Then $oFSO.CreateFolder($sSourceDirPath)
If $oFSO.GetFolder($sSourceDirPath).Files.Count > 0 Then
	MessageToAdmin($oUSER,$oMAIL,"I;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ";" & @UserName & ";" & @ComputerName & ";" & $sSAPSystemName,"Residual files detected in the " & $sSourceDirPath & "  Exit code: 11",$SYS_ADMINS)
	Exit(201)
ElseIf $oFSO.GetFolder($sProcesseDirPath).Files.Count > 0 Then
	MessageToAdmin($oUSER,$oMAIL,"I;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ";" & @UserName & ";" & @ComputerName & ";" & $sSAPSystemName,"Residual files detected in the " & $sProcesseDirPath & "  Exit code: 11",$SYS_ADMINS)
	Exit(202)
EndIf
; Check if the local directory is empty. DO NOT continue if there are residual files !!!
; This can happen if the previous run failed. Needs attention !
; Network share "\\siljun003\SCANNER\HR02_HADES_DND"
$oFOLDER = $oFSO.GetFolder($sNetworkSharePath)
For $folder in $oFOLDER.SubFolders
	If $oFSO.GetFolder($sNetworkSharePath & $folder.Name).Files.Count > 0 Then
		For $file in $oFSO.GetFolder($sNetworkSharePath & $folder.Name).Files
			$sFileMoved = $file.Name
			If StringRegExp($file.Name,"(.*?)\.(PDF|pdf|Pdf)$") Then ; move valid pdf file and rename it accordingly to the network share subfolder
				$oFSO.MoveFile($sNetworkSharePath & $folder.Name & "\" & $file.Name,$sSourceDirPath & $folder.Name & "_" & $file.Name) ; Prepend prefix to each file
				If $oFSO.FileExists($sSourceDirPath & $folder.Name & "_" & $sFileMoved) Then
					$dictFilesMovedFromShare.Add($folder.Name & "_" & $sFileMoved, $sFileMoved)
				EndIf
			EndIf
		Next
	EndIf
Next

If $dictFilesMovedFromShare.Count = 0 Then
	MessageToAdmin($oUSER,$oMAIL,"I;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ":" & @SEC & ";" & @UserName & ";" & @ComputerName & ";" & $sSAPSystemName,"No files found in the " & $sSourceDirPath & " Exit code: 100",$SYS_ADMINS)
	Exit(100) ; Nothing to do. Network share empty. No files were moved
EndIf

If _CheckSAPLogon($nSAPLogonPID) Then
	; OK, Logon is running
	; PID is saved
Else
	$nSAPLogonPID = _LaunchSAPLogon()
	If $nSAPLogonPID == 0 Then
		Exit(1) ; SAPLogon couldn't be started
	EndIf
EndIf
; At this point we can obtain SAPGUI object
$oGUI = ObjGet("SAPGUI")
While Not IsObj($oGUI) ; Possibly dangerous loop
	$oGUI = ObjGet("SAPGUI")
WEnd
$oSAP = $oGUI.GetScriptingEngine
ObjEvent($oSAP,SAPErrorHandler) ; Register event handler
If Not IsObj($oSAP) Then
	Exit(2); Can't get scripting engine
EndIf

If _FindSAPSystemDescription($sSAPSystemName,$sSAPSystemDescription) Then
	; OK, Description string saved
Else
	Exit(3) ; Can't obtain SAP system description from landscape file
EndIf


If $oSAP.Children.Count == 0 Then ; No connections opened. Open a new one
	$oCON = $oSAP.OpenConnection($sSAPSystemDescription,True,False); Open new connection asynchronously
	If Not IsObj($oCON) Or @error Then
		Exit(4) ; Can't open connection. Connection is not an object
	ElseIf IsObj($oCON) And $oCON.Children.Count == 0 Then
		Exit(5) ; Can't get session object. Connection has zero children. Check permissions to the SAP system
	EndIf
	$oSES = $oCON.Children.Item(0) ; Initialize session variable to the first child of the new connection
	KillPopups($oSES)
	$nSessionsCount = $oCON.Children.Count
	$oSES.CreateSession()
	While $oCON.Children.Count = $nSessionsCount
		Sleep(100)
	WEnd
	$nSessionsCount = $oCON.Children.Count
	$oSES2 = $oCON.Children.Item($nSessionsCount - 1)
	KillPopups($oSES2)
ElseIf $oSAP.Children.Count > 0 Then ; At least one connection exists
	If _GetSAPSession($sSAPSystemName,$oSAP,$oSES,$oCON) Then ; Functions returns True if the target session has been found
		$nSessionsCount = $oCON.Children.Count
		$oSES.CreateSession()
		While $oCON.Children.Count = $nSessionsCount
			Sleep(100)
		WEnd
		$nSessionsCount = $oCON.Children.Count
		$oSES2 = $oCON.Children.Item($nSessionsCount - 1)
		KillPopups($oSES2)
	Else
		; None of the sessions fits. Open a new connection to the target SAP system
		$oCON = $oSAP.OpenConnection($sSAPSystemDescription,True,False); Open a new connection asynchronously
		If Not IsObj($oCON) Or @error Then
			Exit(4) ; Can't open connection. Connection is not an object
		ElseIf (IsObj($oCON) And $oCON.Children.Count == 0) Then
			Exit(5) ; Can't get session object. Connection has zero children. Check permissions to the SAP system
		EndIf
		$oSES = $oCON.Children.Item(0)
		KillPopups($oSES)
		$nSessionsCount = $oCON.Children.Count
		$oSES.CreateSession()
		While $oCON.Children.Count = $nSessionsCount
			Sleep(100)
		WEnd
		$nSessionsCount = $oCON.Children.Count
		$oSES2 = $oCON.Children.Item($nSessionsCount - 1)
		KillPopups($oSES2)
	EndIf
EndIf



PreWork($oWND0, $oWND1, $hwndExp, $hwndDnD, $oSES, $oSES2, $nPIDexp, $sOawdFolder, "Incoming invoice prel posting (PDF)", "HR02_HADES_DND_SOURCE", $sSourceDirPath)
$nFilesMoved = DoWork($hwndDnD,$hwndExp,$oWND0,$oWND1,$oSES,$oSES2, $sSourceDirPath,$nDuplicateFiles,$dictReport,$sProcesseDirPath)
PostWork($hwndExp,$oSES,$dictReport,$SYS_ADMINS,$nFilesMoved,$hTimer,$SP_SITE_URL,$SP_FOLDER_NAME,$oHTTP,$dictFilesMovedFromShare,$dictFilesUploaded,$dictFilesFailedToUpload,$sProcesseDirPath)

Exit(0)


;**************************************************************
;************************ End Main ****************************


;----------------------- Misc functions -------------------------
;----------------------------------------------------------------
Func KillPopups($_oSES)
	While $_oSES.Children.Count > 1
		If StringInStr($_oSES.ActiveWindow.Text, "System Message") > 0 Then
			$_oSES.ActiveWindow.sendVKey(12)
		ElseIf StringInStr($_oSES.ActiveWindow.Text, "Information") > 0 And StringInStr($_oSES.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
			$_oSES.ActiveWindow.sendVKey(0)
		ElseIf StringInStr($_oSES.ActiveWindow.Text, "Copyright") > 0 Then
			$_oSES.ActiveWindow.sendVKey(0)
		ElseIf StringInStr($_oSES.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
			$_oSES.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
			$_oSES.ActiveWindow.sendVKey(0)
		;ElseIF   'Insert next type of popup windows which you want to kill
		Else
			ExitLoop
		EndIf
	Wend
EndFunc

Func MessageToAdmin(ByRef $_oLDAP, ByRef $_oMAIL, $_sSubject, $_sMessage, $_sAdmins)

For $__sAdmin in StringSplit($_sAdmins,";")
	With $_oMAIL
		.From = $_oLDAP.Mail
		.To = $__sAdmin
		.Subject = $_sSubject
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailgot.it.volvo.net"
		.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		.HTMLBody = $_sMessage
		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
		.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
		.Configuration.Fields.Update()
		.Send()
	EndWith
Next

EndFunc

Func PreWork(ByRef $_oWND0, ByRef $_oWND1, ByRef $_hwndExpWindow, ByRef $_hwndDnDWindow, ByRef $_oSAPSession, ByRef $_oSAPSession2, $_nPID, $_sSAPFolderName, $_sSAPAction, $_sSourceDir, $_sSourceDirPath)
	Local $__aDnDSize, $__aExpSize, $__aWindows
	;$__aWindows = WinList($_sSourceDir)
	;MsgBox(0,"OK",$__aWindows[0][0])
	;If $__aWindows[0][0] == 0 Then ; No explorer window with title that matches our source dir exists. Open a new one
	$_nPID = ShellExecute("explorer.exe",$_sSourceDirPath) ; Open explorer and return PID to the caller
	If $_nPID == 0 Then
		Exit(7) ; Failed to start windows explorer
	EndIf
	$__aWindows = WinList($_sSourceDir)
	While $__aWindows[0][0] == 0
		$__aWindows = WinList($_sSourceDir)
	WEnd
	$_hwndExpWindow = $__aWindows[1][1]
	; Open ZFIDOCWID transaction
	$_oSAPSession2.findById("wnd[0]/tbar[0]/okcd").text = "/nzfidocwid"
	$_oSAPSession2.findById("wnd[0]").sendVKey(0)   ; Enter
	KillPopups($_oSAPSession2)
	; Open OAWD transaction and prepare Archive from Frontend window
	$_oSAPSession.findById("wnd[0]/tbar[0]/okcd").text = "/NOAWD"	;Opens SAP tcode for Drag and Drop
	$_oSAPSession.findById("wnd[0]").sendVKey(0)	;Enter key
	KillPopups($_oSAPSession)
	$_oSAPSession.findById("wnd[0]").sendVKey(71)	;Ctrl+F to find string
	$_oSAPSession.findById("wnd[1]/usr/txtRSYSF-STRING").text = $_sSAPFolderName	;string in find control
	$_oSAPSession.findById("wnd[1]").sendVKey(0)	;Enter key
	$_oSAPSession.findById("wnd[2]").sendVKey(84)	;Ctrl+G to point to the searched string
	$_oSAPSession.findById("wnd[2]").sendVKey(2)	;F2 key to select the pointed string and return to previous wnd
	$_oSAPSession.findById("wnd[0]").sendVKey(2)	;F2 key to expand the pointed position
	$_oSAPSession.findById("wnd[0]").sendVKey(71)	;Ctrl+F to find string
	$_oSAPSession.findById("wnd[1]/usr/txtRSYSF-STRING").text = $_sSAPAction	;string in find control
	$_oSAPSession.findById("wnd[1]").sendVKey(0)	;Enter key
	$_oSAPSession.findById("wnd[2]").sendVKey(84)	;Ctrl+G to point to the searched string
	$_oSAPSession.findById("wnd[2]").sendVKey(2)	;F2 key to select the pointed string and return to previous wnd
	$_oSAPSession.findById("wnd[0]").sendVKey(2)	;F2 key to expand the pointed position
	$_oSAPSession.findById("wnd[1]/usr/txtCONFIRM_DATA-NOTE").text = "" ;nazov PDF suboru bez PDF a max 50 znakov
	$_oWND1 = $_oSAPSession.findById("wnd[1]") ; Drag'n'Drop window (wnd[1]) returned to the caller.
	$_oWND0 = $_oSAPSession.findById("wnd[0]") ; Main transaction window (wnd[0]) returned to the caller.
	$__aWindows = WinList($_oWND1.Text)
	While $__aWindows[0][0] == 0
		$__aWindows = WinList($_oWND1.Text)
	WEnd
	$_hwndDnDWindow = $__aWindows[1][1] ; Catch the first window in the list and return to the caller the handle
	WinActivate($_hwndExpWindow,"")
	ExplorerSetView(2,$_hwndExpWindow)
EndFunc



Func DoWork(ByRef $_hwndDnDWindow, ByRef $_hwndExpWindow, ByRef $_oWND0, ByRef $_oWND1, ByRef $_oSES, ByRef $_oSES2, $_sSourceDirectoryPath, $_nDuplicateFiles, ByRef $_dictReport, $_sProcessedDirPath)
	Local $__xstart, $__ystart, $__xstop, $__ystop
	Local $__sPane = ""
	Local $__sShell = ""
	Local $__sShell2 = ""
	Local $__sPreviousFile = "\.xxx" ; bogus file name
	Local $__oFso = ObjCreate("Scripting.FileSystemObject")
	Local $__oFolder = $__oFso.GetFolder($_sSourceDirectoryPath)
	Local $__colFiles = $__oFolder.Files
	Local $__aESize ; Array holding dims and coords of Explorer window
	Local $__aDSize ; Array holding dims and coords of SAP Drag'n'Drop window
	Local $__aMSize ; Arry holding dims and coords of the main working area of the explorer window i.e where the files are displayed
	Local $__aWorkItemID
	Local $__sWorkItemID
	Local $__nFilesMoved = 0


	For $__file In $__colFiles
		$_dictReport.Add($__file.Name,0)
		ReopenArchiveWindow($_oWND1, $__file.Name, "Storing for subsequent entry", "Archive from Frontend")
		KillPopups($_oSES)
		WinActivate($_hwndDnDWindow,"")
		WinMove($_hwndDnDWindow,"",0,0)
		$__aDSize = WinGetPos($_hwndDnDWindow) ; Get position of the drag and drop sap window
		WinActivate($_hwndDnDWindow,"")
		WinActivate($_hwndExpWindow,"")
		WinMove($_hwndExpWindow,"",$__aDSize[2] + 10,0,1000,500)
		$__aMSize = ControlGetPos($_hwndExpWindow,"","DirectUIHWND3") ; Get position of the main working area of the explorer window
		$__aESize = WinGetPos($_hwndExpWindow)			; Get position of the whole explorer window
		$__xstart = $__aESize[0] + $__aMSize[0] + 35; X position of the window + relative position of the main area within the window + offset
		$__ystart = $__aESize[1] + $__aMSize[1] + 15; Y position of the window + relative position of the main area within the window + offset
		$__xstop = $__aDSize[0] + Floor(($__aDSize[2] / 100) * 75)
		$__ystop = $__aDSize[1] + Floor(($__aDSize[3] / 100) * 50)
		DragAndDrop($__xstart,$__ystart,$__xstop,$__ystop)
		$__sPane = $_oWND0.findById("sbar/pane[0]").Text ; Should generate COM error. Catch it using error handler which sets $errno
		While $_oSES.Busy == True Or StringLen($__sPane) == 0 ; Session is busy. usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[1]/shell not yet available. Could generate error if referenced while session is busy
			$__sPane = $_oWND0.findById("sbar/pane[0]").Text ; Should generate COM error. Catch it using error handler which sets $errno
		WEnd
		If $__sPane == "Action completed" Then
			$__sShell = $_oWND1.findById("usr/ssubSUB110:SAPLALINK_DRAG_AND_DROP:0110/cntlSPLITTER/shellcont/shellcont/shell/shellcont[1]/shell").Text
			If StringRegExp($__sShell,$__sPreviousFile) == 1 Then
				; If this condition is true it means that deleting a file in the previous cycle hasn't yet been reflected in the windows explorer
				; At this point file is already uploaded to the SAP, there is no way to prevent duplicit files. Only report them
				$_nDuplicateFiles += 1
				$_oWND1.sendVKey(0) ; Hit enter
			ElseIf StringRegExp($__sShell,$__sPreviousFile) == 0 Then ; OK
				$_oWND1.sendVKey(0) ; Hit enter
			EndIf
		EndIf
		; Check if upload was successfull

		$__aWorkItemID = StringRegExp($__sShell,":[0-9]{12}",$STR_REGEXPARRAYMATCH)
		; We really expect only one match, use the first one
        If UBound($__aWorkItemID) >= 1 Then ; WorkID found
			$__sWorkItemID = StringRight($__aWorkItemID[0],StringLen($__aWorkItemID[0]) - 1 )
			$_oSES2.findById("wnd[0]/usr/txtS_WORKID-LOW").text = $__sWorkItemID
			$_oSES2.findById("wnd[0]").sendVKey(8)
			If $oSES2.findById("wnd[0]/shellcont/shell").RowCount > 0 Then
				For $i = 0 To $oSES2.findById("wnd[0]/shellcont/shell").RowCount - 1
					If $oSES2.findById("wnd[0]/shellcont/shell").GetCellValue($i,"WORKID") = Abs($__sWorkItemID) Then
						If StringRegExp($oSES2.findById("wnd[0]/shellcont/shell").GetCellValue($i,"ZSTATUS"),"Work Flow Created") Then
							$_dictReport.Item($__file.Name) = 0 ; OK
							ExitLoop
						ElseIf StringRegExp($oSES2.findById("wnd[0]/shellcont/shell").GetCellValue($i,"ZSTATUS"),"WFC Failed-Corrupted File") Then
							$_dictReport.Item($__file.Name) = $SAP_FCORUPTED
							ExitLoop
						EndIf
					EndIf
				Next
			Else
				$_dictReport.Item($__file.Name) = $SAP_FNOTFOUND_REPORT
			EndIf
		Else
			$_dictReport.Item($__file.Name) = $_dictReport.Item($__file.Name) + $SAP_FNOTFOUND_SHELL
		EndIf
		$_oSES2.findById("wnd[0]").sendVKey(3)
		$__sPreviousFile = $__file.Name
		$__oFso.MoveFile($_sSourceDirectoryPath & $__file.Name,$_sProcessedDirPath & $__file.Name) ; Move processed file and continue with the next one
		WinActivate($_hwndExpWindow,"")
		ExplorerRefreshWindow($_hwndExpWindow,2000)
		$__nFilesMoved += 1
	Next
	$_oWND1.sendVKey(12) ; Exit
	Return $__nFilesMoved

EndFunc

Func PostWork(ByRef $_hwndExpWindow, ByRef $_oSES, ByRef $_dictReport, $_sAdmins, $_nFilesMoved, $_hTimer,  $_sSpSiteUrl, $_sSpFolder, ByRef $_oHTTP, ByRef $_dictFilesMovedFromShare, ByRef $_dictFilesUploaded, ByRef $_dictFilesFailedToUpload, $_sProcessedDirPath)
	If StringRight($_sProcessedDirPath,1) = "\" Then
		; ok
	Else
		$_sProcessedDirPath = $_sProcessedDirPath & "\"
	EndIf
	Local $__sSubject, $__sBody, $__aSharepointIDs, $__sSecurityToken, $__sXDigest, $__oHTTP, $__oFOLDER, $__nDuration
	Local $__oFso = ObjCreate("Scripting.FileSystemObject")
	$__oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$__aSharepointIDs = SPGetTenantRealmID($__oHTTP, $_sSpSiteUrl)
	If UBound($__aSharepointIDs) = 0 Then
		; do not copy to sharepoint
	Else
		$__sSecurityToken = SPGetSecurityToken($_oHTTP, "volvogroup.sharepoint.com", $__aSharepointIDs[1], $__aSharepointIDs[2], "462ad7ed-2e5e-4175-b808-18c6f33fadd7","dWHEl4AMp8qHX/oxeFcY4RyFJJRD7z1cIavjDH53yIE=")
		If $__sSecurityToken = 1 Then
			; do not copy. Can't get security token
		Else
			$__sXDigest = SPGetXDigestValue($_oHTTP, $_sSpSiteUrl, $__sSecurityToken)
			If $__sXDigest = 1 Then
				; do not copy. Can't get xdigest value
			Else
				$__oFOLDER = $__oFso.GetFolder($sProcesseDirPath)
				For $file in $__oFOLDER.Files
					SPFileUpload($__oHTTP, $_sSpSiteUrl, $__sXDigest, $__sSecurityToken, $_sProcessedDirPath & $file.Name, $file.Name, "HR02_HADES_ARCHIVE_QA", $_dictFilesUploaded, $_dictFilesFailedToUpload)
				Next
			EndIf
		EndIf
	EndIf
	$__nDuration = Floor(TimerDiff($_hTimer) / 1000)
	WinClose($_hwndExpWindow)
	$_oSES.SendCommand("/nex")
	$__aSharepointIDs = SPGetTenantRealmID($_oHTTP,"https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it")
	$__sBody =  "<HEAD><TITLE>HR02_HADES_DND</TITLE></HEAD>" _
	& "<p style=""color:red;"">Total number of files downloaded from " & $sNetworkSharePath & ": " & $dictFilesMovedFromShare.Count & "</p>" _
	& "<p style=""color:green;""><b>SAP file name   |   Original file name</b></p><p>"
	For $key in $_dictFilesMovedFromShare.Keys()
		$__sBody = $__sBody & $key & "   |   " & $_dictFilesMovedFromShare.Item($key) & "<br>"
	Next
	$__sBody = $__sBody & "<p style=""color:red;"">Total number of files moved to SAP: " & $_nFilesMoved & "</p>" _
	& "<p style=""color:green;""><b>File name   |   Status</b></p><p>"
	$__sSubject = "I;" & @ScriptName & ";" & @YEAR & "-" & @MON & "-" & @MDAY & ";" & @HOUR & ":" & @MIN & ":" & @SEC & ";" & @UserName & ";" & @ComputerName & ";" & $sSAPSystemName & ";" & $__nDuration
	For $key in $_dictReport.Keys()
		$__sBody = $__sBody & $key & "   |   " & $_dictReport.Item($key) & "<br>"
	Next
	$__sBody = $__sBody & "<p style=""color:red;"">Total number of files uploaded to Sharepoint: " & $_dictFilesUploaded.Count & "</p><p style=""color:green;""><b>Sharepoint file name   |   Filesystem name</b></p>"
	For $key in $_dictFilesUploaded.Keys()
		$__sBody = $__sBody & $key & "   |   " & $_dictFilesUploaded.Item($key) & "<br>"
	Next
	$__sBody = $__sBody & "<p style=""color:red;"">Total number of duplicate files: " & $nDuplicateFiles & "</p><br><br>" _
	& "<p style=""color:green;"">Processing time</p><p>" & $__nDuration & " seconds</p></BODY></HTML>"
	MessageToAdmin($oUSER, $oMAIL, $__sSubject, $__sBody, $_sAdmins)
EndFunc


;-------------------------- Exception handlers ------------------------
;----------------------------------------------------------------------
Func SAPErrorHandler($_lErrorId, $_sD1, $_sD2, $_sD3, $_sD4)
	;MsgBox(0,"Error",$_lErrorId)
EndFunc

Func COMErrHandler($oError)
#comments-start
Dummy excepction handler
  MsgBox(0,"ERROR","COM Error!"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
#comments-end

Endfunc
;-------------------------------- End -----------------------------------
;------------------------------------------------------------------------


#comments-start
Function opens Archive window because after each uploaded file
this window is closed i.e. it changes it's layout
#comments-end
Func ReopenArchiveWindow(ByRef $_oWND, $_sFileName, $_sTitle1, $_sTitle2)
	If StringRight($_sFileName,4) == ".pdf" Or StringRight($_sFileName,4) == ".PDF" Then
		$_sFileName = StringLeft($_sFileName,StringLen($_sFileName) - 4) ; remove file extension
		$_sFileName = StringLeft($_sFileName,50) ; trim to 50 chars
	EndIf
	If StringRegExp($_oWND.Text,$_sTitle1) Then
		$_oWND.findById("usr/txtCONFIRM_DATA-NOTE").text = $_sFileName
		$_oWND.sendVKey(9) ; F9
	ElseIf StringRegExp($_oWND.Text,$_sTitle2) Then
		$_oWND.sendVKey(12) ; F12
		$_oWND.findById("usr/txtCONFIRM_DATA-NOTE").text = $_sFileName
		$_oWND.sendVKey(9) ; F9
	EndIf
EndFunc


Func DragAndDrop($_nXstart, $_nYstart, $_nXStop, $_nYstop)
	MouseClickDrag("",$_nXstart,$_nYstart,$_nXStop,$_nYstop)
EndFunc


;----------------------------------------------------------------
;---------------------- SAP Functions ---------------------------
Func _CheckSAPLogon(ByRef $_nPidStorage)
	Local $_oWMI = ObjGet("winmgmts:\\.\root\cimv2")
	Local $_colProc = $_oWMI.ExecQuery("Select Name, ProcessId From Win32_Process Where Name Like '%saplogon%'")
	If IsObj($_colProc) and $_colProc.count > 0 Then
		For $_proc in $_colProc
			If StringInStr($_proc.Name,"saplogon",$STR_NOCASESENSE) > 0 Then ; Substring found
				$_nPidStorage = $_proc.ProcessId
				Return True
				ExitLoop
			EndIf
		Next
	EndIf
		$_nPidStorage = Null
		Return False ; If the execution reaches this point no SAPLogon was found
EndFunc

Func _LaunchSAPLogon()
	Local $_pid
	Run("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
	$_pid = ProcessWait("saplogon.exe", 10)
	If $_pid == 0 Then
		Return 0 ; SAPLogon couldn't be started
	Else ; PID is returned
		Return $_pid
	EndIf
EndFunc

Func _FindSAPSystemDescription($_sSystemName, ByRef $_sSystemDescription)
	Local $_oXML = ObjCreate("MSXML2.DOMDocument")
	Local $_colNodes,$_node
	$_oXML.load($SAP_LOCAL_LANDSCAPE_PATH)
	$_colNodes = $_oXML.GetElementsByTagName("Service")
	For $_node in $_colNodes
		If(StringLeft(StringLower($_node.attributes.getNamedItem("name").text),3)) == StringLower($_sSystemName) Then
			$_sSystemDescription = $_node.attributes.getNamedItem("name").text
			Return True
			ExitLoop
		EndIf
	Next
	Return False
EndFunc

Func _GetSAPSession($_sSystemName, ByRef $_oSAP, ByRef $_oSES, ByRef $_oCON)
	For $_oCON in $_oSAP.Children
		For $__oSES in $_oCON.Children
			If $__oSES.Info.SystemName == $_sSystemName Then
				$_oSES = $__oSES ; Return/Set the session object
				$_oCON = $_oSES.Parent ; Return/Set the connection object
				Return True
			EndIf
		Next
	Next

	Return False
EndFunc

;------------------------------------------------
;----------------- end --------------------------

;--------------- Windows Explorer functions ------------
;-------------------------------------------------------
Func ExplorerSetView($_nView, ByRef $_hwndWindow)
	SendKeepActive($_hwndWindow)
	Send("^")
	Send("+")
	Send($_nView)
	SendKeepActive("")
EndFunc

Func ExplorerRefreshWindow(ByRef $_hwndWindow, $_nDelayms)
	SendKeepActive($_hwndWindow)
	Send("{F5}")       ; F5
	SendKeepActive("")
	Sleep($_nDelayms)
EndFunc


;-------------------------------------------------------
;---------------------- end ----------------------------


;-------------------------------------------------------
;-------------- SharePoint functions -------------------
#comments-start
SPGetTenantRealmID(ByRef $ojbect, $string)
Function obtains Bearer realm aka Tenat ID
and ResourceID which in turn are used in order
to obtain the security token
Return value: one-dimensional array with two elements
1st element is number of tokens returned
2nd element is Bearer/TenantID
3rd element is ResourceID
#comments-end
Func SPGetTenantRealmID(ByRef $_oHTTP, $_sSiteUrl)
	Local $__sResponseHeader
	Local $__aRxMatch
	Local $__aReturnValues = [0,"",""]

	If StringRight($_sSiteUrl,1) == "/" Then
		$_sSiteUrl = $_sSiteUrl & "_vti_bin/client.svc"
	Else
		$_sSiteUrl = $_sSiteUrl & "/_vti_bin/client.svc"
	EndIf

	With $_oHTTP
		.open("GET", $_sSiteUrl, False)
		.setRequestHeader("Authorization", "Bearer")
		.send()
	EndWith

	If $_oHTTP.Status == 401 Then ; 401 is expected at this stage
		$__sResponseHeader = $_oHTTP.getResponseHeader("WWW-Authenticate")
		$__aRxMatch = StringRegExp($__sResponseHeader,"realm=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)
		If UBound($__aRxMatch) >= 1 Then
			; Sould really match only once
			$__aReturnValues[1] = StringRight($__aRxMatch[0],StringLen($__aRxMatch[0]) - 7)
			$__aReturnValues[0] = $__aReturnValues[0] + 1
		Else
			$__aReturnValues[0] = $__aReturnValues[0] + 0
		EndIf
		$__aRxMatch = StringRegExp($__sResponseHeader,"client_id=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)
		If UBound($__aRxMatch) >= 1 Then
			; Should really match only once
			$__aReturnValues[2] = StringRight($__aRxMatch[0],StringLen($__aRxMatch[0]) - 11)
			$__aReturnValues[0] = 2
		Else
			$__aReturnValues[0] = $__aReturnValues[0] + 1
		EndIf
		Return $__aReturnValues
	Else
		$__aReturnValues[0] = 0
		Return $__aReturnValues
	EndIf
EndFunc

#comments-start
SPGetSecurityToken(ByRef $object, $string, $string)
Function obtains Security token used in the subsequent
operations
#comments-end
Func SPGetSecurityToken(ByRef $_oHTTP, $_sTenantDomainName, $_sTenantID, $_sResourceID, $_sClientID, $_sClientSecret)
	Local $__aStringSplit1
	Local $__aStringSplit2
	Local $__sToken
	Local $__sHttpBody
	Local $__sAuthUrl1 = "https://accounts.accesscontrol.windows.net/"
	Local $__sAuthUrl2 = "/tokens/OAuth/2"

	If StringLeft($_sTenantDomainName,1) <> "/" Then
		$_sTenantDomainName = "/" & $_sTenantDomainName
	EndIf

	$__sHttpBody = "grant_type=client_credentials&client_id=" & $_sClientID & "@" & $_sTenantID & "&client_secret=" & $_sClientSecret & "&resource=" & $_sResourceID & $_sTenantDomainName & "@" & $_sTenantID

	With $_oHTTP
		.open("POST", $__sAuthUrl1 & $_sTenantID & $__sAuthUrl2, False)
		.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		.setRequestHeader("Content-Length", StringLen($__sHttpBody))
		.send($__sHttpBody)
	EndWith

	If $_oHTTP.status == 200 Then
		$__aStringSplit1 = StringSplit($_oHTTP.responseText,",")
		If $__aStringSplit1[0] == 0 Then
			Return 1 ; Couldn't split response text or there is nothing in the responsetext. Token not found
		Else
			For $__i = 1 To $__aStringSplit1[0]
				$__aStringSplit2 = StringSplit($__aStringSplit1[$__i],":")
				For $__j = 1 To $__aStringSplit2[0]
					If $__aStringSplit2[$__j] == """access_token""" Then
						$__sToken = StringLeft($__aStringSplit2[$__j + 1],StringLen($__aStringSplit2[$__j + 1]) - 2)
						$__sToken = StringRight($__sToken,StringLen($__sToken) - 1)
						Return $__sToken
					EndIf
				Next
			Next
			Return 1 ; Couldn't find token.
		EndIf
	Else
		Return $_oHTTP.status ; HTTP error. Token not found
	EndIf
EndFunc

#comments-start
SPGetXDigestValue(ByRef $object, $string)
#comments-end
Func SPGetXDigestValue(ByRef $_oHTTP, $_sSiteUrl, $_sSecurityToken)
	Local $_aRxMatch
	Local $_sDigest
	If StringRight($_sSiteUrl,1) == "/" Then
		$_sSiteUrl = $_sSiteUrl & "_api/contextinfo"
	Else
		$_sSiteUrl = $_sSiteUrl & "/_api/contextinfo"
	EndIf

	With $_oHTTP
		.open("POST", $_sSiteUrl, False)
		.setRequestHeader("accept", "application/json;odata=verbose")
		.setRequestHeader("authorization", "Bearer " & $_sSecurityToken)
		.send()
	EndWith

	If $_oHTTP.status == 200 Then
		$_aRxMatch = StringRegExp($_oHTTP.responseText,"FormDigestValue"":""0x[a-fA-F0-9]+,",$STR_REGEXPARRAYMATCH)
		If UBound($_aRxMatch) > 0 Then
			$_sDigest = StringRight($_aRxMatch[0], StringLen($_aRxMatch[0]) - 18)
			$_sDigest = StringLeft($_sDigest, StringLen($_sDigest) - 1)
			Return $_sDigest
		EndIf
	EndIf

	Return 1 ; Couldn't find x digest value
EndFunc


#comments-start
SPFolderExists(ByRef $object, $string, $string, $string)
Return value: 1 if folder exists 0 otherwise
#comments-end
Func SPFolderExists(ByRef $_oHTTP, $_sSiteUrl, $_sDirName, $_sSecurityToken)
	If StringRight($_sSiteUrl,1) == "/" Then
		$_sSiteUrl = $_sSiteUrl & "_api/web/GetFolderByServerRelativeUrl('" & $_sDirName & "')"
	Else
		$_sSiteUrl = $_sSiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & $_sDirName & "')"
	EndIf

	With $_oHTTP
		.open("GET", $_sSiteUrl, False)
		.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Accept", "application/json;odata=verbose")
		.send()
	EndWith

	If $_oHTTP.status == 404 Then
		; Folder not found and @error is set
		SetError($_oHTTP.status)
		Return 0
	EndIf

	If StringRegExp($_oHTTP.responseText,"""Exists"":true") Then
		Return 1 ; True
	EndIf

	Return 0 ; False
EndFunc

#comments-start
SPDownloadFolder(ByRef $object, $string, $string, $string, $string)
Checks if the folder exists. If it does downloads it to the target folder
Sharepoint folder is the leaf of the local tree structure
If the target path is C:\Some folder\MyFolder and the sharepoint folder name
is also MyFolder than the final path will be C:\Some folder\MyFolder\MyFolder
#comments-end
Func SPDownloadFolder(ByRef $_oHTTP, $_sSiteUrl, $_sDirName, $_sTargetPath, $_sSecurityToken)
	Local $__nFileCount, $__oXML, $__colItems, $__colPaths, $__oStream, $__sDownloadUrl
	$__oStream = ObjCreate("ADODB.Stream")
	$__colFiles = ObjCreate("scripting.dictionary")

	If StringRight($_sSiteUrl,1) == "/" Then
		$__sDownloadUrl = $_sSiteUrl & "_layouts/15/download.aspx?SourceUrl=" & $_sSiteUrl
		$_sSiteUrl = $_sSiteUrl & "_api/web/GetFolderByServerRelativeUrl('" & $_sDirName & "')"
	Else
		$__sDownloadUrl = $_sSiteUrl & "/_layouts/15/download.aspx?SourceUrl=" & $_sSiteUrl & "/"
		$_sSiteUrl = $_sSiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & $_sDirName & "')"
	EndIf

	With $_oHTTP
		.open("GET", $_sSiteUrl, False)
		.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Accept", "application/json;odata=verbose")
		.send()
	EndWith

	If $_oHTTP.status == 404 Then
		; Folder not found and @error is set
		SetError($_oHTTP.status)
		Return 0
	EndIf

	If StringRegExp($_oHTTP.responseText,"""Exists"":true") Then
		With $_oHTTP
			.open("GET", $_sSiteUrl & "/Files", False)
			.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
			.setRequestHeader("Accept", "application/atom+xml;odata=verbose")
			.send()
		EndWith

		If $_oHTTP.status == 200 Then
			$__oXML = ObjCreate("MSXML2.DOMDocument")
			$__oXML.loadXML($_oHTTP.responseText)
			$__colItems = $__oXML.getElementsByTagName("d:Name")
			$__colPaths = $__oXML.getElementsByTagName("d:ServerRelativeUrl")
			$__nFileCount = $__colItems.length

			If $__nFileCount == 0 Then
				Return 0 ; Nothing to do
			EndIf

			For $__i = 0 To $__nFileCount - 1

				With $_oHTTP
					.open("GET", $__sDownloadUrl & $_sDirName & "/" & $__colItems($__i).text, False)
					.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
					.send()
				EndWith

				If Not $_oHTTP.status == 200 Then
					; Write down which file coulnd't be downloaded
					MsgBox(0,"Error donwloading file", $__colItems($__i).text)
				EndIf

				$__oStream.Open()
				$__oStream.Type = 1
				$__oStream.Write($_oHTTP.responseBody)
				$__oStream.SaveToFile($_sTargetPath & $__colItems($__i).text)
				$__oStream.Close()
			Next

			Return $__nFileCount

		Else
			SetError($_oHTTP.status)
			Return 0
		EndIf
	EndIf

	Return 0


EndFunc
;-------------------------------------------------------
;---------------------- end ----------------------------

#comments-start
SPFileUpload(ByRef $object, $string, $string, $string, $string, $strin, $string, ByRef int)
#comments-end
Func SPFileUpload(ByRef $_oHTTP, $_sSiteUrl, $_sXRequestDigest, $_sSecurityToken, $_sSourceFilePath, $_sTargetFileName, $_sTargetFolder, ByRef $_dictFilesUploaded, ByRef $_dictFilesFailedToUpload)
	If StringRight($_sSiteUrl,1) <> "/" Then
		$_sSiteUrl = $_sSiteUrl & "/"
	EndIf
	Local $__sBuffer, $__nBufferLen, $__oFile, $__sOriginalFileName
	$__sOriginalFileName = $_sTargetFileName
	$__nBufferLen = FileGetSize($_sSourceFilePath)
	$__oFile = FileOpen($_sSourceFilePath,16)
	$__sBuffer = FileRead($__oFile)
    $_sTargetFileName = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & "_" & $_sTargetFileName
	With $_oHTTP
		.open("POST",$_sSiteUrl & "_api/Web/GetFolderByServerRelativeUrl('" & $_sTargetFolder & "')/Files/add(overwrite=false, url='" & $_sTargetFileName & "')", False)
		.setRequestHeader("accept", "application/json;odata=verbose")
		.setRequestHeader("X-RequestDigest", $_sXRequestDigest)
		.setRequestHeader("Authorization", "Bearer " & $_sSecurityToken)
		.setRequestHeader("Content-Length", $__nBufferLen)
		.send($__sBuffer)
	EndWith

	If $_oHTTP.Status = 200 Then
		$_dictFilesUploaded.Add($_sTargetFileName,$__sOriginalFileName)
	Else
		$_dictFilesFailedToUpload.Add($_sTargetFileName,$__sOriginalFileName)
		FileClose($__oFile)
	EndIf
	FileDelete($_sSourceFilePath)
	FileClose($__oFile)
EndFunc
