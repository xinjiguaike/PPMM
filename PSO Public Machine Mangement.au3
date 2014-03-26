#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.10.2
	Author:         Rudy

	Script Function:
	Manage PSO public machines.

#ce ----------------------------------------------------------------------------
#include <MsgBoxConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ColorConstants.au3>
#include <GuiListView.au3>
#include <GuiImageList.au3>
#include <Excel.au3>
#include <File.au3>
#include <Array.au3>
#include <_XMLDomWrapper_.au3>

#cs
	Global Const $REMOTE_CONNECT_TITLE_CH = "远程桌面连接"
	Global Const $SECURITY_WARNING_TITLE_CH = "Windows 安全"
	Global Const $CONFIRM_CONNECT_TITLE_CH = "[TITLE:远程桌面连接;CLASS:#32770;INSTANCE:1]"
	Global Const $INPUT_COMPUTER_TITLE = "[CLASS:ComboBoxEx32;ID:5012]"
	Global Const $BUTTON_CONNECT_TITLE = "[CLASS:Button;ID:1]"
	Global Const $DISPLAY_OPTIONS_TITLE = "[CLASS:ToolbarWindow32;ID:5017]"
	Global Const $IN_INPUT_COMPUTER_TITLE = "[CLASS:ComboBoxEx32;INSTANCE:1]"
	Global Const $IN_INPUT_USERNAME_TITLE = "[CLASS:Edit;ID:13064]"
	Global Const $EDIT_PASSWORD_TITLE = "[CLASS:Edit;INSTANCE:1]"
	Global Const $BUTTON_CONFIRM_TITLE = "[CLASS:Button;INSTANCE:2]"
	Global Const $BUTTON_YES_TITLE = "[CLASS:Button;ID:14004]"
#ce

Global Const $PPMM_TITLE = "PSO Public Machines Management"
Global Const $PPMM_PATH = "\\pso.hz.webex.com\PSO_Share\DOC_Center\Individual\PPMM"
Global Const $LAUNCH_RDP_PATH = @ScriptDir & "\LaunchRDP.exe"
Global Const $PPMM_FILE_PATH = @ScriptDir & "\PSOPublicMachines.xlsx"
Global Const $LAUNCH_RDP_PREFIX_TITLE = "LaunchRDP - "

Global $g_aPCName[0]
Global $g_aDomain[0]
Global $g_aUserName[0]
Global $g_aPassword[0]

Global $g_bPPMMExcelOpened = False
Global $g_strLogPath = @ScriptDir & "\PPMM.log"
Global $g_nCurSelectedIndex = -1

Global $g_strXMLPath = @ScriptDir & "\PPMM.xml"

;AddNewDesktop($g_strXMLPath, "XP_18", "10.224.168.18", "Cisco", "cisco", "pass")
;_XMLFileOpen($g_strXMLPath)
;Local $nCount = _XMLGetAttrib("/PPMM/Desktop[3]", "ConnectionName")
;_ArrayDisplay($nCount)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $nCount = ' & $nCount & @crlf & '>Error code: ' & @error & @crlf) ;### Debug Console
;Exit

Func AddNewDesktop($strFilePath, $strConnectionName, $strPCName, $strDomain, $strUserName, $strPassword)
	_XMLFileOpen($strFilePath)
	_XMLCreateRootNodeWAttr("Desktop", "ConnectionName", $strConnectionName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "PCName", $strPCName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Domain", $strDomain)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "UserName", $strUserName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Password", $strPassword)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Available", "Y")
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Owner", "")
EndFunc


Local $hGUI = GUICreate($PPMM_TITLE, 400, 450, 400, 180)
#Region================= Layer 1 ============================================
Local $btnNewConnection = GUICtrlCreateButton("New", 50, 20, 60, 30)
Local $btnEditConnection = GUICtrlCreateButton("Edit", 130, 20, 60, 30)
Local $btnDeleteConnection = GUICtrlCreateButton("Delete", 210, 20, 60, 30)
Local $btnRefreshConnection = GUICtrlCreateButton("Refresh", 290, 20, 60, 30)
GUICtrlSetState($btnEditConnection, $GUI_DISABLE)
GUICtrlSetState($btnDeleteConnection, $GUI_DISABLE)

Local $listviewConnections = GUICtrlCreateListView("", 50, 80, 300, 250)
GUICtrlSetBkColor($listviewConnections, 0xffffee)
Local $hImage = _GUIImageList_Create(16, 32)
_GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap(GUICtrlGetHandle($listviewConnections), 0xFF0000, 16, 32)); 0 for red
_GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap(GUICtrlGetHandle($listviewConnections), 0x00FF00, 16, 32)); 1 for green
_GUICtrlListView_SetImageList($listviewConnections, $hImage, 1)
_GUICtrlListView_AddColumn($listviewConnections, "Machines", 150)
_GUICtrlListView_AddColumn($listviewConnections, "Owner", 145)

Local $lableYourName = GUICtrlCreateLabel("Your Name:", 110, 350, 60, 23)
Local $inputYourName = GUICtrlCreateInput("", 175, 347, 100, 20)

Local $btnStartConnection = GUICtrlCreateButton("Start Connection", 125, 390, 150, 30)
GUICtrlSetState($btnStartConnection, $GUI_DISABLE)

GUISetOnEvent($GUI_EVENT_CLOSE, "OnEventClose")
GUICtrlSetOnEvent($btnNewConnection, "OnBtnNewClicked")
GUICtrlSetOnEvent($btnEditConnection, "OnBtnEditClicked")
GUICtrlSetOnEvent($btnDeleteConnection, "OnBtnDeleteClicked")
GUICtrlSetOnEvent($btnRefreshConnection, "OnBtnRefreshClicked")
GUICtrlSetOnEvent($btnStartConnection, "OnBtnStartConnectionClicked")
#EndRegion

#Region================= Layer 2 ============================================
Local $lableEditRemoteDesktop = GUICtrlCreateLabel("Edit Remote Desktop", 140, 20, 160, 23)
GUICtrlSetFont($lableEditRemoteDesktop, 10, $FW_BOLD)

Local $lableConnectionName = GUICtrlCreateLabel("Connection Name:", 45, 70, 100, 23)
Local $inputConnectionName = GUICtrlCreateInput("", 150, 65, 200, 25)

Local $lablePCName = GUICtrlCreateLabel("PC Name:", 95, 120, 60, 23)
Local $inputPCName = GUICtrlCreateInput("", 150, 115, 200, 25)

Local $lableDomain = GUICtrlCreateLabel("Domain:", 100, 170, 50, 23)
Local $inputDomain = GUICtrlCreateInput("", 150, 165, 200, 25)

Local $lableUserName = GUICtrlCreateLabel("User Name:", 85, 220, 60, 23)
Local $inputUserName = GUICtrlCreateInput("", 150, 215, 200, 25)

Local $lablePassword = GUICtrlCreateLabel("Password:", 92, 270, 60, 23)
Local $inputPassword = GUICtrlCreateInput("", 150, 265, 200, 25)

Local $btnSaveNew = GUICtrlCreateButton("Save", 100, 350, 100, 30)
Local $btnSaveEdit = GUICtrlCreateButton("Save", 100, 350, 100, 30)
Local $btnCancel = GUICtrlCreateButton("Cancel", 220, 350, 100, 30)

GUICtrlSetOnEvent($btnCancel, "OnBtnCancelClicked")
GUICtrlSetOnEvent($btnSaveNew, "OnBtnSaveClicked")
GUICtrlSetOnEvent($btnSaveEdit, "OnBtnSaveClicked")
SetLayer2State($GUI_HIDE)
#EndRegion

SyncFromServer($PPMM_FILE_PATH)

Opt("GUIOnEventMode", 1)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

GUISetState(@SW_SHOW)

While 1
	Sleep(100)
WEnd


Func SetLayer1State($nState)
	GUICtrlSetState($btnNewConnection, $nState)
	GUICtrlSetState($btnEditConnection, $nState)
	GUICtrlSetState($btnDeleteConnection, $nState)
	GUICtrlSetState($btnRefreshConnection, $nState)
	GUICtrlSetState($listviewConnections, $nState)
	GUICtrlSetState($btnStartConnection, $nState)
	GUICtrlSetState($lableYourName, $nState)
	GUICtrlSetState($inputYourName, $nState)
EndFunc   ;==>SetLayer1State

Func SetLayer2State($nState, $strType = "All")
	GUICtrlSetState($lableEditRemoteDesktop, $nState)
	GUICtrlSetState($lableConnectionName, $nState)
	GUICtrlSetState($inputConnectionName, $nState)
	GUICtrlSetState($lablePCName, $nState)
	GUICtrlSetState($inputPCName, $nState)
	GUICtrlSetState($lableDomain, $nState)
	GUICtrlSetState($inputDomain, $nState)
	GUICtrlSetState($lableUserName, $nState)
	GUICtrlSetState($inputUserName, $nState)
	GUICtrlSetState($lablePassword, $nState)
	GUICtrlSetState($inputPassword, $nState)
	GUICtrlSetState($btnCancel, $nState)
	If $strType == "New" Then
		GUICtrlSetState($btnSaveNew, $nState)
	ElseIf $strType == "Edit" Then
		GUICtrlSetState($btnSaveEdit, $nState)
	Else
		GUICtrlSetState($btnSaveNew, $nState)
		GUICtrlSetState($btnSaveEdit, $nState)
	EndIf
EndFunc   ;==>SetLayer2State

Func OnEventClose()
	GUIDelete($hGUI)
	Exit
EndFunc   ;==>OnEventClose

Func OnBtnNewClicked()
	SetLayer1State($GUI_HIDE)
	GUICtrlSetData($inputConnectionName, "")
	GUICtrlSetData($inputPCName, "")
	GUICtrlSetData($inputDomain, "")
	GUICtrlSetData($inputUserName, "")
	GUICtrlSetData($inputPassword, "")
	SetLayer2State($GUI_SHOW, "New")
EndFunc   ;==>OnBtnNewClicked

Func OnBtnEditClicked()
	SetLayer1State($GUI_HIDE)
	Local $strConnectionName = _GUICtrlListView_GetItemText($listviewConnections, $g_nCurSelectedIndex)
	GUICtrlSetData($inputConnectionName, $strConnectionName)
	GUICtrlSetData($inputPCName, $g_aPCName[$g_nCurSelectedIndex])
	GUICtrlSetData($inputDomain, $g_aDomain[$g_nCurSelectedIndex])
	GUICtrlSetData($inputUserName, $g_aUserName[$g_nCurSelectedIndex])
	GUICtrlSetData($inputPassword, $g_aPassword[$g_nCurSelectedIndex])

	SetLayer2State($GUI_SHOW, "Edit")
EndFunc   ;==>OnBtnEditClicked

Func OnBtnDeleteClicked()
	_GUICtrlListView_DeleteItemsSelected($listviewConnections)
EndFunc   ;==>OnBtnDeleteClicked

Func OnBtnRefreshClicked()
	SyncFromServer($PPMM_FILE_PATH)
EndFunc   ;==>OnBtnRefreshClicked

Func OnBtnStartConnectionClicked()
	ConnectRemoteComputer($g_nCurSelectedIndex)
EndFunc   ;==>OnBtnStartConnectionClicked

Func OnBtnCancelClicked()
	SetLayer1State($GUI_SHOW)
	SetLayer2State($GUI_HIDE)
EndFunc   ;==>OnBtnCancelClicked

Func OnBtnSaveClicked()
	If @GUI_CtrlId = $btnSaveNew Then
		ConsoleWrite("Save New." & @CR)
	ElseIf @GUI_CtrlId = $btnSaveEdit Then
		ConsoleWrite("Save Edit." & @CR)
	EndIf
EndFunc   ;==>OnBtnSaveClicked

Func SyncFromServer($strFilePath)
	_GUICtrlListView_DeleteAllItems($listviewConnections);Before Syncing, clear all
	;================ To Do ===================

	;Clear all Array

	;==========================================
	_XMLFileOpen($g_strXMLPath)
	Local $nDesktopCount = _XMLGetNodeCount("/PPMM/Desktop")
	Local $bGray = True

	For $i = 1 To $nDesktopCount Step 1
		Local $strConnectionName = _XMLGetAttrib("/PPMM/Desktop[" & $i & "]", "ConnectionName")
		Local $strOwner = _XMLGetValue("/PPMM/Desktop[" & $i & "]/Owner")[1]
		Local $strAvailable = _XMLGetValue("/PPMM/Desktop[" & $i & "]/Available")[1]
		Local $strPCName = _XMLGetValue("/PPMM/Desktop[" & $i & "]/PCName")[1]
		Local $strDomain = _XMLGetValue("/PPMM/Desktop[" & $i & "]/Domain")[1]
		Local $strUserName = _XMLGetValue("/PPMM/Desktop[" & $i & "]/UserName")[1]
		Local $strPassword = _XMLGetValue("/PPMM/Desktop[" & $i & "]/Password")[1]

		$itemListView = GUICtrlCreateListViewItem("", $listviewConnections)
		If $strAvailable == "N" Then
			_GUICtrlListView_AddSubItem($listviewConnections, $i-1, $strConnectionName, 0, 0)
			_GUICtrlListView_AddSubItem($listviewConnections, $i-1, $strOwner, 1)
		Else
			_GUICtrlListView_AddSubItem($listviewConnections, $i-1, $strConnectionName, 0, 1)
		EndIf
		$bGray = Not $bGray
		If $bGray Then GUICtrlSetBkColor($itemListView, 0xE9F0FE)

		_ArrayAdd($g_aPCName, $strPCName)
		_ArrayAdd($g_aDomain, $strDomain)
		_ArrayAdd($g_aUserName, $strUserName)
		_ArrayAdd($g_aPassword, $strPassword)
	Next
EndFunc   ;==>SyncFromServer

Func SyncToServer($nIndex)
	Return 0
EndFunc   ;==>SyncToServer

Func AddNewMachine()

EndFunc   ;==>AddNewMachine


Func ConnectRemoteComputer($nIndex)
	Local $aItemInfo = _GUICtrlListView_GetItem($listviewConnections, $nIndex)
	If $aItemInfo[4] = 0 Then
		MsgBox($MB_ICONERROR, "PPMM", "Current desktop is being used, please select another desktop!")
		Return 0
	EndIf

	Local $strYourName = GUICtrlRead($inputYourName)
	If $strYourName == "" Then
		MsgBox($MB_ICONWARNING, "PPMM", "Before connecting the desktop you must input your name!")
		Return 0
	EndIf

	Local $strCMDLine = $LAUNCH_RDP_PATH & " " & $g_aPCName[$nIndex] & " 3389 " & $g_aUserName[$nIndex] & " " & $g_aDomain[$nIndex] & " " & $g_aPassword[$nIndex] & " 0 0 0"
	;ConsoleWrite($strCMDLine & @CR)
	Run($strCMDLine)
	Local $strRemoteConnectionTitle = $LAUNCH_RDP_PREFIX_TITLE & $g_aPCName[$nIndex]
	Local $hwndLaunchRDP = WinWait($strRemoteConnectionTitle, "", 30)
	If $hwndLaunchRDP = 0 Then Return 0

	_GUICtrlListView_SetItemImage($listviewConnections, $nIndex, 0)
	_GUICtrlListView_AddSubItem($listviewConnections, $nIndex, $strYourName, 1)

	SyncToServer($nIndex)

	Return 1

EndFunc   ;==>ConnectRemoteComputer

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam
	Local $hWndFrom, $iCode, $tNMHDR, $hWndListView, $tInfo
	$hWndListView = $listviewConnections
	If Not IsHWnd($listviewConnections) Then $hWndListView = GUICtrlGetHandle($listviewConnections)

	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case $NM_CLICK ; Sent by a list-view control when the user clicks an item with the left mouse button
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
					$g_nCurSelectedIndex = DllStructGetData($tInfo, "Index")
					If $g_nCurSelectedIndex = -1 Then
						GUICtrlSetState($btnStartConnection, $GUI_DISABLE)
						GUICtrlSetState($btnEditConnection, $GUI_DISABLE)
						GUICtrlSetState($btnDeleteConnection, $GUI_DISABLE)
					Else
						GUICtrlSetState($btnStartConnection, $GUI_ENABLE)
						GUICtrlSetState($btnEditConnection, $GUI_ENABLE)
						GUICtrlSetState($btnDeleteConnection, $GUI_ENABLE)
					EndIf
					Return 0
				Case $NM_DBLCLK ; Sent by a list-view control when the user clicks an item with the right mouse button
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
					$g_nCurSelectedIndex = DllStructGetData($tInfo, "Index")
					ConnectRemoteComputer($g_nCurSelectedIndex)
					Return 0 ; allow the default processing
				Case $LVN_KEYDOWN
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
					;$g_nCurSelectedIndex = DllStructGetData($tInfo, "Index")
					ConsoleWrite("$g_nCurSelectedIndex = " & $g_nCurSelectedIndex & @CR)
					Return 0
			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY



#cs
	Local $bRet = 0
	Local $hwndRemoteConnect = CheckWindowExist($REMOTE_CONNECT_TITLE_CH, 30)
	If $hwndRemoteConnect = 0 Then Return 0
	WinActivate($hwndRemoteConnect)

	#Region==================================
	;click 'show options'
	$bRet = CheckControlAction($hwndRemoteConnect, $DISPLAY_OPTIONS_TITLE, "click")
	If $bRet = 0 Then Return 0

	Sleep(3000)

	;input computer name
	$bRet = CheckControlAction($hwndRemoteConnect, $IN_INPUT_COMPUTER_TITLE, "text", $strPCName)
	If $bRet = 0 Then Return 0

	;input user name
	$bRet = CheckControlAction($hwndRemoteConnect, $IN_INPUT_USERNAME_TITLE, "text", $strUserName)
	If $bRet = 0 Then Return 0

	;click 'connect'
	$bRet = CheckControlAction($hwndRemoteConnect, $BUTTON_CONNECT_TITLE, "click")
	If $bRet = 0 Then Return 0
	#EndRegion

	Local $hwndSecurityWarning = CheckWindowExist($SECURITY_WARNING_TITLE_CH, 30)
	If $hwndSecurityWarning = 0 Then Return 0

	#Region=======================
	;input password
	$bRet = CheckControlAction($hwndSecurityWarning, $EDIT_PASSWORD_TITLE, "text", $strPassword)
	If $bRet = 0 Then Return 0

	;click 'confirm' button
	$bRet = CheckControlAction($hwndSecurityWarning, $BUTTON_CONFIRM_TITLE, "click")
	If $bRet = 0 Then Return 0
	#EndRegion

	Local $hwndConfirmConnect = CheckWindowExist($CONFIRM_CONNECT_TITLE_CH, 30)
	If $hwndConfirmConnect Then
	$bRet = CheckControlAction($hwndConfirmConnect, $BUTTON_YES_TITLE, "click")
	If $bRet Then Return 0
	EndIf

	EndFunc

	Func CheckControlAction($title, $controlID, $strAction, $strText="")
	Local $hControl = ControlGetHandle($title, "", $controlID)
	If $hControl Then
	If $strAction == "text" Then
	ControlSetText($title, "", $hControl, $strText)
	ElseIf $strAction == "click" Then
	ControlClick($title, "", $hControl)
	EndIf
	Else
	MsgBox($MB_ICONWARNING, $PPMM_TITLE, "No contorl " & $controlID & " found!")
	Return 0
	EndIf

	Return 1
	EndFunc


	Func CheckWindowExist($title, $nTimeOut)
	Local $hwndConnectWindow = WinWait($title, "", $nTimeOut)
	If $hwndConnectWindow = 0 Then
	MsgBox($MB_ICONWARNING, $PPMM_TITLE, "No [" & $title & "] window pop up!")
	Return 0
	EndIf
	$hwndConnectWindow = WinGetHandle($title)

	Return $hwndConnectWindow
#ce


