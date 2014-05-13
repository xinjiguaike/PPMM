#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ICO File\Microsoft Remote Desktop Connection.ico
#AutoIt3Wrapper_Outfile=PPMM_Admin.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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
#include <Misc.au3>
#include <Excel.au3>
#include <File.au3>
#include <Array.au3>
#include <TrayConstants.au3>
#include <_XMLDomWrapper_.au3>

_Singleton("PPMM"); Just run one instance of PPMM

Global Const $PPMM_TITLE = "PSO Public Machines Management"
Global Const $PPMM_PATH = "\\10.224.172.65\PSO_Share\DOC_Center\Individual\PPMM"
Global Const $LAUNCH_RDP_PATH = $PPMM_PATH & "\LaunchRDP.exe"
Global Const $LAUNCH_RDP_PREFIX_TITLE = "LaunchRDP - "
Global Const $PPMM_XML_PATH = $PPMM_PATH & "\PPMM.xml"
Global Const $PPMM_LOG_PATH = $PPMM_PATH & "\PPMM.log"

Global $g_aPCName[0]
Global $g_aDomain[0]
Global $g_aUserName[0]
Global $g_aPassword[0]
Global $g_nCurSelectedIndex = -1
Global $g_nConnectedIndex = -1
Global $g_bBeginCheckWindow = False
Global $g_strRemoteConnectionTitle = ""

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
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

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
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("GUIOnEventMode", 1)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 3)

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

TrayCreateItem("Show PPMM")
TrayItemSetOnEvent(-1, "ShowPPMM")
TrayCreateItem("")
TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "ExitScript")

TraySetOnEvent($TRAY_EVENT_PRIMARYDOUBLE, "OnTrayEvent")
TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "OnTrayEvent")

GUISetState(@SW_SHOW)
SyncFromServer($PPMM_XML_PATH)

<<<<<<< HEAD
=======
AdlibRegister("OnBtnRefreshClicked", 60000)

>>>>>>> fcf24b453819edd4484c559c4363561fab7667c0
While 1
	If $g_bBeginCheckWindow Then
		If WinExists($g_strRemoteConnectionTitle) = 0 Then
			UpdateDesktopState($g_nConnectedIndex)
		EndIf
	EndIf

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

Func ShowPPMM()
	GUISetState(@SW_SHOW)
EndFunc   ;==>ShowPPMM

Func ExitScript()
	Local $nPressed = MsgBox($MB_OKCANCEL, "PPMM", "If exit the PPMM, current connection will be closed, really exit?")
	If $nPressed = $IDOK Then
		ProcessClose("mstsc.exe")
		UpdateDesktopState($g_nConnectedIndex)
	Else
		Return
	EndIf
	AdlibUnRegister("OnBtnRefreshClicked")
	GUIDelete($hGUI)
	Exit
EndFunc   ;==>ExitScript

Func OnTrayEvent()
	If @TRAY_ID = $TRAY_EVENT_PRIMARYDOUBLE Then	GUISetState(@SW_SHOW)
EndFunc   ;==>OnTrayEvent

Func OnEventClose()
	GUISetState(@SW_HIDE)
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
	Local $aItemInfo = _GUICtrlListView_GetItem($listviewConnections, $g_nCurSelectedIndex)
	If $aItemInfo[4] = 0 Then
		MsgBox($MB_ICONWARNING, "PPMM", "Current desktop is being used, you can not edit it!")
		Return 0
	EndIf
	GUICtrlSetData($inputConnectionName, $aItemInfo[3])
	GUICtrlSetData($inputPCName, $g_aPCName[$g_nCurSelectedIndex])
	GUICtrlSetData($inputDomain, $g_aDomain[$g_nCurSelectedIndex])
	GUICtrlSetData($inputUserName, $g_aUserName[$g_nCurSelectedIndex])
	GUICtrlSetData($inputPassword, $g_aPassword[$g_nCurSelectedIndex])

	SetLayer1State($GUI_HIDE)
	SetLayer2State($GUI_SHOW, "Edit")
EndFunc   ;==>OnBtnEditClicked

Func OnBtnDeleteClicked()
	Local $aItemInfo = _GUICtrlListView_GetItem($listviewConnections, $g_nCurSelectedIndex)
	If $aItemInfo[4] = 0 Then
		MsgBox($MB_ICONWARNING, "PPMM", "Current desktop is being used, you can not delete it!")
		Return 0
	EndIf
	_XMLDeleteNode("PPMM/Desktop[" & $g_nCurSelectedIndex + 1 & "]")
	SyncFromServer($PPMM_XML_PATH)
EndFunc   ;==>OnBtnDeleteClicked

Func OnBtnRefreshClicked()
	SyncFromServer($PPMM_XML_PATH)
EndFunc   ;==>OnBtnRefreshClicked

Func OnBtnStartConnectionClicked()
	ConnectRemoteComputer($g_nCurSelectedIndex)
EndFunc   ;==>OnBtnStartConnectionClicked

Func OnBtnCancelClicked()
	SetLayer1State($GUI_SHOW)
	SetLayer2State($GUI_HIDE)
EndFunc   ;==>OnBtnCancelClicked

Func OnBtnSaveClicked()
	Local $strConnectionName = GUICtrlRead($inputConnectionName)
	Local $strPCName = GUICtrlRead($inputPCName)
	Local $strDomain = GUICtrlRead($inputDomain)
	Local $strUserName = GUICtrlRead($inputUserName)
	Local $strPassword = GUICtrlRead($inputPassword)
	If $strConnectionName == "" Or $strPCName == "" Or $strDomain == "" Or $strUserName == "" Or $strPassword == "" Then
		MsgBox($MB_ICONWARNING, "PPMM", "Please don't leave any field empty!")
		Return 0
	EndIf

	If @GUI_CtrlId = $btnSaveNew Then
		AddNewDesktop($PPMM_XML_PATH, $strConnectionName, $strPCName, $strDomain, $strUserName, $strPassword)
		SyncFromServer($PPMM_XML_PATH)
		SetLayer1State($GUI_SHOW)
		SetLayer2State($GUI_HIDE)
	ElseIf @GUI_CtrlId = $btnSaveEdit Then
		UpdateSlecetedDesktop($PPMM_XML_PATH, $g_nCurSelectedIndex, $strConnectionName, $strPCName, $strDomain, $strUserName, $strPassword)
		SyncFromServer($PPMM_XML_PATH)
		SetLayer1State($GUI_SHOW)
		SetLayer2State($GUI_HIDE)
	EndIf
EndFunc   ;==>OnBtnSaveClicked

Func SyncFromServer($strFilePath)
	;================ Clear all Array =========
	ReDim $g_aPCName[0]
	ReDim $g_aDomain[0]
	ReDim $g_aUserName[0]
	ReDim $g_aPassword[0]
	;==========================================
	_GUICtrlListView_DeleteAllItems($listviewConnections);Before Syncing, clear all

	_XMLFileOpen($strFilePath)
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

		Local $itemListView = GUICtrlCreateListViewItem("", $listviewConnections)
		If $strAvailable == "N" Then
			_GUICtrlListView_AddSubItem($listviewConnections, $i - 1, $strConnectionName, 0, 0)
			_GUICtrlListView_AddSubItem($listviewConnections, $i - 1, $strOwner, 1)
		Else
			_GUICtrlListView_AddSubItem($listviewConnections, $i - 1, $strConnectionName, 0, 1)
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

Func AddNewDesktop($strFilePath, $strConnectionName, $strPCName, $strDomain, $strUserName, $strPassword)
	_XMLFileOpen($strFilePath)
	_XMLCreateRootNodeWAttr("Desktop", "ConnectionName", $strConnectionName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "PCName", $strPCName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Domain", $strDomain)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "UserName", $strUserName)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Password", $strPassword)
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Available", "Y")
	_XMLCreateChildNode("/PPMM/Desktop[@ConnectionName='" & $strConnectionName & "']", "Owner", "")
EndFunc   ;==>AddNewDesktop

Func UpdateSlecetedDesktop($strFilePath, $nIndex, $strConnectionName, $strPCName, $strDomain, $strUserName, $strPassword)
	_XMLFileOpen($strFilePath)
	_XMLSetAttrib("/PPMM/Desktop", "ConnectionName", $strConnectionName, $nIndex)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/PCName", $strPCName)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/Domain", $strDomain)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/UserName", $strUserName)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/Password", $strPassword)
EndFunc   ;==>UpdateSlecetedDesktop

Func UpdateDesktopState($nIndex, $strAvailable = "Y", $strOwner = "", $nItemImage = 1)
	_GUICtrlListView_SetItemImage($listviewConnections, $nIndex, $nItemImage)
	If $strAvailable == "N" Then
		_GUICtrlListView_AddSubItem($listviewConnections, $nIndex, $strOwner, 1)
	Else
		_GUICtrlListView_SetItemText($listviewConnections, $nIndex, $strOwner, 1)
	EndIf

	_XMLFileOpen($PPMM_XML_PATH)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/Available", $strAvailable)
	_XMLUpdateField("/PPMM/Desktop[" & $nIndex + 1 & "]/Owner", $strOwner)
EndFunc   ;==>UpdateDesktopState

Func ConnectRemoteComputer($nIndex)
	Local $aItemInfo = _GUICtrlListView_GetItem($listviewConnections, $nIndex)
	If $aItemInfo[4] = 0 Then
		Local $strOwner = _GUICtrlListView_GetItemText($listviewConnections, $nIndex, 1)
		MsgBox($MB_ICONWARNING, "PPMM", "Current desktop is being used by " & $strOwner & ", please select another desktop!")
		Return 0
	EndIf

	Local $strYourName = GUICtrlRead($inputYourName)
	If $strYourName == "" Then
		MsgBox($MB_ICONWARNING, "PPMM", "Before connecting the desktop you must input your name!")
		Return 0
	EndIf

	Local $strCMDLine = $LAUNCH_RDP_PATH & " " & $g_aPCName[$nIndex] & " 3389 " & $g_aUserName[$nIndex] & " " & $g_aDomain[$nIndex] & " " & $g_aPassword[$nIndex] & " 0 0 0"
	Run($strCMDLine)
	$g_strRemoteConnectionTitle = $LAUNCH_RDP_PREFIX_TITLE & $g_aPCName[$nIndex]
	Local $hwndLaunchRDP = WinWait($g_strRemoteConnectionTitle, "", 30)
	If $hwndLaunchRDP = 0 Then Return 0

	$g_nConnectedIndex = $nIndex
	$g_bBeginCheckWindow = True

	UpdateDesktopState($g_nConnectedIndex, "N", $strYourName, 0)

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
				Case $LVN_ITEMCHANGED ; Sent by a list-view control when the user clicks an item with the left mouse button
					Local $aItemsIndex = _GUICtrlListView_GetSelectedIndices($listviewConnections, True)
					If $aItemsIndex[0] == 1 Then
						$g_nCurSelectedIndex = $aItemsIndex[1]
					Else
						$g_nCurSelectedIndex = -1
					EndIf
					If $g_nCurSelectedIndex = -1 Then
						GUICtrlSetState($btnStartConnection, $GUI_DISABLE)
						GUICtrlSetState($btnEditConnection, $GUI_DISABLE)
						GUICtrlSetState($btnDeleteConnection, $GUI_DISABLE)
					Else
						GUICtrlSetState($btnStartConnection, $GUI_ENABLE)
						GUICtrlSetState($btnEditConnection, $GUI_ENABLE)
						GUICtrlSetState($btnDeleteConnection, $GUI_ENABLE)
					EndIf
				Case $NM_DBLCLK ; Sent by a list-view control when the user clicks an item with the right mouse button
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
					$g_nCurSelectedIndex = DllStructGetData($tInfo, "Index")
					If $g_nCurSelectedIndex <> -1 Then ConnectRemoteComputer($g_nCurSelectedIndex)
			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

