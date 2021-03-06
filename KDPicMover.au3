﻿#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile=KDPicMover.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.1.1
#AutoIt3Wrapper_Res_Language=1028
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs

Changelog:
2013/05/06 1.0.1.1 by tsaikd@gmail.com
support view multiple image once

2013/04/28 1.0.1.0 by tsaikd@gmail.com
add change screen button

2013/01/19 1.0.0.9 by tsaikd@gmail.com
Open image file with other App

2009/08/19 1.0.0.8 by tsaikd@gmail.com
Show rest file count

2009/08/18 1.0.0.7 by tsaikd@gmail.com
Show appver in app title

2009/08/13 1.0.0.6 by tsaikd@gmail.com
Change hotkey

2008/12/17 1.0.0.5 by tsaikd@gmail.com
Auto fit desktop size

2008/09/22 1.0.0.4 by tsaikd@gmail.com
Fix Bug: Stack overflow if misc dir exists non picture files

2008/09/21 1.0.0.3 by tsaikd@gmail.com
user can change the dir path
add special directory
Fix Bug: When moving the last picture in misc dir will enter a infinite loop
Change parameter of image display size

2008/09/19 1.0.0.2 by tsaikd@gmail.com
Fix Bug: If Misc Picture is empty will not show application window at startup

2008/09/07 1.0.0.1 by tsaikd@gmail.com
First Release

#ce

#include <GUIConstants.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#include <file.au3>
#include <Date.au3>
#include <SQLite.au3>
#include <Math.au3>
#Include <GDIPlus.au3>
#include <Misc.au3>
#include "Display_library_functions.au3"

; Variable Definition
Global Const $appname = "KDPicMover"
Global Const $appver = "1.0.0.9"
Global Const $appdate = "2013/01/19"
Global Const $author = "tsaikd@gmail.com"

Global Const $appsql = @WorkingDir&"\"&$appname&".sqlite"
Global Const $appini = @WorkingDir&"\"&$appname&".ini"

; Initialization
Global $ie, $ieActiveX, $iIEW, $iIEH

Global $sql
Global $appgui
Global $hWndHotKey
Global Const $app = $appname&" "&$appver

Global $hPicList = -1
Global $sCurPicPath

Global $iFileCount = 0
Global $lblFileCount
Global $lblPicPath

Global $appwidth = 1000
Global $appheight = 700
Global $sPicMiscDir
Global $sPicRenameBeautyDir
Global $sPicRenamePrettyDir
Global $sPicRenameSpecialDir
Global $sOpenAppPath

Global $btnBeauty
Global $btnPretty
Global $btnSpecial
Global $btnOpen
Global $btnBrowse
Global $btnDelete
Global $btnReload
Global $btnScreen

#cs
Pleace set AutoItWinSetTitle($appname) first
Or return value will always be false
#ce
Func IsAppActive()
	Return BitAND(WinGetState(AutoItWinGetTitle()), 8) == 8
EndFunc

Func Main()
	_GDIPlus_Startup()
	_IEErrorHandlerRegister()
	If Not InitPath() Then Return MsgBox(0x10, $app, _("Initialize path failed"))
	If Not InitSQL() Then Return MsgBox(0x10, $app, _("Initialize SQL failed"))
	InitPicList() ; return false if misc picture dir is empty

	$screenInfo = _NumberAndNameMonitors()
	If $screenInfo[0][0] > 1 Then
		Dim $iScreen = IniRead($appini, "Global", "iScreen", 2)
		$ks = _DisplayKeySettings($screenInfo)
		Dim $a = _NumberAndNameMonitors()
		$appwidth = $ks[$iScreen][4] * 0.95
		$appheight = $ks[$iScreen][5] * 0.95 - 50
		$appgui = GUICreate($appname, $appwidth, $appheight, $ks[$iScreen][2] + $ks[$iScreen][4] * 0.025, $ks[$iScreen][3] + $ks[$iScreen][5] * 0.025)
	Else
		$appwidth = @DesktopWidth * 0.95
		$appheight = @DesktopHeight * 0.95 - 50
		$appgui = GUICreate($appname, $appwidth, $appheight)
	EndIf

	$ie = _IECreateEmbedded()
	AutoItWinSetTitle($appname)

	$iWinBH = 5
	$iCtrlGap = 10
	$iCtrlH = 30
	$iLblH = 14
	$iLblO = ($iCtrlH-$iLblH)/2
	$iBtnH = 25
	$iBtnW = 100
	$iBtnG = 150
	$iFileCountW = 40
	$iIEW = $appwidth-$iCtrlGap*2
	$iIEH = $appheight-$iCtrlH*2-$iCtrlGap-$iWinBH

	$ieActiveX = GUICtrlCreateObj($ie, $iCtrlGap, $iCtrlGap, $iIEW, $iIEH)
	$lblFileCount = GUICtrlCreateLabel($iFileCount, $iCtrlGap, $appheight-$iCtrlH*2+$iLblO-$iWinBH, $iFileCountW, $iLblH)
	$lblPicPath = GUICtrlCreateLabel("", $iCtrlGap*2+$iFileCountW, $appheight-$iCtrlH*2+$iLblO-$iWinBH, $appwidth-$iCtrlGap*3-$iFileCountW, $iLblH)
	$btnBeauty = GUICtrlCreateButton(_("Beauty(&1)"), $appwidth/2-$iBtnW/2-$iBtnG*2.25, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnPretty = GUICtrlCreateButton(_("Pretty(&2)"), $appwidth/2-$iBtnW/2-$iBtnG*1.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnSpecial = GUICtrlCreateButton(_("Special(&3)"), $appwidth/2-$iBtnW/2-$iBtnG*0.75, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnOpen = GUICtrlCreateButton(_("&Open"), $appwidth/2-$iBtnW/2-$iBtnG*0, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnBrowse = GUICtrlCreateButton(_("Browse"), $appwidth/2-$iBtnW/2+$iBtnG*0.75, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnDelete = GUICtrlCreateButton(_("Delete(&D)"), $appwidth/2-$iBtnW/2+$iBtnG*1.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnReload = GUICtrlCreateButton(_("Reload(&R)"), $appwidth/2-$iBtnW/2+$iBtnG*2.25, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	If $screenInfo[0][0] > 1 Then
		$btnScreen = GUICtrlCreateButton(_("Screen(&N)"), $appwidth/2-$iBtnW/2+$iBtnG*3, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	EndIf

	$btnMsgExit = GUICtrlCreateButton("Exit", 0, 0)
	GUIUpdatePath()
	GUICtrlSetState(-1, $GUI_HIDE)
	_IENavigate($ie, "about:blank")

	Dim $aAccelKeys[9][2] = [ _
		["{ESC}"	, $btnMsgExit], _
		["1"		, $btnBeauty], _
		["2"		, $btnPretty], _
		["3"		, $btnSpecial], _
		["d"		, $btnDelete], _
		["o"		, $btnOpen], _
		["r"		, $btnReload], _
		["n"		, $btnScreen], _
		["{ENTER}"	, $btnBrowse] _
	]

	GUISetAccelerators($aAccelKeys)
	GUISetState()

	ShowPicInApp()

	While 1
		$msg = GUIGetMsg()

		Select
		Case $msg == $GUI_EVENT_CLOSE Or $msg == $btnMsgExit
			ExitLoop
		Case $msg == $btnOpen
			btnOpen()
		Case $msg == $btnBrowse
			btnBrowse()
		Case $msg == $btnDelete
			btnDelete()
		Case $msg == $btnReload
			btnReload()
		Case $msg == $btnBeauty
			btnBeauty()
		Case $msg == $btnPretty
			btnPretty()
		Case $msg == $btnSpecial
			btnSpecial()
		Case $msg == $btnScreen
			btnScreen()
			ExitLoop
		EndSelect
	WEnd

	GUIDelete()
	DestroyPicList()
	DestroySQL()
	_GDIPlus_ShutDown()

	If $msg == $btnScreen Then
		Main()
	EndIf
EndFunc

Func InitPath()
	$sPicMiscDir = IniRead($appini, "Global", "sPicMiscDir", "C:\Users\tsaikd\Desktop\picture\misc")
	$sPicRenameBeautyDir = IniRead($appini, "Global", "sPicRenameBeautyDir", "C:\Users\tsaikd\Desktop\picture\new\beauty")
	$sPicRenamePrettyDir = IniRead($appini, "Global", "sPicRenamePrettyDir", "C:\Users\tsaikd\Desktop\picture\new\pretty")
	$sPicRenameSpecialDir = IniRead($appini, "Global", "sPicRenameSpecialDir", "C:\Users\tsaikd\Desktop\picture\new\special")
	$sOpenAppPath = IniRead($appini, "Global", "sOpenAppPath", "C:\Users\tsaikd\Desktop\DTP\Imagine_1.0.9_x64_Unicode\Imagine64.exe")

	If Not FileExists($sPicMiscDir) Then
		MsgBox(0x10, $app, _("Misc Picture Directory no found"))
		If Not SetDirPath("browse") Then Return False
	EndIf
	If Not FileExists($sPicRenameBeautyDir) Then
		MsgBox(0x10, $app, _("Beauty Picture Directory no found"))
		If Not SetDirPath("beauty") Then Return False
	EndIf
	If Not FileExists($sPicRenamePrettyDir) Then
		MsgBox(0x10, $app, _("Pretty Picture Directory no found"))
		If Not SetDirPath("pretty") Then Return False
	EndIf
	If Not FileExists($sPicRenameSpecialDir) Then
		MsgBox(0x10, $app, _("Special Picture Directory no found"))
		If Not SetDirPath("special") Then Return False
	EndIf

	Return True
EndFunc

Func GUIUpdatePath()
	GUICtrlSetTip($btnBrowse, StringFormat(_("Press [Shift] can set the path\nNow: %s"), $sPicMiscDir))
	GUICtrlSetTip($btnBeauty, StringFormat(_("Press [Shift] can set the path\nNow: %s"), $sPicRenameBeautyDir))
	GUICtrlSetTip($btnPretty, StringFormat(_("Press [Shift] can set the path\nNow: %s"), $sPicRenamePrettyDir))
	GUICtrlSetTip($btnSpecial, StringFormat(_("Press [Shift] can set the path\nNow: %s"), $sPicRenameSpecialDir))
EndFunc

Func InitSQL()
	_SQLite_Startup()
	If @error > 0 Then
		MsgBox(0x10, $app, "SQLite.dll Can't be Loaded!")
		Return False
	EndIf

	$sql = _SQLite_Open($appsql)
	If @error > 0 Then
		MsgBox(0x10, $app, "Can't Load Database!")
		Return False
	EndIf

	Local $cmd = _
		'CREATE TABLE IF NOT EXISTS DirMap' & _
		'	( i INTEGER PRIMARY KEY AUTOINCREMENT' & _
		'	, s TEXT UNIQUE' & _
		');' & @CRLF & _
		'' & _
		'CREATE TABLE IF NOT EXISTS History'& _
		'	( sName TEXT' & _
		'	, iSrcDir INTEGER' & _
		'	, iTarDir INTEGER' & _
		'	, iTime INTEGER' & _
		');' & @CRLF & _
		''

	_SQLite_Exec($sql, $cmd)
	If @error <> 0 Then
		MsgBox(0x10, $app, "Create SQL table failed!")
		Return False
	EndIf

	Return True
EndFunc

Func DestroySQL()
	_SQLite_Close($sql)
	_SQLite_Shutdown()
EndFunc

Func SQLEscapePath($path)
	If StringRight($path, 1) == "\" Then $path = StringTrimRight($path, 1)
	Return $path
EndFunc

; if not found, will insert into database and return index of it
Func SQLGetDirIndex($path)
	Local $aBuf, $ret
	$path = SQLEscapePath($path)

	_SQLite_QuerySingleRow($sql, 'SELECT i FROM DirMap WHERE s = "'&$path&'";', $aBuf)
	If $aBuf[0] == "" Then
		SQLAddDirMap($path)
		Return SQLGetDirIndex($path)
	Else
		$ret = $aBuf[0]
	EndIf

	Return $ret
EndFunc

Func SQLAddDirMap($path)
	$path = SQLEscapePath($path)
	Local $cmd = _
		'INSERT INTO DirMap ( s ) VALUES ( "'&$path&'" );' & _
		''
	_SQLite_Exec($sql, $cmd)
	If @error <> 0 Then
		MsgBox(0x10, $app, "SQLAddDirMap failed!")
		Return False
	EndIf
EndFunc

Func SQLAddHistory($srcpath, $tardir)
	$srcpath = SQLEscapePath($srcpath)
	$tardir = SQLEscapePath($tardir)
	Local $a, $b, $c, $d
	_PathSplit($srcpath, $a, $b, $c, $d)
	SQLGetDirIndex($a&$b)
	Local $cmd = '' & _
		'INSERT INTO History' & _
		'	( sName' & _
		'	, iSrcDir' & _
		'	, iTarDir' & _
		'	, iTime' & _
		') VALUES' & _
		'	( "'&$c&$d&'"' & _
		'	, '&SQLGetDirIndex($a&$b) & _
		'	, '&SQLGetDirIndex($tardir) & _
		'	, '&NowTime() & _
		');' & _
		''
	_SQLite_Exec($sql, $cmd)
	If @error <> 0 Then
		MsgBox(0x10, $app, "SQLAddHistory failed!")
		Return False
	EndIf
EndFunc

; InitPicList() must be after InitSQL() because of FileChangeDir()
Func InitPicList($path = "")
	DestroyPicList()
	If $path == "" Then $path = $sPicMiscDir

	FileChangeDir($path)
	$hPicList = FileFindFirstFile("*.*")
	If $hPicList == -1 Then
		If @WorkingDir == $sPicMiscDir Then
			Return False
		Else
			FileChangeDir("..")
			ConsoleWrite(DirRemove($path)&@CRLF)
			Return InitPicList()
		EndIf
	EndIf

	Local $tmp = DirGetSize($sPicMiscDir, 1)
	If IsArray($tmp) Then
		$iFileCount = $tmp[1]
	Else
		$iFileCount = 0
	EndIf

	Return True
EndFunc

Func DestroyPicList()
	If $hPicList <> -1 Then
		FileClose($hPicList)
		$hPicList = -1
	EndIf
EndFunc

Func PicListGetNextPicPath($bFindNextDir)
	Local $path
	Local $att

	While True
		If $hPicList == -1 Then Return ""

		$path = FileFindNextFile($hPicList)
		If @error Then
			If @WorkingDir == $sPicMiscDir Then
				Return ""
			ElseIf $bFindNextDir Then
				InitPicList()
				Return PicListGetNextPicPath($bFindNextDir)
			Else
				Return ""
			EndIf
		EndIf

		$att = FileGetAttrib($path)
		If @error Then Return ""

		If StringInStr($att, "D") Then
			If InitPicList(@WorkingDir&"\"&$path) Then
				Return PicListGetNextPicPath($bFindNextDir)
			Else
				MsgBox(0x10, $app, _("PicList Can't enter sub directory"))
				InitPicList()
				Return PicListGetNextPicPath($bFindNextDir)
			EndIf
		EndIf

		Return @WorkingDir&"\"&$path
	WEnd
EndFunc

Global $picList[5] = ["", "", "", "", ""]
Global $picIdx = -1
Func ShowPicInApp()
	Local $path
	While $picIdx < 4
		$path = PicListGetNextPicPath($picIdx == -1)
		If $path == "" Then
			ExitLoop
		EndIf
		$picIdx = $picIdx + 1
		$picList[$picIdx] = $path
	WEnd

	If $picList[0] == "" Then
		MsgBox(0x40, $app, _("Misc Picture Directory is Empty"))
		_IENavigate($ie, "about:blank")
		GUICtrlSetData($lblPicPath, "")
		Return
	EndIf

	$sCurPicPath = $picList[0]
	$picList[0] = $picList[1]
	$picList[1] = $picList[2]
	$picList[2] = $picList[3]
	$picList[3] = $picList[4]
	$picList[4] = ""
	$picIdx = $picIdx - 1

	Local $iPicMaxHeight = $iIEH - 45
	Local $hPic = _GDIPlus_ImageLoadFromFile($sCurPicPath)
	Local $picw, $pich
	Local $sSize
	If $hPic <> -1 Then
		$picw = _GDIPlus_ImageGetWidth($hPic)
		$pich = _GDIPlus_ImageGetHeight($hPic)
		If $picw == -1 Or $pich == -1 Then
			$hPic = -1
		Else
			$sSize = StringFormat("(%dx%d)\t", $picw, $pich)
		EndIf
	EndIf

	If $hPic == -1 Then
		MsgBox(0x10, $app, "GDIPlus load image failed: "&$sCurPicPath)
		_IENavigate($ie, $sCurPicPath)
	Else
		_GDIPlus_ImageDispose($hPic)
		$hPic = -1
		Local $extPicBody = ""
		If $picList[0] <> "" Then $extPicBody = $extPicBody & '<img class="ext" src="file://'&$picList[0]&'">'
		If $picList[1] <> "" Then $extPicBody = $extPicBody & '<img class="ext" src="file://'&$picList[1]&'">'
		If $picList[2] <> "" Then $extPicBody = $extPicBody & '<img class="ext" src="file://'&$picList[2]&'">'
		If $picList[3] <> "" Then $extPicBody = $extPicBody & '<img class="ext" src="file://'&$picList[3]&'">'
		_IEBodyWriteHTML($ie, '' & _
			'<body style="margin:0; padding: 0;">' & _
				'<style>' & _
					'img.ext {' & _
						'margin-left: 10px;' & _
						'height: expression(this.height > '&$iPicMaxHeight&' ? '&$iPicMaxHeight&' : true);' & _
					'}' & _
				'</style>' & _
				'<div style="white-space: nowrap;">' & _
					'<img height="'&$iPicMaxHeight&'"' & _
					' style="margin-right: 40px;"' & _
					' src="file://'&$sCurPicPath&'"' & _
					' onclick="javascript: ' & _
						'var h1 = '&$pich&';' & _
						'var h2 = '&$iPicMaxHeight&';' & _
						'if (this.clicktag) {' & _
						'    this.height = h2;' & _
						'} else {' & _
						'    this.height = h1;' & _
						'}' & _
						'this.clicktag = !this.clicktag;' & _
					'">' & _
					$extPicBody & _
				'</div>' & _
			'</body>' & _
			'')
	EndIf

	GUICtrlSetData($lblPicPath, $sSize&$sCurPicPath)
	$iFileCount -= 1
	GUICtrlSetData($lblFileCount, $iFileCount)
EndFunc

Func MovePic($srcpath, $tardir, $bSplash = True)
	If $bSplash Then SplashMsgBegin(_("Moving picture"))

	If 1 == FileMove($srcpath, $tardir) Then
		SQLAddHistory($srcpath, $tardir)
		Return SplashMsgEnd(True)
	Else
		SplashMsgEnd()
		MsgBox(0x10, $app, _("Move Picture failed"))
		Return False
	EndIf
EndFunc

Func SplashMsgBegin($msg)
	SplashTextOn($app, $msg, 300, 120, -1, -1, 0x30)
EndFunc

#cs
	@param $ret return value, if Default then return nothing
#ce
Func SplashMsgEnd($ret = Default)
	SplashOff()
	If $ret == Default Then
		Return
	Else
		Return $ret
	EndIf
EndFunc

#cs
$name can be
	"browse"
	"beauty"
	"pretty"
	"special"

@retval true if set new directory
@retval false if cancel or invalid parameters
#ce
Func SetDirPath($name)
	Local $path
	Local $key
	Local $prompt = StringFormat(_("Please select %s directory path"), $name)

	Switch($name)
	Case "browse"
		$key = "sPicMiscDir"
	Case "beauty"
		$key = "sPicRenameBeautyDir"
	Case "pretty"
		$key = "sPicRenamePrettyDir"
	Case "special"
		$key = "sPicRenameSpecialDir"
	Case Else
		MsgBox(0x40, $app, _("SetDirPath(): Invalid parameter: $name"))
		Return False
	EndSwitch

	$path = FileSelectFolder($prompt, "", 0x03, Execute("$"&$key))
	If @error == 1 Then Return False

	If FileExists($path) Then
		IniWrite($appini, "Global", $key, $path)
	Else
		IniDelete($appini, "Global", $key)
	EndIf
	InitPath()
	GUIUpdatePath()
	Return True
EndFunc

Func btnOpen()
	ShellExecute($sOpenAppPath, $sCurPicPath)
EndFunc

Func btnBrowse()
	Local $a, $b, $c, $d

	If _IsPressed("10") Then Return SetDirPath("browse")

	$path = _PathSplit($sCurPicPath, $a, $b, $c, $d)
	ShellExecute($a&$b)
EndFunc

Func btnDelete()
	SplashMsgBegin(_("Delete file ..."))
	If FileRecycle($sCurPicPath) Then
		SQLAddHistory($sCurPicPath, "Trash")
	Else
		If FileExists($sCurPicPath) Then
			MsgBox(0x10, $app, _("Delete file failed"))
			Return SplashMsgEnd(False)
		EndIf
	EndIf
	ShowPicInApp()
	Return SplashMsgEnd(True)
EndFunc

Func btnReload()
	SplashMsgBegin(_("Reloading ..."))
	$picList[0] = $picList[1] = $picList[2] = $picList[3] = $picList[4] = ""
	$picIdx = -1
	InitPicList()
	ShowPicInApp()
	SplashMsgEnd()
EndFunc

Func btnBeauty()
	If _IsPressed("10") Then Return SetDirPath("beauty")

	If MovePic($sCurPicPath, $sPicRenameBeautyDir) Then
		ShowPicInApp()
	EndIf
EndFunc

Func btnPretty()
	If _IsPressed("10") Then Return SetDirPath("pretty")

	If MovePic($sCurPicPath, $sPicRenamePrettyDir) Then
		ShowPicInApp()
	EndIf
EndFunc

Func btnSpecial()
	If _IsPressed("10") Then Return SetDirPath("special")

	If MovePic($sCurPicPath, $sPicRenameSpecialDir) Then
		ShowPicInApp()
	EndIf
EndFunc

Func btnScreen()
	$picList[0] = $picList[1] = $picList[2] = $picList[3] = $picList[4] = ""
	$picIdx = -1
	Dim $screenInfo = _NumberAndNameMonitors()
	If $screenInfo[0][0] > 1 Then
		Dim $iScreen = IniRead($appini, "Global", "iScreen", 2)
		If $iScreen == 1 Then
			IniWrite($appini, "Global", "iScreen", 2)
		Else
			IniWrite($appini, "Global", "iScreen", 1)
		EndIf
	EndIf
EndFunc

; unit of offset is second
Func NowTime($offset=0)
	Local $ret
	$offset = Int($offset)
	If $offset == 0 Then
		$ret = Number(@YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC)
	Else
		$ret = _DateAdd("s", $offset, StringFormat("%s/%s/%s %s:%s:%s", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
		$ret = StringReplace($ret, "/", "")
		$ret = StringReplace($ret, " ", "")
		$ret = StringReplace($ret, ":", "")
	EndIf
	Return $ret
EndFunc

Func _($s)
	Switch($s)
	Case "Open(&O)"
		Return "開啟檔案(&O)"
	Case "Browse"
		Return "瀏覽目錄"
	Case "Delete(&D)"
		Return "刪除檔案(&D)"
	Case "Reload(&R)"
		Return "重新載入(&R)"
	Case "Press [Shift] can set the path\nNow: %s"
		Return "按住 [Shift] 可以設定路徑\n目前設定: %s"
	Case "Please select %s directory path"
		Return "請選擇 %s 的資料夾"
	EndSwitch
	Return $s
EndFunc

Main()
