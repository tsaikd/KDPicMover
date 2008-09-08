#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile=KDPicMover.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Language=1028
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs

Changelog:
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

; Variable Definition
Global Const $appname = "KDPicMover"
Global Const $appver = "1.0.0.1"
Global Const $appdate = "2008/09/07"
Global Const $author = "tsaikd@gmail.com"

Global Const $appsql = $appname&".sqlite"

; Initialization
Global $ie, $ieActiveX, $iIEW, $iIEH

Global $sql
Global $appgui
Global $hWndHotKey
Global Const $app = $appname&" "&$appver

Global $hPicList = -1
Global $sCurPicPath

Global $lblPicPath

$appwidth = 1000
$appheight = 700
$sHomeUrl = "C:\Documents and Settings\tsaikd\桌面\picture\local\name\Fanni\w_DCseries_Fanni_20040313_001.jpg"
$sPicMiscDir = "C:\Documents and Settings\tsaikd\桌面\picture\misc"
$sPicRenameBeautyDir = "C:\Documents and Settings\tsaikd\桌面\picture\new\beauty"
$sPicRenamePrettyDir = "C:\Documents and Settings\tsaikd\桌面\picture\new\pretty"

#cs
Pleace set AutoItWinSetTitle($appname) first
Or return value will always be false
#ce
Func IsAppActive()
	Return BitAND(WinGetState(AutoItWinGetTitle()), 8) == 8
EndFunc

Func Main()
	If Not FileExists($sPicMiscDir) Then Return MsgBox(0x10, $app, _("Misc Picture Directory no found"))
	If Not FileExists($sPicRenameBeautyDir) Then Return MsgBox(0x10, $app, _("Beauty Picture Directory no found"))
	If Not FileExists($sPicRenamePrettyDir) Then Return MsgBox(0x10, $app, _("Pretty Picture Directory no found"))

	_GDIPlus_Startup()
	_IEErrorHandlerRegister()
	If Not InitSQL() Then Return
	If Not InitPicList() Then Return

	$ie = _IECreateEmbedded()
	$appgui = GUICreate($appname, $appwidth, $appheight)
	AutoItWinSetTitle($appname)

	$iWinBH = 5
	$iCtrlGap = 10
	$iCtrlH = 30
	$iLblH = 14
	$iLblO = ($iCtrlH-$iLblH)/2
	$iBtnH = 25
	$iBtnW = 100
	$iBtnG = 150
	$iIEW = $appwidth-$iCtrlGap*2
	$iIEH = $appheight-$iCtrlH*2-$iCtrlGap-$iWinBH

	$ieActiveX = GUICtrlCreateObj($ie, $iCtrlGap, $iCtrlGap, $iIEW, $iIEH)
	$lblPicPath = GUICtrlCreateLabel("", $iCtrlGap, $appheight-$iCtrlH*2+$iLblO-$iWinBH, $appwidth-$iCtrlGap*2, $iLblH)
	$btnBeauty = GUICtrlCreateButton(_("&Beauty"), $appwidth/2-$iBtnW/2-$iBtnG*2.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnPretty = GUICtrlCreateButton(_("&Pretty"), $appwidth/2-$iBtnW/2-$iBtnG*1.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnOpen = GUICtrlCreateButton(_("&Open"), $appwidth/2-$iBtnW/2-$iBtnG*0.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnBrowse = GUICtrlCreateButton(_("Browse"), $appwidth/2-$iBtnW/2+$iBtnG*0.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnDelete = GUICtrlCreateButton(_("&Delete"), $appwidth/2-$iBtnW/2+$iBtnG*1.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)
	$btnReload = GUICtrlCreateButton(_("Reload"), $appwidth/2-$iBtnW/2+$iBtnG*2.5, $appheight-$iCtrlH-$iWinBH, $iBtnW, $iBtnH)

	$btnMsgExit = GUICtrlCreateButton("Exit", 0, 0)
	GUICtrlSetState(-1, $GUI_HIDE)
	_IENavigate($ie, "about:blank")

	Dim $aAccelKeys[6][2] = [ _
		["{ESC}", $btnMsgExit], _
		["b", $btnBeauty], _
		["d", $btnDelete], _
		["o", $btnOpen], _
		["{ENTER}", $btnBrowse], _
		["p", $btnPretty] _
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
		EndSelect
	WEnd

	GUIDelete()
	DestroyPicList()
	DestroySQL()
	_GDIPlus_ShutDown()
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
	If $hPicList <> -1 Then DestroyPicList()
	If $path == "" Then $path = $sPicMiscDir

	FileChangeDir($path)
	$hPicList = FileFindFirstFile("*.*")
	If $hPicList == -1 Then
		If @WorkingDir == $sPicMiscDir Then
			Return False
		Else
			FileChangeDir("..")
			DirRemove($path)
			Return InitPicList()
		EndIf
	EndIf
	Return True
EndFunc

Func DestroyPicList()
	If $hPicList <> -1 Then
		FileClose($hPicList)
		$hPicList = -1
	EndIf
EndFunc

Func PicListGetNextPicPath()
	If $hPicList == -1 Then Return ""

	Local $path = FileFindNextFile($hPicList)
	If @error Then
		If @WorkingDir == $sPicMiscDir Then
			Return ""
		Else
			InitPicList()
			Return PicListGetNextPicPath()
		EndIf
	EndIf

	Local $att = FileGetAttrib($path)
	If @error Then Return ""

	If StringInStr($att, "D") Then
		If InitPicList(@WorkingDir&"\"&$path) Then
			Return PicListGetNextPicPath()
		Else
			MsgBox(0x10, $app, _("PicList Can't enter sub directory"))
			InitPicList()
			Return PicListGetNextPicPath()
		EndIf
	EndIf

	Return @WorkingDir&"\"&$path
EndFunc

Func ShowPicInApp()
	Local $path = PicListGetNextPicPath()
	If $path == "" Then
		MsgBox(0x40, $app, _("Misc Picture Directory is Empty"))
		Return
	EndIf

	$sCurPicPath = $path
	Local $hPic = _GDIPlus_ImageLoadFromFile($sCurPicPath)
	Local $picw, $pich
	Local $npicw, $npich
	Local $sSize
	If $hPic <> -1 Then
		$picw = _GDIPlus_ImageGetWidth($hPic)
		$pich = _GDIPlus_ImageGetHeight($hPic)
		If $picw == -1 Or $pich == -1 Then
			$hPic = -1
		Else
			$sSize = StringFormat("(%dx%d)\t", $picw, $pich)
			Local $fw = ($iIEW-10) / $picw
			Local $fh = ($iIEH-35) / $pich
			Local $fmin = _Min($fw, $fh)
			If $fmin < 1 Then
				$npicw = Int($picw * $fmin)
				$npich = Int($pich * $fmin)
			Else
				$npicw = $picw
				$npich = $pich
			EndIf
		EndIf
	EndIf

	If $hPic == -1 Then
		_IENavigate($ie, $sCurPicPath)
	Else
		_GDIPlus_ImageDispose($hPic)
		$hPic = -1
		_IEBodyWriteHTML($ie, '' & _
			'<body style="margin:0; padding: 0;">' & _
				'<img width="'&$npicw&'" height="'&$npich&'"' & _
				' style="margin: 0; padding: 0;"' & _
				' src="file://'&$sCurPicPath&'"' & _
				' onclick="javascript: ' & _
					'var w1 = '&$picw&';' & _
					'var h1 = '&$pich&';' & _
					'var w2 = '&$npicw&';' & _
					'var h2 = '&$npich&';' & _
					'if (this.width == w1) {' & _
					'    this.width = w2;' & _
					'    this.height = h2;' & _
					'} else {' & _
					'    this.width = w1;' & _
					'    this.height = h1;' & _
					'}' & _
				'">' & _
			'</body>' & _
			'')
	EndIf

	GUICtrlSetData($lblPicPath, $sSize&$sCurPicPath)
EndFunc

Func MovePic($srcpath, $tardir)
	If 1 == FileMove($srcpath, $tardir) Then
		SQLAddHistory($srcpath, $tardir)
		Return True
	Else
		MsgBox(0x10, $app, _("Move Picture failed"))
		Return False
	EndIf
EndFunc

Func SplashMsgBegin($msg)
	SplashTextOn($app, $msg, 300, 120, -1, -1, 0x30)
EndFunc

Func SplashMsgEnd($ret = Default)
	SplashOff()
	If $ret == Default Then
		Return
	Else
		Return $ret
	EndIf
EndFunc

Func btnOpen()
	ShellExecute($sCurPicPath)
EndFunc

Func btnBrowse()
	Local $a, $b, $c, $d
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
	InitPicList()
	ShowPicInApp()
	SplashMsgEnd()
EndFunc

Func btnBeauty()
	SplashMsgBegin(_("Moving picture to beauty"))
	If MovePic($sCurPicPath, $sPicRenameBeautyDir) Then
		ShowPicInApp()
	EndIf
	SplashMsgEnd()
EndFunc

Func btnPretty()
	SplashMsgBegin(_("Moving picture to pretty"))
	If MovePic($sCurPicPath, $sPicRenamePrettyDir) Then
		ShowPicInApp()
	EndIf
	SplashMsgEnd()
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
	Case "&Open"
		Return "開啟檔案(&O)"
	Case "Browse"
		Return "瀏覽目錄"
	Case "&Delete"
		Return "刪除檔案(&D)"
	Case "Reload"
		Return "重新載入"
	EndSwitch
	Return $s
EndFunc

Main()
