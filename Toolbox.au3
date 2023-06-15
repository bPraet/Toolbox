#pragma compile(FileVersion, 0.0.1.4)

#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#Include <GuiListBox.au3>
#include <SendMessage.au3>
#include <AD.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <SliderConstants.au3>
#include <String.au3>
#include <StringCompareVersions.au3>
#include <FileConstants.au3>
#include <Excel.au3>
#include <InetConstants.au3>

DllCall("kernel32.dll", "int", "Wow64DisableWow64FsRedirection", "int", 1)

Call("selfUpdate")

; GUI
$Main = GUICreate("Toolbox", 497, 332, 470, 185)
; WinSetOnTop($Main, "", 1)

;$Bg = GUICtrlCreatePic("C:\Users\bpraea\Pictures\bg.jpg", 0, 0, 497, 332)
;GUICtrlSetState(-1, $GUI_DISABLE)

$AntiAFK = GUICtrlCreateSlider(193, 168, 57, 30)
GUICtrlSetLimit(-1, 1, 0)

$Separator = GUICtrlCreateGroup("", 264, 16, 1, 295)

$Login = GUICtrlCreateInput("Login", 16, 40, 233, 21)
$Password = GUICtrlCreateInput("Password", 16, 96, 233, 21, $ES_PASSWORD)

$Interval = GUICtrlCreateInput("30", 16, 168, 145, 21)

$User = GUICtrlCreateInput("", 16, 240, 97, 21)
$UserAction = GUICtrlCreateCombo("Search", 120, 240, 97, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData($UserAction, "Search", "Search")
$GoUser = GUICtrlCreateButton("->", 224, 240, 25, 25)
$Pc = GUICtrlCreateInput("", 16, 296, 97, 21)
$PcAction = GUICtrlCreateCombo("", 120, 296, 97, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData($PcAction, "Search|Explorer|TeamViewer|Zoho Install|Zoho Uninstall", "Search")
$GoPc = GUICtrlCreateButton("->", 224, 296, 25, 25)

$AnswerList = GUICtrlCreateList("", 280, 40, 193, 201, BitOR($WS_BORDER, $WS_VSCROLL))
$AddBtn = GUICtrlCreateButton("Add", 280, 256, 89, 33)
$RemoveBtn = GUICtrlCreateButton("Remove", 384, 256, 89, 33)

$Name = GUICtrlCreateInput("", 280, 296, 193, 21)

$Label1 = GUICtrlCreateLabel("Login", 16, 16, 37, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Password", 16, 72, 64, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("Interval", 16, 144, 47, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("User", 16, 216, 33, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Label5 = GUICtrlCreateLabel("PC", 16, 272, 22, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Label6 = GUICtrlCreateLabel("Answers", 280, 16, 55, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")

GUISetState(@SW_SHOW)
; End GUI

Global Const $SC_DRAGMOVE = 0xF012

; Shortcuts
HotKeySet("{ESC}", "On_Exit")
HotKeySet("$", "autoInput")
HotKeySet("µ", "login")
HotKeySet("²", "answerPaste")

; Config File
$tbConfigPath = @ScriptDir & "\tbConfig.xlsx"

; Excel
if FileExists($tbConfigPath) then
   msgbox(0, "File", "Config file loaded")
else
   $oExcel = _Excel_Open()
   $oWorkbook = _Excel_BookNew($oExcel, 2)
   _Excel_BookSaveAs($oWorkbook, $tbConfigPath, $xlWorkbookDefault, True)
   msgbox(0, "File", "Config file created")
endif

$oExcel = _Excel_Open()
$oWorkbook = _Excel_BookOpen($oExcel, $tbConfigPath)
$answer = "answer"
$i = 1
$AnswerListTemp = ""

Call("answersRefresh")

; Events
While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_PRIMARYDOWN
             _SendMessage($Main, $WM_SYSCOMMAND, $SC_DRAGMOVE, 0)
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $AntiAFK
		 While GUICtrlRead($AntiAFK) = 1
			$i = Number(GUICtrlRead($Interval)) * 1000
			MouseMove(@DesktopWidth/2 - 150, @DesktopHeight/2)
			MouseClick("")
			Sleep(2000)
			MouseMove(@DesktopWidth/2 + 150, @DesktopHeight/2)
			MouseClick("")
			While GUICtrlRead($AntiAFK) = 1 And $i > 0
			   Sleep(1000)
			   $i -= 1000
			WEnd
		 WEnd
	  Case $AddBtn
		 Call("addAnswer")
	  Case $RemoveBtn
		 Call("delAnswer")
	  Case $GoPc
		 Call("adPcAction")
	  Case $GoUser
		 Call("adUserAction")
   EndSwitch
WEnd

Func On_Exit()
   Exit
EndFunc

Func autoInput()
   Send(GUICtrlRead($Login), 1)
   Send("{TAB}")
   Send(GUICtrlRead($Password), 1)
EndFunc

Func adUserAction()
   _AD_Open(GUICtrlRead($Login), GUICtrlRead($Password))

   If @error <> 0 Then
	  MsgBox(16, "Error", "Can't connect to AD")
	  Return
   EndIf

   Switch GUICtrlRead($UserAction)
	  Case "Search"
		 $userFound = _AD_GetObjectProperties(GUICtrlRead($user), "")

		 If @error <> 0 Then MsgBox(16, "Error", "No results found")
		 _AD_Close()

		 _ArrayDisplay($userFound)
   EndSwitch
EndFunc

Func adPcAction()
   Switch GUICtrlRead($PcAction)
	  Case "Search"
		 _AD_Open(GUICtrlRead($Login), GUICtrlRead($Password))

		 If @error <> 0 Then
			MsgBox(16, "Error", "Can't connect to AD")
			Return
		 EndIf

		 $pcFound = _AD_GetObjectProperties(GUICtrlRead($pc) & "$", "")

		 If @error <> 0 Then MsgBox(16, "Error", "No results found")
		 _AD_Close()

		 $pingResult = Ping(GUICtrlRead($pc))

		 If $pingResult Then
			MsgBox($MB_SYSTEMMODAL, "", "IP: "& TCPNameToIP(GUICtrlRead($pc)) &" Temps de réponse: " & $pingResult & "ms")
		 Else
			MsgBox($MB_SYSTEMMODAL, "", "No ping")
		 EndIf

		 _ArrayDisplay($pcFound)
	  Case "Explorer"
		 $pingResult = Ping(GUICtrlRead($pc))

		 If $pingResult Then
			Run("Explorer \\"& GUICtrlRead($pc) &"\c$")
		 Else
			MsgBox($MB_SYSTEMMODAL, "", "No ping")
		 EndIf
	  Case "TeamViewer"
		 Run("C:\Program Files (x86)\TeamViewer\Version9\TeamViewer.exe --id "& GUICtrlRead($pc))
	  Case "Zoho Install"
		 Call("zohoInstall")
	  Case "Zoho Uninstall"
		 Call("zohoUninstall")
	  Case Else
		 $sMsg = "No action selected"
   EndSwitch
EndFunc

Func login()
   _AD_Open(GUICtrlRead($Login), GUICtrlRead($Password))

   If @error <> 0 Then
	  MsgBox(16, "Error", "Can't connect to AD")
	  Return
   EndIf

   $givenName = _AD_GetObjectAttribute(@username, "givenName")

   If @error <> 0 Then MsgBox(16, "Error", "No results found")
   _AD_Close()

   GUICtrlSetData($Name, _StringProper($givenName))
EndFunc

Func answersRefresh()
   $i = 1
   $AnswerListTemp = ""

   While $answer <> ""
   $answer = _Excel_RangeRead($oWorkbook, Default, "A" & $i)
	  If $answer <> "" Then
		 $AnswerListTemp = $AnswerListTemp & $answer & "|"
		 $i += 1
	  EndIf
   WEnd

   $answer = "answer"
   GUICtrlSetData($AnswerList, $AnswerListTemp)
EndFunc

Func addAnswer()
   $answerName = InputBox("Name", "Insert answer name")
   $answerDescription = ClipGet()
   $copyAnswer = MsgBox(4, "Answer", "The following data is stored in the clipboard: " & @CRLF & $answerDescription)
   ConsoleWrite($copyAnswer)
   If $copyAnswer == 6 Then
	  $AnswerListTemp = $AnswerListTemp & $answerName & "|"

	  _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $answerName,"A" & $i)
	  _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $answerDescription,"B" & $i)
	  _Excel_BookSave ( $oWorkbook )

	  $i += 1
	  GUICtrlSetData($AnswerList, "")
	  Call("answersRefresh")
   EndIf

EndFunc

Func delAnswer()
   $answerIndex = _GUICtrlListBox_GetCurSel($answerList)

   _GUICtrlListBox_DeleteString($answerList, $answerIndex)
   $answerIndex += 1
   $i -= 1
   ConsoleWrite($answerIndex)
   _Excel_RangeDelete($oWorkbook.activesheet, $answerIndex & ":" & $answerIndex, $xlShiftUp, 1)
   _Excel_BookSave ( $oWorkbook )

EndFunc

Func answerPaste()
   $currentAnswerIndex = _GUICtrlListBox_GetCurSel ($answerList) + 1
   $currentAnswer = _Excel_RangeRead($oWorkbook, Default, "B" & $currentAnswerIndex, 1, True)

   If GUICtrlRead($Name) == "" Then
	  $currentAnswer = StringReplace($currentAnswer, "%name% ", "")
   Else
	  $currentAnswer = StringReplace($currentAnswer, "%name%", GUICtrlRead($Name))
   EndIf

   ClipPut($currentAnswer)
EndFunc

Func zohoInstall()
   RunAsWait(GUICtrlRead($Login), "FINBEL", GUICtrlRead($Password), 0, @ComSpec & " /c " & "xcopy C:\tools\Zoho \\" & GUICtrlRead($Pc) & "\c$\temp\Zoho\", "", @SW_HIDE)
   ;FileCopy("C:\tools\Zoho", "\\" & GUICtrlRead($Pc) & "\c$\temp\Zoho\", $FC_OVERWRITE + $FC_CREATEPATH)
   RunAsWait(GUICtrlRead($Login), "FINBEL", GUICtrlRead($Password), 0, "psexec \\" & GUICtrlRead($Pc) & " -h -i msiexec.exe /i C:\temp\Zoho\ZA_Access.msi /qn", '', @SW_HIDE)

   If @error Then
      MsgBox(16, "Erreur", "Installation not completed")
      Exit
   EndIf
EndFunc

Func zohoUninstall()
   RunAsWait(GUICtrlRead($Login), "FINBEL", GUICtrlRead($Password), 0, "psexec \\" & GUICtrlRead($Pc) & " -h -i msiexec.exe /x C:\temp\Zoho\ZA_Access.msi /qn", '', @SW_HIDE)

   If @error Then
      MsgBox(16, "Erreur", "Uninstallation not completed")
      Exit
   EndIf

   RunAsWait(GUICtrlRead($Login), "FINBEL", GUICtrlRead($Password), 0, @ComSpec & " /c " & "rmdir /Q /S \\" & GUICtrlRead($Pc) & "\c$\temp\Zoho\", "", @SW_HIDE)
   ;DirRemove("\\" & GUICtrlRead($Pc) & "\c$\temp\Zoho\", $DIR_REMOVE)
EndFunc

Func selfUpdate()
   If Not IsDeclared("updatePath") Then
	  InetGet("https://github.com/bPraet/Toolbox/releases/download/toolbox/Toolbox.exe", "update.exe")

	  Global $path = FileGetShortName(@ScriptDir & "\update.exe")
	  If Not FileExists($path) Then
		 MsgBox(48, "Update Error!", "Unable to locate an update file!")
	  Else
		 $updatePath = $path
	  EndIf

	  $newVersion = FileGetVersion($updatePath)
	  ConsoleWrite($newVersion)
	  $oldVersion = FileGetVersion(@ScriptFullPath)
	  ConsoleWrite($oldVersion)
	  $results = _StringCompareVersions($oldVersion, $newVersion)
	  If $results = "-1" Then
		 SplashTextOn("Updating Loader", "Please wait... This may take a few moments.", "400", "100", "-1", "-1", 50, "", "", "")
		 Sleep(1000)
		 FileCopy($updatePath, @ScriptFullPath & ".new")
		 Local $batchPath = @ScriptDir & '\update.bat'
		 $fileData = "@echo off" & @CRLF & _
		 "ping localhost -n 2 > nul" & @CRLF & _
		 ":loop" & @CRLF & _
		 'del /Q "' & @ScriptFullPath & '"' & @CRLF & _
		 'if exist "' & @ScriptFullPath & '" goto loop' & @CRLF & _
		 'move "' & @ScriptFullPath & '.new" "' & @ScriptFullPath & '"' & @CRLF & _
		 'start /B "Loading" "' & @ScriptFullPath & '"' & @CRLF & _
		 'del /Q "' & $batchPath & '"' & @CRLF & _
		 "exit"
		 FileWrite($batchPath, $fileData)
		 Run($batchPath, "", @SW_HIDE)
		 SplashOff()
		 Exit
	  EndIf
	  FileDelete("update.exe")
   EndIf
EndFunc
