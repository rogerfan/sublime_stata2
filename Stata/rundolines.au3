; AutoIt v3 script to run Stata commands from an external text editor
; Version 4.0, 14 August 2011
; Friedrich Huebler, fhuebler@gmail.com, www.huebler.info

; Declare variables
Global $ini, $statapath, $statawin, $statacmd, $clippause, $winpause, $keypause, $commands, $tempfile, $tempfile2

; File locations
; Path to INI file
$ini = @ScriptDir & "\rundolines.ini"
; Path to Stata executable
$statapath = IniRead($ini, "Stata", "statapath", "C:\Program Files\Stata11\StataSE.exe")
; Title of Stata window
$statawin = IniRead($ini, "Stata", "statawin", "Stata/SE 11.2")

; Keyboard shortcut for Stata command window
$statacmd = IniRead($ini, "Stata", "statacmd", "^1")

; Delays
; Pause after copying of Stata commands to clipboard
$clippause = IniRead($ini, "Delays", "clippause", "100")
; Pause between window-related operations
$winpause = IniRead($ini, "Delays", "winpause", "200")
; Pause between keystrokes sent to Stata
$keypause = IniRead($ini, "Delays", "keypause", "1")

; Set WinWaitDelay and SendKeyDelay to speed up or slow down script
Opt("WinWaitDelay", $winpause)
Opt("SendKeyDelay", $keypause)

; If more than one Stata window is open, the window that was most recently active will be matched
Opt("WinTitleMatchMode", 2)

; Clear clipboard
ClipPut("")
; Copy selected lines from editor to clipboard
Send("^c")
; Pause avoids problem with clipboard, may be AutoIt or Windows bug
Sleep($clippause)
$commands = ClipGet()

; Terminate script if no commands selected in editor
If $commands = "" Then 
  Exit
EndIf

; Create file name in system temporary directory
$tempfile = EnvGet("TEMP") & "\statacmd.tmp"

; Open file for writing and check that it worked
$tempfile2 = FileOpen($tempfile, 2)
If $tempfile2 = -1 Then
  MsgBox(0, "Error: Cannot open temporary file", "at [" & $tempfile & "]")
  Exit
EndIf

; Write commands to temporary file, add CR-LF at end to ensure last line is executed by Stata
FileWrite($tempfile2, $commands & @CRLF)
FileClose($tempfile2)

; Check if Stata is already open, start Stata if not
If WinExists($statawin) Then
  WinActivate($statawin)
  WinWaitActive($statawin)
  ; Activate Stata command window and select text (if any)
  Send($statacmd)
  Send("^a")
  ; Run temporary file
  ; Double quotes around $tempfile needed in case path contains blanks
  ClipPut("do " & '"' & $tempfile & '"')
  ; Pause avoids problem with clipboard, may be AutoIt or Windows bug
  Sleep($clippause)
  Send("^v" & "{Enter}")
Else
  Run($statapath)
  WinWaitActive($statawin)
  ; Activate Stata command window
  Send($statacmd)
  ; Run temporary file
  ; Double quotes around $dofile needed in case path contains blanks
  ClipPut("do " & '"' & $tempfile & '"')
  ; Pause avoids problem with clipboard, may be AutoIt or Windows bug
  Sleep($clippause)
  Send("^v" & "{Enter}")
EndIf

; End of script
