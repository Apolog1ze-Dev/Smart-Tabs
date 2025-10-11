; AutoIt Helper for Smart Tabs - Standalone Executable
; This script will be compiled to a standalone .exe that requires no AutoIt installation
; Handles file selection detection and shortcut creation triggers

#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\build\icon.ico
#AutoIt3Wrapper_Outfile=autoit-helper.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Smart Tabs File Detection Helper
#AutoIt3Wrapper_Res_Description=Background helper for Smart Tabs file detection and shortcut creation
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductName=Smart Tabs File Helper
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_CompanyName=Smart Tabs
#AutoIt3Wrapper_Res_LegalCopyright=Copyright Smart Tabs
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; String constants (instead of including StringConstants.au3)
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Global Const $STR_NOCOUNT = 2

; File constants (instead of including File.au3)
Global Const $FO_APPEND = 8
Global Const $FO_CREATEPATH = 16

; Global variables for communication
Global $g_sOutputFile = @TempDir & "\smarttabs-autoit-output.txt"
Global $g_sCommandFile = @TempDir & "\smarttabs-autoit-commands.txt"
Global $g_bRunning = True
Global $g_sRegisteredHotkey = ""
Global $g_sOriginalClipboard = "" ; Store original clipboard for restoration

; Initialize and start main loop
Initialize()
MainLoop()

; Initialize the helper
Func Initialize()
    ; Clean up any existing files
    FileDelete($g_sOutputFile)
    FileDelete($g_sCommandFile)
    
    ; Send startup notification
    WriteOutput("HELPER_STARTED")
    
    ; Register default error handler
    OnAutoItExitRegister("OnExit")
EndFunc

; Main program loop
Func MainLoop()
    While $g_bRunning
        CheckForCommands()
        Sleep(100)
    WEnd
EndFunc

; Check for commands from Electron main process
Func CheckForCommands()
    If FileExists($g_sCommandFile) Then
        Local $sCommand = FileRead($g_sCommandFile)
        If $sCommand And StringLen($sCommand) > 0 Then
            FileDelete($g_sCommandFile)
            ProcessCommand(StringStripWS($sCommand, $STR_STRIPLEADING + $STR_STRIPTRAILING))
        EndIf
    EndIf
EndFunc

; Process commands from Electron
Func ProcessCommand($sCommand)
	Local $aParts = StringSplit($sCommand, "|", $STR_NOCOUNT)
	
	Switch $aParts[0]
		Case "REGISTER_SHORTCUT_HOTKEY"
			If UBound($aParts) > 1 Then
				RegisterShortcutHotkey($aParts[1])
			Else
				WriteOutput("ERROR|ERR_MISSING_PARAM|Missing hotkey parameter")
			EndIf
			
		Case "UNREGISTER_HOTKEY"
			UnregisterCurrentHotkey()

		Case "UNREGISTER_SHORTCUT_HOTKEY"
			UnregisterCurrentHotkey()
			WriteOutput("HOTKEY_UNREGISTERED")

		Case "EXIT"
			WriteOutput("SHUTTING_DOWN")
			$g_bRunning = False
			
		Case "PING"
			WriteOutput("PONG")
			
		Case "GET_STATUS"
			WriteOutput("STATUS|RUNNING|" & $g_sRegisteredHotkey)
			
		Case Else
			WriteOutput("ERROR|ERR_UNKNOWN_COMMAND|Unknown command: " & $aParts[0])
	EndSwitch
EndFunc

; Register the shortcut creation hotkey
Func RegisterShortcutHotkey($sHotkey)
    ; Unregister any existing hotkey first
    UnregisterCurrentHotkey()
    
    ; Convert Electron format to AutoIt format
    Local $sAutoItHotkey = ConvertHotkeyFormat($sHotkey)
    
	If $sAutoItHotkey Then
		; Register the new hotkey
		HotKeySet($sAutoItHotkey, "HandleShortcutCreation")
		$g_sRegisteredHotkey = $sHotkey
		WriteOutput("HOTKEY_REGISTERED|" & $sHotkey)
	Else
		WriteOutput("ERROR|ERR_INVALID_HOTKEY|Invalid hotkey format: " & $sHotkey)
	EndIf
EndFunc

; Unregister the current hotkey
Func UnregisterCurrentHotkey()
    If $g_sRegisteredHotkey Then
        Local $sAutoItHotkey = ConvertHotkeyFormat($g_sRegisteredHotkey)
        If $sAutoItHotkey Then
            HotKeySet($sAutoItHotkey)
        EndIf
        WriteOutput("HOTKEY_UNREGISTERED|" & $g_sRegisteredHotkey)
        $g_sRegisteredHotkey = ""
    EndIf
EndFunc

; Handle the shortcut creation hotkey press
Func HandleShortcutCreation()
    WriteOutput("HOTKEY_TRIGGERED")
    
    Local $selectedData = GetSelectedContent()
    If $selectedData Then
        WriteOutput("SHORTCUT_DATA|" & $selectedData)
    Else
        WriteOutput("NO_SELECTION_ERROR")
    EndIf
EndFunc

; Main function to get selected content from various sources
; PRIORITY ORDER: Text/URLs FIRST, then files, then folders
Func GetSelectedContent()
    Local $result = ""
    
    WriteOutput("DEBUG|Starting content detection (URL priority mode)...")
    
    ; Method 1: Get selected text (HIGHEST PRIORITY - URLs and file paths)
    WriteOutput("DEBUG|Method 1: Checking for selected text/URLs...")
    $result = GetSelectedText()
    If $result Then 
        WriteOutput("DEBUG|✓ Found and processed selected text (URL/path)")
        Return $result
    EndIf
    WriteOutput("DEBUG|✗ No valid selected text found")
    
    ; Method 2: Get selected files from Windows Explorer
    WriteOutput("DEBUG|Method 2: Checking for selected files in Explorer...")
    $result = GetSelectedFilesFromExplorer()
    If $result Then 
        WriteOutput("DEBUG|✓ Found selected file(s) in Explorer")
        Return $result
    EndIf
    WriteOutput("DEBUG|✗ No selected files found in Explorer")
    
    ; Method 3: Get current folder if in Explorer window (LAST RESORT)
    WriteOutput("DEBUG|Method 3: Checking current Explorer folder...")
    $result = GetCurrentFolder()
    If $result Then 
        WriteOutput("DEBUG|✓ Found current Explorer folder")
        Return $result
    EndIf
    WriteOutput("DEBUG|✗ No current Explorer folder found")
    
    ; All methods failed
    WriteOutput("DEBUG|❌ All detection methods failed - nothing valid was selected")
    Return ""
EndFunc

; Get selected files from Windows Explorer using COM
Func GetSelectedFilesFromExplorer()
	Local $oShell = ObjCreate("Shell.Application")
	If @error Then 
		WriteOutput("DEBUG|ERR_COM_FAILED|COM Shell.Application creation failed")
		Return ""
	EndIf
	
	Local $colWindows = $oShell.Windows
	If @error Then
		WriteOutput("DEBUG|ERR_COM_FAILED|Failed to get Shell windows")
		Return ""
	EndIf
    
    ; Find active Explorer window
    For $oWindow In $colWindows
        If IsObj($oWindow) And $oWindow.Visible Then
            Local $sLocationName = String($oWindow.LocationName)
            Local $sFullName = String($oWindow.FullName)
            
            ; Check if this is Windows Explorer or File Explorer
            If StringInStr($sFullName, "explorer.exe") > 0 Then
                Local $oDocument = $oWindow.Document
                If IsObj($oDocument) Then
                    Local $oSelectedItems = $oDocument.SelectedItems()
                    
                    If IsObj($oSelectedItems) And $oSelectedItems.Count > 0 Then
                        ; Get first selected item
                        Local $oItem = $oSelectedItems.Item(0)
                        Local $sPath = String($oItem.Path)
                        
                        If $sPath Then
                            ; Process the selected file/folder
                            Local $itemData = ProcessSelectedFile($sPath)
                            If $itemData Then
                                WriteOutput("DEBUG|Found selected file: " & $sPath)
                                Return $itemData
                            EndIf
                        EndIf
                    Else
                        ; No items selected, try to get current folder
                        Local $currentFolder = String($oWindow.LocationURL)
                        If $currentFolder Then
                            ; Convert file:// URL to path
                            $currentFolder = StringReplace($currentFolder, "file:///", "")
                            $currentFolder = StringReplace($currentFolder, "file://", "")
                            $currentFolder = StringReplace($currentFolder, "/", "\")
                            ; URL decode
                            $currentFolder = StringReplace($currentFolder, "%20", " ")
                            
                            If FileExists($currentFolder) Then
                                WriteOutput("DEBUG|Using current Explorer folder: " & $currentFolder)
                                Return ProcessSelectedFile($currentFolder)
                            EndIf
                        EndIf
                    EndIf
                EndIf
            EndIf
        EndIf
    Next
    
    Return ""
EndFunc

; Get selected text (for URLs and file paths) - ENHANCED FOR URL PRIORITY
Func GetSelectedText()
    ; Save current clipboard content to global variable
    $g_sOriginalClipboard = ClipGet()
    
    WriteOutput("DEBUG|Attempting to capture selected text...")
    
    ; Try multiple copy attempts for better reliability
    Local $selectedText = ""
    Local $attempts = 0
    
    While $attempts < 3 And Not $selectedText
        $attempts += 1
        WriteOutput("DEBUG|Copy attempt " & $attempts & "/3...")
        
        ; Send copy command
        Send("^c")
        Sleep(30 + ($attempts * 10)) ; Slightly longer delay for each attempt
        
        Local $clipContent = ClipGet()
        
        ; Check if we got new content (different from original)
        If $clipContent And $clipContent <> $g_sOriginalClipboard Then
            $selectedText = $clipContent
            WriteOutput("DEBUG|✓ Captured text on attempt " & $attempts & ": " & StringLeft($selectedText, 60) & "...")
            ExitLoop
        EndIf
        
        WriteOutput("DEBUG|✗ Attempt " & $attempts & " failed - no new content")
    WEnd
    
    ; Immediately restore clipboard to reduce disruption
    ClipPut($g_sOriginalClipboard)
    
    ; If we got content, process it with URL priority
    If $selectedText Then
        ; PRIORITY 1: URL detection (most important)
        If IsURL($selectedText) Then
            Local $urlData = CreateUrlData($selectedText)
            If $urlData Then
                WriteOutput("DEBUG|✓ SUCCESS: Detected URL from selected text")
                Return $urlData
            EndIf
        EndIf
        
        ; PRIORITY 2: File path detection
        If IsFilePath($selectedText) Then
            Local $cleanPath = CleanFilePath($selectedText)
            If FileExists($cleanPath) Then
                WriteOutput("DEBUG|✓ SUCCESS: Detected file path from selected text: " & $cleanPath)
                Return ProcessSelectedFile($cleanPath)
            EndIf
        EndIf
        
        WriteOutput("DEBUG|✗ Selected text not recognized as URL or valid file path")
    Else
        WriteOutput("DEBUG|✗ Failed to capture any selected text after " & $attempts & " attempts")
    EndIf
    
    Return ""
EndFunc

; Restore clipboard helper function - NOW HANDLED INLINE FOR BETTER RELIABILITY
; (This function is no longer used - clipboard restoration happens immediately)

; Get current folder when nothing else is available
Func GetCurrentFolder()
    ; Try to get the active window and determine current folder
    Local $hWnd = WinGetHandle("[ACTIVE]")
    Local $sTitle = WinGetTitle($hWnd)
    
    ; If it looks like an Explorer window, try to get the path from COM
    If StringInStr($sTitle, "File Explorer") Or StringInStr($sTitle, "Windows Explorer") Then
        ; Use COM to get current folder
        Local $oShell = ObjCreate("Shell.Application")
        If Not @error Then
            Local $colWindows = $oShell.Windows
            
            For $oWindow In $colWindows
                If IsObj($oWindow) And $oWindow.HWND = $hWnd Then
                    Local $currentPath = String($oWindow.LocationURL)
                    If $currentPath Then
                        $currentPath = StringReplace($currentPath, "file:///", "")
                        $currentPath = StringReplace($currentPath, "file://", "")
                        $currentPath = StringReplace($currentPath, "/", "\")
                        $currentPath = StringReplace($currentPath, "%20", " ")
                        
                        If FileExists($currentPath) Then
                            WriteOutput("DEBUG|Using current folder from active Explorer: " & $currentPath)
                            Return ProcessSelectedFile($currentPath)
                        EndIf
                    EndIf
                EndIf
            Next
        EndIf
    EndIf
    
    ; NO MORE DOCUMENTS FALLBACK - return empty if nothing found
    WriteOutput("DEBUG|No valid selection or folder detected")
    Return ""
EndFunc

; Process a selected file and determine its type
Func ProcessSelectedFile($filePath)
	If Not FileExists($filePath) Then
		WriteOutput("ERROR|ERR_FILE_NOT_FOUND|File does not exist: " & $filePath)
		Return ""
	EndIf
    
    Local $itemData = '{"type":"'
    Local $fileName = ""
    Local $fullExt = ""
    
    ; Get the full extension for better detection
    Local $dotPos = StringInStr($filePath, ".", 0, -1)
    If $dotPos > 0 Then
        $fullExt = StringLower(StringMid($filePath, $dotPos))
    EndIf
    
    ; Check if it's a directory
    If StringInStr(FileGetAttrib($filePath), "D") > 0 Then
        ; It's a directory
        $itemData &= 'file","path":"' & JsonEscape($filePath) & '"'
        $fileName = GetFileNameFromPath($filePath)
        $itemData &= ',"displayName":"' & JsonEscape($fileName) & '"}'
        
    ElseIf $fullExt = ".exe" Or $fullExt = ".lnk" Or $fullExt = ".msi" Then
        ; It's an application or installer
        $itemData &= 'app","path":"' & JsonEscape($filePath) & '"'
        $fileName = GetFileNameWithoutExtension($filePath)
        $itemData &= ',"displayName":"' & JsonEscape($fileName) & '"}'
        
    ElseIf $fullExt = ".url" Then
        ; It's a URL shortcut file
        Local $urlContent = FileRead($filePath)
        If $urlContent Then
            ; Simple regex replacement since we can't use StringRegExp easily
            Local $urlStart = StringInStr(StringUpper($urlContent), "URL=")
            If $urlStart > 0 Then
                Local $urlLine = StringMid($urlContent, $urlStart + 4)
                Local $urlEnd = StringInStr($urlLine, @CR)
                If $urlEnd = 0 Then $urlEnd = StringInStr($urlLine, @LF)
                If $urlEnd = 0 Then $urlEnd = StringLen($urlLine) + 1
                
                Local $extractedUrl = StringLeft($urlLine, $urlEnd - 1)
                $extractedUrl = StringStripWS($extractedUrl, $STR_STRIPLEADING + $STR_STRIPTRAILING)
                
				If $extractedUrl Then
					$itemData = '{"type":"url","url":"' & JsonEscape($extractedUrl) & '"'
					$fileName = GetFileNameWithoutExtension($filePath)
					$itemData &= ',"displayName":"' & JsonEscape($fileName) & '"}'
				Else
					Return '{"type":"error","code":"ERR_URL_PARSE_FAILED","errorMessage":"Could not extract URL from .url file"}'
				EndIf
			Else
				Return '{"type":"error","code":"ERR_URL_PARSE_FAILED","errorMessage":"Could not extract URL from .url file"}'
			EndIf
		Else
			Return '{"type":"error","code":"ERR_FILE_READ_FAILED","errorMessage":"Could not read .url file"}'
		EndIf
        
    Else
        ; It's a regular file
        $itemData &= 'file","path":"' & JsonEscape($filePath) & '"'
        $fileName = GetFileNameWithoutExtension($filePath)
        If $fileName = "" Then
            $fileName = GetFileNameFromPath($filePath)
        EndIf
        $itemData &= ',"displayName":"' & JsonEscape($fileName) & '"}'
    EndIf
    
    Return $itemData
EndFunc

; Create URL data structure
Func CreateUrlData($url)
    Local $urlData = '{"type":"url","url":"' & JsonEscape($url) & '"'
    
    ; Extract display name from URL
    Local $displayName = $url
    If StringInStr($url, "://") Then
        Local $protocolEnd = StringInStr($url, "://") + 3
        Local $hostPart = StringMid($url, $protocolEnd)
        Local $slashPos = StringInStr($hostPart, "/")
        If $slashPos > 0 Then
            $displayName = StringLeft($hostPart, $slashPos - 1)
        Else
            $displayName = $hostPart
        EndIf
    EndIf
    
    $urlData &= ',"displayName":"' & JsonEscape($displayName) & '"}'
    Return $urlData
EndFunc

; Enhanced URL detection function - IMPROVED FOR BETTER RELIABILITY
Func IsURL($text)
    ; Trim whitespace and clean the text
    $text = StringStripWS($text, $STR_STRIPLEADING + $STR_STRIPTRAILING)
    
    ; Skip empty or very short text
    If StringLen($text) < 4 Then Return False
    
    Local $lower = StringLower($text)
    
    ; Method 1: Protocol-based detection (most reliable)
    If StringInStr($text, "://") > 0 Then
        ; Check common protocols (expanded list)
        If StringRegExp($lower, "^(https?|ftp|ftps|file|steam|spotify|discord|slack|zoom|teams|mailto|tel|sms)://") Then
            WriteOutput("DEBUG|URL detected by protocol: " & StringLeft($text, 50) & "...")
            Return True
        EndIf
        
        ; Catch any other protocol format
        If StringRegExp($lower, "^[a-z][a-z0-9+.-]*://") Then
            WriteOutput("DEBUG|URL detected by generic protocol: " & StringLeft($text, 50) & "...")
            Return True
        EndIf
    EndIf
    
    ; Method 2: Common web domains (without protocol)
    If StringRegExp($lower, "^(www\.|[a-z0-9-]+\.)(com|org|net|edu|gov|mil|int|co\.|ac\.|me|io|ly|gl|to|tk|cc)") Then
        WriteOutput("DEBUG|URL detected by domain pattern: " & StringLeft($text, 50) & "...")
        Return True
    EndIf
    
    ; Method 3: IP addresses with ports
    If StringRegExp($text, "^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}(:[0-9]+)?") Then
        WriteOutput("DEBUG|URL detected by IP pattern: " & StringLeft($text, 50) & "...")
        Return True
    EndIf
    
    ; Method 4: Localhost patterns
    If StringRegExp($lower, "^(localhost|127\.0\.0\.1)(:[0-9]+)?") Then
        WriteOutput("DEBUG|URL detected by localhost pattern: " & StringLeft($text, 50) & "...")
        Return True
    EndIf
    
    WriteOutput("DEBUG|Text not recognized as URL: " & StringLeft($text, 30) & "...")
    Return False
EndFunc

; Utility function to check if text is a file path
Func IsFilePath($text)
    ; Check for Windows file path patterns
    If StringLen($text) > 2 And StringMid($text, 2, 2) = ":\" Then
        Return True ; C:\ format
    EndIf
    If StringLeft($text, 2) = "\\" Then
        Return True ; UNC path
    EndIf
    If StringInStr($text, "\") > 0 Or StringInStr($text, "/") > 0 Then
        Return True ; Contains path separators
    EndIf
    Return False
EndFunc

; Clean file path (remove quotes, etc.)
Func CleanFilePath($path)
    Local $cleaned = StringStripWS($path, $STR_STRIPLEADING + $STR_STRIPTRAILING)
    
    ; Remove surrounding quotes
    If StringLen($cleaned) > 2 Then
        If (StringLeft($cleaned, 1) = '"' And StringRight($cleaned, 1) = '"') Or _
           (StringLeft($cleaned, 1) = "'" And StringRight($cleaned, 1) = "'") Then
            $cleaned = StringMid($cleaned, 2, StringLen($cleaned) - 2)
        EndIf
    EndIf
    
    Return $cleaned
EndFunc

; Convert Electron hotkey format to AutoIt format
Func ConvertHotkeyFormat($sHotkey)
    Local $sResult = StringLower($sHotkey)
    
    ; Convert modifier keys
    $sResult = StringReplace($sResult, "control+", "^")
    $sResult = StringReplace($sResult, "ctrl+", "^")
    $sResult = StringReplace($sResult, "alt+", "!")
    $sResult = StringReplace($sResult, "shift+", "+")
    $sResult = StringReplace($sResult, "win+", "#")
    $sResult = StringReplace($sResult, "cmd+", "#")
    $sResult = StringReplace($sResult, "meta+", "#")
    
    ; Handle special keys
    $sResult = StringReplace($sResult, "space", "{SPACE}")
    $sResult = StringReplace($sResult, "enter", "{ENTER}")
    $sResult = StringReplace($sResult, "tab", "{TAB}")
    $sResult = StringReplace($sResult, "escape", "{ESC}")
    $sResult = StringReplace($sResult, "backspace", "{BS}")
    $sResult = StringReplace($sResult, "delete", "{DEL}")
    
    ; Handle function keys
    For $i = 1 To 12
        $sResult = StringReplace($sResult, "f" & $i, "{F" & $i & "}")
    Next
    
    ; Validate the result
    If StringLen($sResult) = 0 Then
        Return ""
    EndIf
    
    Return $sResult
EndFunc

; Helper function to get filename from path
Func GetFileNameFromPath($path)
    Local $lastSlash = StringInStr($path, "\", 0, -1)
    If $lastSlash > 0 Then
        Return StringMid($path, $lastSlash + 1)
    EndIf
    Return $path
EndFunc

; Helper function to get filename without extension
Func GetFileNameWithoutExtension($path)
    Local $fileName = GetFileNameFromPath($path)
    Local $lastDot = StringInStr($fileName, ".", 0, -1)
    If $lastDot > 0 Then
        Return StringLeft($fileName, $lastDot - 1)
    EndIf
    Return $fileName
EndFunc

; Escape JSON special characters
Func JsonEscape($text)
    Local $escaped = $text
    $escaped = StringReplace($escaped, "\", "\\")
    $escaped = StringReplace($escaped, '"', '\"')
    $escaped = StringReplace($escaped, @CR, "\r")
    $escaped = StringReplace($escaped, @LF, "\n")
    $escaped = StringReplace($escaped, @TAB, "\t")
    Return $escaped
EndFunc

; Write output to communication file
Func WriteOutput($sMessage)
    ; Simple file write approach that should work with admin privileges
    Local $result = FileWrite($g_sOutputFile, $sMessage & @CRLF)
    
    ; Fallback: try direct path construction if first attempt fails
    If Not $result Then
        Local $tempPath = @TempDir & "\smarttabs-autoit-output.txt"
        $result = FileWrite($tempPath, $sMessage & @CRLF)
    EndIf
    
    ; Ultimate fallback: write to local directory
    If Not $result Then
        FileWrite("autoit-debug-output.txt", $sMessage & @CRLF)
    EndIf
    
EndFunc

; Cleanup function
Func OnExit()
    WriteOutput("HELPER_EXITING")
    
    ; Clean up hotkey
    UnregisterCurrentHotkey()
    
    ; Clean up temp files
    FileDelete($g_sCommandFile)
    ; Leave output file for final messages to be read
EndFunc