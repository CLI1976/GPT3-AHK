; AutoHotkey script that enables you to use GPT3 in any input field on your computer

; -- Configuration --
#SingleInstance  ; Allow only one instance of this script to be running.

; This is the hotkey used to autocomplete prompts
; HOTKEY_AUTOCOMPLETE = #o  ; Win+o
; This is the hotkey used to edit text
; HOTKEY_EDIT = #+o  ; Win+shift+o
; Models settings
global MODEL_ENDPOINT := "https://api.openai.com/v1/chat/completions"
global MODEL_AUTOCOMPLETE_ID := "gpt-4o-mini"
CUSTOM_MODEL_ENDPOINT := "http://0.0.0.0:8000/chat/completions"
CUSTOM_MODEL_ID := "together_ai/togethercomputer/llama-2-70b-chat"
MODEL_AUTOCOMPLETE_MAX_TOKENS := 1000
MODEL_AUTOCOMPLETE_TEMP := 0.8

; -- Initialization --
; Dependencies
; WinHttpRequest: https://www.reddit.com/comments/mcjj4s
; cJson.ahk: https://github.com/G33kDude/cJson.ahk
#Include <Json>
http := WinHttpRequest()

I_Icon = GPT3-AHK.ico
IfExist, %I_Icon%
Menu, Tray, Icon, %I_Icon%
; Create custom menus
Menu, LLMMenu, Add, GPT3.5 turbo, SelectLLMHandler
Menu, LLMMenu, Add, GPT4o, SelectLLMHandler
Menu, LLMMenu, Add, GPT4o mini, SelectLLMHandler
Menu, LLMMenu, Add, Custom, SelectLLMHandler
Menu, Tray, Add, Select LLM, :LLMMenu  
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, NoStandard
Menu, Tray, Standard
; Hotkey, %HOTKEY_AUTOCOMPLETE%, AutocompleteFcn
; Hotkey, %HOTKEY_EDIT%, InstructFcn

IfNotExist, settings.ini     
{
   InputBox, OPENAI_API_KEY, Please insert your OpenAI API key, Open AI API key, , 270, 145
   IniWrite, %OPENAI_API_KEY%, settings.ini, OpenAI, API_KEY
   MsgBox, 4,, Do you want to use a custom model?
   ; Check the result of the message box
   IfMsgBox, Yes
      InputBox, CUSTOM_API_KEY, Please insert your custom model API key, Custom Model API key, , 270, 145
      IniWrite, %CUSTOM_API_KEY%, settings.ini, CustomModel, API_KEY      
   IfMsgBox, No
      IniWrite, %CUSTOM_API_KEY%, settings.ini, CustomModel, ""  
} 
Else
{
   IniRead, OPENAI_API_KEY, settings.ini, OpenAI, API_KEY  
   IniRead, CUSTOM_API_KEY, settings.ini, CustomModel, API_KEY  
}
global API_KEY := OPENAI_API_KEY

; Define the custom menu
Menu, MyCustomMenu, Add, 翻成中文, FunctionA
Menu, MyCustomMenu, Add, 修正文法及錯字, FunctionB
Menu, MyCustomMenu, Add, 翻成英文, FunctionC

; Hotkey to trigger the menu on Ctrl + Right Mouse Button
^RButton::
    ; Get the position of the mouse cursor
    ; global X, Y  ; 定義全局變數
      MouseGetPos, X, Y
    ; Show the custom menu at the cursor position
    Menu, MyCustomMenu, Show, %X%, %Y%
return

OnExit("ExitFunc")


SelectLLMHandler:
   if (A_ThisMenuItem = "GPT3.5 turbo") {
      API_KEY := OPENAI_API_KEY
      MODEL_ENDPOINT := "https://api.openai.com/v1/chat/completions"
      MODEL_AUTOCOMPLETE_ID := "gpt-3.5-turbo"	
   } else if (A_ThisMenuItem = "GPT4o") {
      API_KEY := OPENAI_API_KEY
      MODEL_ENDPOINT := "https://api.openai.com/v1/chat/completions"
      MODEL_AUTOCOMPLETE_ID := "gpt-4"
   } else if (A_ThisMenuItem = "GPT4o mini") {
      API_KEY := OPENAI_API_KEY
      MODEL_ENDPOINT := "https://api.openai.com/v1/chat/completions"
      MODEL_AUTOCOMPLETE_ID := "gpt-4o-mini"
   } else if (A_ThisMenuItem = "Custom") {
      API_KEY := CUSTOM_API_KEY
      MODEL_ENDPOINT := CUSTOM_MODEL_ENDPOINT
      MODEL_AUTOCOMPLETE_ID := CUSTOM_MODEL_ID
   } 
   Return

; -- Main commands --

; Function to show the GUI with the response
ShowResponseGui(responseText, chatMod)
{
    global MyEdit  ; 確保 MyEdit 是全局變量
    CoordMode, Mouse , Screen
    MouseGetPos, X1, Y1
        Gui, Color, 85ddda
    ; Gui, Font, s14
     Gui, Font, s14 , 微軟正黑體   ; 改大字體
    ;  Gui, Font, s14 , Gen Jyuu Gothic Monospace Normal  ; 改大字體
    Gui, Add, Edit, w600 h400 vMyEdit ReadOnly, 
    GuiControl, , MyEdit, %responseText%
    Gui, Show, x%X1% y%Y1% w630 h420, %chatMod% says
    return


}

    GuiClose:
    Gui, Destroy  ; 銷毀 GUI 窗口和相關變數
    return


; Function A 翻成中文
FunctionA:
   GetText(CopiedText, "Copy")
   text2translate := "translate the following text to zh-tw -- " CopiedText
   url := MODEL_ENDPOINT
   body := {}
   body.model := MODEL_AUTOCOMPLETE_ID ; ID of the model to use.   
   body.messages := [{"role": "user", "content": text2translate}] ; The prompt to generate completions for
   body.max_tokens := MODEL_AUTOCOMPLETE_MAX_TOKENS ; The maximum number of tokens to generate in the completion.
   body.temperature := MODEL_AUTOCOMPLETE_TEMP + 0 ; Sampling temperature to use 
   headers := {"Content-Type": "application/json", "Authorization": "Bearer " . API_KEY}
   TrayTip, GPT3-AHK, Asking ChatGPT...
   SetSystemCursor()
   response := http.POST(url, JSON.Dump(body), headers, {Object:true, Encoding:"UTF-8"})
; ======vvv Debug vvvvvv=======
; 檢查並創建目標目錄
targetDir := "C:\temp"
if !FileExist(targetDir)
{
    FileCreateDir, %targetDir%
}
; 將 response.Text 內容寫入到文件中
targetFile := targetDir "\response001.txt"
     ; 刪除現有的 test.txt 文件
    FileDelete, % targetFile
; FileAppend, % response.Headers, % targetFile
; FileAppend, % response.Status, % targetFile
; FileAppend, % response.Text, % targetFile
; msgbox, % response.Text
; msgbox, % "Status Code: " response.Status "`nResponse Text saved to: " targetFile
for key, value in response
{
    FileAppend, %key%: %value%`n, %targetFile%
}

; msgbox, The response has been saved to %targetFile%
; ======^^^ Debug ^^^=======
   obj := JSON.Load(response.Text)
   ans := obj.choices[1].message.content
   mod := obj.model
    ; Show the GUI with the response
    ; msgbox, %text2translate%
    ; msgbox, %url%
    ; msgbox, %API_KEY%
    ; msgbox, %mod%
    ShowResponseGui(ans, mod)

   RestoreCursors()   
   TrayTip
   Return


; Function B 修正文法及錯字
FunctionB:
   GetText(CopiedText, "Copy")
   text2corrent := "Help me check the grammar and spelling of the following sentences, and output the corrected sentences. -- " CopiedText
   url := MODEL_ENDPOINT
   body := {}
   body.model := MODEL_AUTOCOMPLETE_ID ; ID of the model to use.   
   body.messages := [{"role": "user", "content": text2corrent}] ; The prompt to generate completions for
   body.max_tokens := MODEL_AUTOCOMPLETE_MAX_TOKENS ; The maximum number of tokens to generate in the completion.
   body.temperature := MODEL_AUTOCOMPLETE_TEMP + 0 ; Sampling temperature to use 
   headers := {"Content-Type": "application/json", "Authorization": "Bearer " . API_KEY}
   TrayTip, GPT3-AHK, Asking ChatGPT...
   SetSystemCursor()
   response := http.POST(url, JSON.Dump(body), headers, {Object:true, Encoding:"UTF-8"})
   obj := JSON.Load(response.Text)
   ans := obj.choices[1].message.content
   mod := obj.model
    ; Show the GUI with the response
    ShowResponseGui(ans, mod)

   RestoreCursors()   
   TrayTip
   Return

; Function C 翻成英文
FunctionC:
   GetText(CopiedText, "Copy")
   text2eng := "translate the following text to English -- " CopiedText
   url := MODEL_ENDPOINT
   body := {}
   body.model := MODEL_AUTOCOMPLETE_ID ; ID of the model to use.   
   body.messages := [{"role": "user", "content": text2eng}] ; The prompt to generate completions for
   body.max_tokens := MODEL_AUTOCOMPLETE_MAX_TOKENS ; The maximum number of tokens to generate in the completion.
   body.temperature := MODEL_AUTOCOMPLETE_TEMP + 0 ; Sampling temperature to use 
   headers := {"Content-Type": "application/json", "Authorization": "Bearer " . API_KEY}
   TrayTip, GPT3-AHK, Asking ChatGPT...
   SetSystemCursor()
   response := http.POST(url, JSON.Dump(body), headers, {Object:true, Encoding:"UTF-8"})
   obj := JSON.Load(response.Text)
   ans := obj.choices[1].message.content
   mod := obj.model
    ; Show the GUI with the response
    ShowResponseGui(ans, mod)

   RestoreCursors()   
   TrayTip
   Return


; -- Auxiliar functions --
; Copies the selected text to a variable while preserving the clipboard.
GetText(ByRef MyText = "", Option = "Copy")
{
   SavedClip := ClipboardAll
   Clipboard =
   If (Option == "Copy")
   {
      Send ^c
   }
   Else If (Option == "Cut")
   {
      Send ^x
   }
   ClipWait 0.5
   If ERRORLEVEL
   {
      Clipboard := SavedClip
      MyText =
      Return
   }
   MyText := Clipboard
   Clipboard := SavedClip
   Return MyText
}

; Send text from a variable while preserving the clipboard.
PutText(MyText, Option = "")
{
   ; Save clipboard and paste MyText
   SavedClip := ClipboardAll 
   Clipboard = 
   Sleep 20
   Clipboard := MyText
   If (Option == "AddSpace")
   {
      Send {Right}
      Send {end}
      Send {Space}
   }
   Send ^v
   Sleep 100
   Clipboard := SavedClip
   Return
}   


; Change system cursor 
SetSystemCursor()
{
   Cursor = %A_ScriptDir%\GPT3-AHK.ani
   CursorHandle := DllCall( "LoadCursorFromFile", Str,Cursor )

   Cursors = 32512,32513,32514,32515,32516,32640,32641,32642,32643,32644,32645,32646,32648,32649,32650,32651
   Loop, Parse, Cursors, `,
   {
      DllCall( "SetSystemCursor", Uint,CursorHandle, Int,A_Loopfield )
   }
}

RestoreCursors() 
{
   DllCall( "SystemParametersInfo", UInt, 0x57, UInt,0, UInt,0, UInt,0 )
}

ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        RestoreCursors()
    }
}