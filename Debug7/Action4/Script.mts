'I make a change in comments for GIT 
'Dim aIE
'Dim MyDocument
'Set aIE = CreateObject("InternetExplorer.Application")


command = "chrome -disable-session-crashed-bubble " & " https://www.blocket.se/stockholm?ca=11"
Set objBrowser = CreateObject("Wscript.Shell")
objBrowser.Run(command)
'Maximize the Browser Window
Window("nativeclass:=Chrome_WidgetWin_1","regexpwndtitle:= Google Chrome").Maximize



Set objBrowser = Nothing

Set SearchField = Browser("index:=0").Page("index:=0").WebEdit("xpath:=//DIV[1]/DIV[1]/DIV[1]/FORM[1]/DIV[1]/DIV[@role=""combobox""][1]/INPUT[1]")
SearchField.Highlight
SearchField.Set "damcykel"

Set SearchButton = Browser("index:=0").Page("index:=0").WebButton("xpath:=//DIV/DIV/DIV/FORM/DIV/DIV[@role=""combobox""]/DIV/BUTTON[normalize-space()=""Sök""]")
SearchButton.Highlight
SearchButton.Click

'SearchField.Set "Amer"






'While IE.ReadyState <> 4 
'While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend  
'aIE.Visible = true 
'aIE.navigate " https://www.blocket.se/stockholm?ca=11"
'Browser("index:=0").Page("index:=0").WebButton("innertext:=Dela").Highlight

'Browser("index:=0").Page("index:=0").Highlight

'Browser.Page("index:=0").WebButton("innertext:=0")

'Wend

'SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe", "https://www.blocket.se/stockholm?ca=11"
'Browser("index:=0").Page("index:=0").WebEdit("placeholder:=Vad vill du söka efter?").Highlight
'

'Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
'IE.Navigate "url"
'With IE.Document
'
'Do
'if not CreateObject(.getElementByID("formButton2343255")) is nothing then            
'.getElementByID("formButton2343255").Click()
'Exit Do
'End if
'WScript.Sleep 500 
'Loop
'
'SET objWshShell = Nothing
'End With
'End Function
'
'Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
'
'IE.Navigate "https://stackoverflow.com"
'IE.Visible = True
'
''Wait til DOM is ready
'Do Until IE.ReadyState = 4 : Loop
'
'If IsObject(IE.Document.GetElementById("nav-tags")) Then
'    IE.Document.GetElementById("nav-tags").Click()
'End If
'
'Set IE = Nothing
