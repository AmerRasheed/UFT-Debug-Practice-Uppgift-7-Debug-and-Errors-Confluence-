command = "chrome -disable-session-crashed-bubble " & " https://www.blocket.se/stockholm?ca=11"
Set objBrowser = CreateObject("Wscript.Shell")
objBrowser.Run(command)
'Maximize the Browser Window
Window("nativeclass:=Chrome_WidgetWin_1","regexpwndtitle:= Google Chrome").Maximize

Set objBrowser = Nothing

Set SearchField = Browser("index:=0").Page("index:=0").WebEdit("xpath:=//DIV[1]/DIV[1]/DIV[1]/FORM[1]/DIV[1]/DIV[@role=""combobox""][1]/INPUT[1]")
SearchField.Highlight
SearchField.Set "stratocaster"

Set SearchButton = Browser("index:=0").Page("index:=0").WebButton("xpath:=//DIV/DIV/DIV/FORM/DIV/DIV[@role=""combobox""]/DIV/BUTTON[normalize-space()=""Sök""]")
SearchButton.Highlight
SearchButton.Click


'Set Focus = Description.Create()

'Filter("micclass").Value = "Highlight"

'Set allCheckboxes = Browser("Job").Page("QTP").ChildObjects(Filter)

Set oDesc = Description.Create()
 
oDesc("micclass").Value = "Link"
oDesc("html tag").Value = "A"
'oDesc("class").Value = "item_link"

Set childCollection = Browser("index:=0").Page("index:=0").ChildObjects(oDesc)
 
For x = 0 to childCollection.Count - 1
 	
If InStr(childCollection.Item(x).GetROproperty("text"),"Fender") Then
	childCollection.Item(x).Highlight
End If
    
Next
