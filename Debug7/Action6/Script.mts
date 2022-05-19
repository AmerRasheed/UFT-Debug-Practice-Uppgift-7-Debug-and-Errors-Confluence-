'Function BestWayToOpenABrowser
''    CreateObject(...)
''    Navigate...
''    Print..
''   
''   
'EndFunction
'Function RemoveBlancs(stringToRemoveBlancsFrom)
'    stringToRemoveBlancsFrom =Replace(stringToRemoveBlancsFrom, " ", "")
'    RemoveBlancs =Replace(stringToRemoveBlancsFrom, vbTAB, "")
'''white_check_mark
''eyes
''raised_hands
''1:57
'Endfunction
'
stringToPlayWith ="Hej Hej"&vbTAB&" Hej"
 Print "Before: "& stringToPlayWith
 RemoveBlancs(stringToRemoveBlancsFrom)
 stringToPlayWith = RemoveBlancs(stringToPlayWith)
 Print "After: "& stringToPlayWith
'One liner descriptitive programming for Test Objects
'Browser("index:=0").Highlight
 Browser("index:=0").Page("index:=0").WebButton("innertext:=Dela").Highlight
'Dynamic desriptive programming with filter
'Set oDesc =Description.Create()
' oDesc("micclass").Value ="Browser"
'Set childCollection = Desktop.ChildObjects(oDesc)
' Print "We found "& childCollection.count &" objects matching the description"
	For x = 0 to childCollection.Count -1
	    childCollection.Item(x).Highlight 'What??
	Next
	
