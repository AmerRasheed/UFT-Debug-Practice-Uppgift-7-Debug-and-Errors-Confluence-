'On Error Resume Next
'a = functionA
'
'
'On Error Goto 0
'
'Function functionA
'	Print "functionA started"
'	
'	On Error Resume Next
'	
'		Call functionB
'		
'		functionA = 4/0	
'	
'End Function
'
'Function functionB
'	Print "functionB started"
'	On Error Resume Next
'	functionB = 20/0
'	
'	Call functionC
'
'End Function
'
'Function functionC
'	'On Error Resume Next
'       'functionC = 20/0
'	Print "Hammarby är bäst!"
'End Function
