andreas = 30
fredrik = "barnet"
'fredrik = 20

Print AddAges(andreas, fredrik)

Function AddAges(age1, age2)
	
	If Not (AgeIsNumeric(age1) And AgeIsNumeric(Age2)) Then
		Err.Raise 1000, "AddAges", "One or both ages are not numeric..."
	End If	
	
	End Function

Function AgeIsNumeric(age)

If Not IsNumeric(Age) Then
		AgeIsNumeric = False
	Else
		AgeIsNumeric = True
	End If
	
End Function
