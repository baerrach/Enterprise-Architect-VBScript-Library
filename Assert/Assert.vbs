'[path=\Assert]
'[group=Assert]

sub assertTrue(message, condition)
	if not condition then
		Err.Raise vbObjectError + 1, "assertTrue", message & " " & "expected=true" & " but was actual=" & condition
	end if
end sub

sub assertFalse(message, condition)
	if condition then
		Err.Raise vbObjectError + 1, "assertFalse", message & " " & "expected=false" & " but was actual=" & condition
	end if
end sub

sub assertSame(message, expected, actual)
	if not expected is actual then
		Err.Raise vbObjectError + 1, "assertSame", message & " (can't output objects)"
	end if
end sub

sub assertEquals(message, expected, actual)
	if expected <> actual then
		Err.Raise vbObjectError + 1, "assertEquals", message & " " & "expected=" & expected & " but was actual=" & actual
	end if
end sub

sub assertNotNothing(message, actual)
	if actual is nothing then
		Err.Raise vbObjectError + 1, "assertNotNothing", message & " " & "expected=something but was actual=nothing"
	end if
end sub