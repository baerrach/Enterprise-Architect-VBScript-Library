'[path=\Framework\Utils]
'[group=Utils]

!INC Local Scripts.EAConstants-VBScript

dim defaultColor

defaultColor 					= -1

' From http://www.sparxsystems.com.au/enterprise_architect_user_guide/11/automation_and_scripting/diagramobjects.html
' The color value is a decimal representation of the hex RGB value, where Red=FF, Green=FF00 and Blue=FF0000
' Who would write an RGB as BGR. YAEAB
' red, green, blue are string for the hex values 00-FF
function SparxColorFromRGB(red, green, blue)
	SparxColorFromRGB = CLng("&h" & blue & green & red)
end function

' Convert a Sparx color value into array[red, green, blue] 
' the values of red, green, blue are decimal values
function SparxColorToRGB(color)
	dim red, green, blue

	red = color mod 16^2
	green = (color \ (16^2)) mod 16^2
	blue = (color \ (16^4)) mod 16^2

	SparxColorToRGB = Array(red, green, blue)
end function

' Convert a red, green, blue in hex values (or decimal) to a SparxColor
function RGBtoSparxColor(red, green, blue)
	RGBtoSparxColor = RGB(red, green, blue)
end function

sub main
	dim rgb
	rgb = SparxColorToRGB(15138790)
	Session.Output rgb(0) & ", " & rgb(1) & ", " & rgb(2)
	
	Session.Output SparxColorFromRGB(HEX(230), HEX(255), HEX(230))
	Session.Output RGBtoSparxColor(230, 255, 230)
	Session.Output RGBtoSparxColor(&HE6, &HFF, &HE6)	
end sub
