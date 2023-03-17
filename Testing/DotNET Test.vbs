option explicit

'[path=\Testing]
'[group=Testing]

!INC Local Scripts.EAConstants-VBScript

sub TestDotNet3_5
	dim dictionary
	set dictionary = CreateObject("Scripting.Dictionary")
end sub

' Can't use Generics in VBScript
sub TestDotNet4_7_GenericDictionary
	dim dictionary
	set dictionary = CreateObject("System.Collections.Generic.Dictionary")
	dictionary.Add "test", "value"
	Session.Output "Dictionary key=test, value=" & dictionary("test")
end sub

sub TestDotNet4_7_Hashtable
	dim dictionary
	set dictionary = CreateObject("System.Collections.Hashtable")
	dictionary.Add "test", "value"
	Session.Output "Dictionary key=test, value=" & dictionary("test")
end sub

sub main
	TestDotNet4_7_Hashtable
end sub

main