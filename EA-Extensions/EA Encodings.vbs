'[path=\EA-Extensions]
'[group=EA-Extensions]

class SparxKeyValueEncodedString
	private DEFAULT_SEPARATOR
	
	private logger
	
	private m_separator
	private m_KeyValues
	private m_KeysInOrder

	Private Sub Class_Initialize
		DEFAULT_SEPARATOR = ";"
		m_separator = DEFAULT_SEPARATOR
		set logger = LogManager.getLogger("SparxKeyValueEncodedString")
		set m_KeyValues = CreateObject("Scripting.Dictionary")
		set m_KeysInOrder = CreateObject("System.Collections.ArrayList")
	End Sub
	
	' Count property.
	Public Property Get Count
		Count = m_KeysInOrder.Count
	End Property
	
	Public Property Get Item(key)
		Item = ""
		if m_KeyValues.Exists(key) then
			Item = m_KeyValues.Item(key)
		end if
	end Property
	
	Public sub Add(key, value)
		if not m_KeysInOrder.contains(key) then
			m_KeysInOrder.Add key
		end if
		m_KeyValues(key) = value
	end sub
	
	public function toString()
		dim key, value, keys, keyValueEncodedStrings, keyAndValueEncoded
		keys = m_KeyValues.Keys()
		set keyValueEncodedStrings = CreateObject("System.Collections.ArrayList")
		for each key in m_KeysInOrder
			value = m_KeyValues.Item(key)
			keyValueEncodedStrings.Add key & "=" & value
		next
		
		' Can't use Join, Sparx has trailing ;
		toString = ""
		for each keyAndValueEncoded in keyValueEncodedStrings.ToArray
			toString = toString & keyAndValueEncoded & m_separator
		next
	end function
	
	' resets this object and initialises using the string provided
	' other form:
	' key1=value;key2=value2;...
	public sub fromString(keyValueEncodedString)
		set m_KeyValues = CreateObject("Scripting.Dictionary")
		set m_KeysInOrder = CreateObject("System.Collections.ArrayList")

		dim keyValueEncodedStrings
		keyValueEncodedStrings = Split(keyValueEncodedString, m_separator)
		
		dim rx, matches, match
		set rx = new RegExp
		rx.Pattern = "^([^=]*)=(.*)$"
		
		dim keyAndValueEncoded, keyValue, key, value
		for each keyAndValueEncoded in keyValueEncodedStrings
			logger.trace "keyAndValueEncoded=" & keyAndValueEncoded
			
			' Can't use Split as EA can encode a value with a different separator for more key/values
			set matches = rx.Execute(keyAndValueEncoded)
			if Matches.count = 1 then
				set match = matches(0)
				key = match.SubMatches(0)
				value = match.SubMatches(1)
				
				logger.trace "key=" & key
				logger.trace "value=" & value

				if key <> "" then
					m_KeysInOrder.Add key
					m_KeyValues.Add key, value
				end if
			end if
		next
	end sub
end class