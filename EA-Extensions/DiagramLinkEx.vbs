'[path=\EA-Extensions]
'[group=EA-Extensions]

!INC Logging.LogManager

' 
' A straight DiagramLink has no Path
' Path values are a list of x:y values (separated by semi-colons) See EA-Extensions/EA Encodings for class to handle these
'
' Geometry values are a list of key=values (separated by semi-colons) See EA-Extensions/EA Encodings for class to handle these
' Geometry.SX/SY are the Start X and Start Y values (range from -(Width/2) on the left to Width/2)
' Geometry.EX/EY are the End X and End Y values (range from -(Height/2) on the bottom to Height/2)
'   These values are relative to the object connected, a 0 value indicates a center value.
' Geometry.EDGE indicate which side the link is connected to; 1=Top, 2=Right, 3=Bottom, 4=Left
'   It is unknown how the end connection edge is determined
' See https://sparxsystems.com/forums/smf/index.php/topic,4921.msg122670.html#msg122670 for more details
'
'
' Example DiagramLink Geometry and Paths:
' 3/14/2023 10:13:17 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Geometry=SX=0;SY=0;EX=0;EY=0;EDGE=1;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;
' 3/14/2023 11:09:52 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Path=
'
' 3/14/2023 10:13:17 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Geometry=SX=-15;SY=-34;EX=1;EY=34;EDGE=3;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=-25:OY=6:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;
' 3/14/2023 11:09:52 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Path=330:-410;155:-410;
'
' 3/14/2023 10:13:18 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Geometry=SX=0;SY=0;EX=0;EY=0;EDGE=3;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;
' 3/14/2023 11:09:53 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Path=
'
' 3/14/2023 10:13:19 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Geometry=SX=15;SY=-34;EX=1;EY=34;EDGE=3;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=5:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;
' 3/14/2023 11:09:12 AM DEBUG [ArchiMate.Style Layout Connectors]   diagramLink.Path=360:-410;536:-410;

class DiagramLinkExtenionClass
	private logger
	private m_DiagramLink

	Private Sub Class_Initialize
		set logger = LogManager.getLogger("DiagramLinkEx")
		set m_Diagram = nothing
	End Sub
	
	Public sub init(diagramLink)
		set m_DiagramLink = diagramLink
	end sub

end class

class DiagramLinkExtensionNamespace
	private logger

	Private Sub Class_Initialize
		set logger = LogManager.getLogger("DiagramLinkEx")
	End Sub
	
	function createForDiagramLink(diagramLink)
		dim result
		set result = new DiagramLinkExtenionClass
		result.init diagramLink
		set createForDiagramLink = result
	end function
end class

dim DiagramLinkExtension
set DiagramLinkExtension = new DiagramLinkExtensionNamespace