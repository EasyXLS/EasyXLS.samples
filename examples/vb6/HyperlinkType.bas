Attribute VB_Name = "HyperlinkType"
Public HYPERLINKTYPE_NOHYPERLINK As String
Public HYPERLINKTYPE_URL As String
Public HYPERLINKTYPE_FILE As String
Public HYPERLINKTYPE_UNC As String
Public HYPERLINKTYPE_CELL As String
Public HYPERLINKTYPE_UNKNOWN As String

Sub Initialize()
	HYPERLINKTYPE_NOHYPERLINK = "nohyperlink"
	HYPERLINKTYPE_URL = "url"
	HYPERLINKTYPE_FILE = "file"
	HYPERLINKTYPE_UNC = "unc"
	HYPERLINKTYPE_CELL = "cell"
	HYPERLINKTYPE_UNKNOWN = "unknown"
End Sub
