Attribute VB_Name = "ErrorType"
Public ERRORTYPE_NULL As Long
Public ERRORTYPE_DIV0 As Long
Public ERRORTYPE_VALUE As Long
Public ERRORTYPE_REF As Long
Public ERRORTYPE_NAME As Long
Public ERRORTYPE_NUM As Long
Public ERRORTYPE_NA As Long

Sub Initialize()
	ERRORTYPE_NULL = 0
	ERRORTYPE_DIV0 = 7
	ERRORTYPE_VALUE = 15
	ERRORTYPE_REF = 23
	ERRORTYPE_NAME = 29
	ERRORTYPE_NUM = 36
	ERRORTYPE_NA = 42
End Sub
