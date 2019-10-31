Attribute VB_Name = "FontSettings"
Public FONTSETTINGS_UNDERLINE_STYLE_NONE As Long
Public FONTSETTINGS_UNDERLINE_STYLE_SINGLE As Long
Public FONTSETTINGS_UNDERLINE_STYLE_DOUBLE As Long
Public FONTSETTINGS_UNDERLINE_STYLE_SINGLE_ACCOUNTING As Long
Public FONTSETTINGS_UNDERLINE_STYLE_DOUBLE_ACCOUNTING As Long
Public FONTSETTINGS_HEADINGS_FONT = "HEADINGS" As String
Public FONTSETTINGS_BODY_FONT = "BODY" As String

Sub Initialize()
	FONTSETTINGS_UNDERLINE_STYLE_NONE = 0
	FONTSETTINGS_UNDERLINE_STYLE_SINGLE = 1
	FONTSETTINGS_UNDERLINE_STYLE_DOUBLE = 2
	FONTSETTINGS_UNDERLINE_STYLE_SINGLE_ACCOUNTING = 33
	FONTSETTINGS_UNDERLINE_STYLE_DOUBLE_ACCOUNTING = 34
	FONTSETTINGS_HEADINGS_FONT = "HEADINGS"
	FONTSETTINGS_BODY_FONT = "BODY"
End Sub