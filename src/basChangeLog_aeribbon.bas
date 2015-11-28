Option Compare Database
Option Explicit

' Custom Usage:
' FRONT END SETUP
Public Const THE_FRONT_END_APP = True
Public Const THE_SOURCE_FOLDER = ".\src\"                  ' "C:\THE\DATABASE\PATH\src\"
Public Const THE_XML_FOLDER = ".\src\xml\"                 ' "C:\THE\DATABASE\PATH\src\xml\"
Public Const THE_XML_DATA_FOLDER = ".\src\xmldata\"        ' "C:\THE\DATABASE\PATH\src\xmldata\"
'Public Const THE_BACK_END_DB1 = "C:\MY\BACKEND\DATA.accdb"
Public Const THE_BACK_END_SOURCE_FOLDER = "NONE"           ' ".\srcbe\"
Public Const THE_BACK_END_XML_FOLDER = "NONE"              ' ".\srcbe\xml\"
Public Const THE_BACK_END_XML_DATA_FOLDER = "NONE"         ' ".\srcbe\xmldata\"

' BACK END SETUP
'Public Const THE_FRONT_END_APP = False
'Public Const THE_SOURCE_FOLDER = "NONE"                     ' ".\src\"
'Public Const THE_XML_FOLDER = "NONE"                        ' ".\src\xml\"
'Public Const THE_XML_DATA_FOLDER = "NONE"                   ' ".\src\xmldata\"
'Public Const THE_BACK_END_DB1 = "NONE"
'Public Const THE_BACK_END_SOURCE_FOLDER = "C:\THE\DATABASE\PATH\srcbe\"             ' ".\srcbe\"
'Public Const THE_BACK_END_XML_FOLDER = "C:\THE\DATABASE\PATH\srcbe\xml\"            ' ".\srcbe\xml\"
'Public Const THE_BACK_END_XML_DATA_FOLDER = "C:\THE\DATABASE\PATH\srcbe\xmldata\"   ' ".\srcbe\xmldata\
'
Public Const gconTHIS_APP_VERSION As String = "0.0.2"
Public Const gconTHIS_APP_VERSION_DATE = "November 28, 2015"
Public Const gconTHIS_APP_NAME = "aeribbon"

Public Function aeribbon_EXPORT(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR
 
    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER
    End If
 
PROC_EXIT:
     Exit Function
 
PROC_ERR:
     MsgBox "Erl=" & Erl & " Err=" & Err & " (" & Err.description & ") in procedure aeribbon_EXPORT"
     Resume Next

End Function

'=============================================================================================================================
' Tasks:
' %010 -
' %009 -
' %008 -
' %007 -
' %006 -
' %005 - Document changes from blank accdb to minimal app template
' %004 - Replace basGDIPlus with latest GDayClass and TimerClass
' %003 - Use splash form with aeternity logo, load from binary table?
' %002 - Create setup tab form for loading images into the binary table
' %001 - Configure code and ribbon to only load pix from internal binary table
' %000 - Deatiled information for Ribbon development can be found here:
'           http://www.accessribbon.de/en/index.php?Downloads:12
'
'=============================================================================================================================
'
'
'20151128 v002 -
'20151128 v001 - First version commit, simple ribbon, minimal code