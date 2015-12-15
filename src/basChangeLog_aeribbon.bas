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
Public Const gconTHIS_APP_VERSION As String = "0.0.6"
Public Const gconTHIS_APP_VERSION_DATE = "December 13, 2015"
Public Const gconTHIS_APP_NAME = "aeribbon"

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = gconTHIS_APP_VERSION
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = gconTHIS_APP_VERSION_DATE
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gconTHIS_APP_NAME
End Function

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
     MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.description & " in procedure aeribbon_EXPORT", vbInformation, gconTHIS_APP_NAME
     Resume Next

End Function

'=============================================================================================================================
' Tasks:
' %020 -
' %019 -
' %018 -
' %017 - Add a form for loading a picture to the binary table
' %009 - Update code to use standard naming protocol
' %006 - Load app logo from attachment table into splash form
' %005 - Document changes from blank accdb to minimal app template
' %002 - Create setup tab form for loading images into the attachment table
' %000 - Detailed information for Ribbon development can be found here:
'           http://www.accessribbon.de/en/index.php?Downloads:12
'           https://msdn.microsoft.com/en-us/library/bb386089.aspx
'           http://www.excelguru.ca/blog/category/the-ribbon/
'
'=============================================================================================================================
'
'
'20151213 v006 -
    ' FIXED - %016 - Relates to %103, remove unused code and defines
'20151212 v005 -
    ' WONTFIX - %014 - Test RemoveSysMenu and related if splash form popup has a menubar - Not needed if popup is required ?
    ' FIXED - %013 - Fade causes flickering, test using repaint to see if it improves
    '           repaint does not help, but fade is ok but fast on Dell laptop so changed sleep to 40
'20151211 v004 -
    ' FIXED - %015 - Repeated click on splash image with fade crashes access if it catches the form still loaded
    ' FIXED - %012 - Form does not display before running code, Ref: https://bytes.com/topic/access/answers/449160-how-get-form-display-first-then-run-code-open
    ' FIXED - %011 - Run-time error 2585 in Transparency - need to explicitly target the correct window
    ' FIXED - %010 - Use function to get handle for splash form in the current event
    ' FIXED - %008 - Debug fade code for splash form
    ' FIXED - %007 - Add code to allow switch between fade and no fade
    ' FIXED - Move old references from form load as they relate to original excel implementation
    ' Reference: Microsoft Knowledge Base Article Q213774, XL2000 - How to Create a Startup Screen with a UserForm
    ' Who:  Peter Ennis
    ' Date: 04/06/2007 - Add error checking code to track down 1004 on splash form close
    '       04/15/2007 - Use gblnClockOn for starting the clock
    '       05/03/2007 - Remove gblnClockOn. cmd30 now used
    '       02/15/2010 - Modify for aeChart
    '       12/06/2015 - Modify to use transparency in Access
'20151203 v003 -
    ' FIXED - %003 - Use splash form with aeternity logo
'20151129 v002 -
    ' FIXED - %004 - Replace basGDIPlus with latest GDayClass and TimerClass
    ' FIXED - %001 - Configure code and ribbon to only load pix from internal image attachment table
'20151128 v001 - First version commit, simple ribbon, minimal code