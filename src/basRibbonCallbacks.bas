Option Compare Database
Option Explicit

Public gobjRibbon As IRibbonUI

Private Const mRIBBON_PIX As String = "tblRibbonPix"

' For Sample Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Sub LoadImages(control, ByRef Image)
' Loads an image with transparency to the ribbon
' Modul basGDIPlus is required

    Dim strImage        As String
    Dim strPicture      As String

    strImage = CStr(control)
    strPicture = getPic(strImage)

    If strImage <> "" Then
        If strPicture <> "" Then
            Set Image = getIconFromTable(strPicture)
        Else
            Set Image = Nothing
        End If
    Else
        Set Image = Nothing
    End If

End Sub

Public Sub GetImages(control As IRibbonControl, ByRef Image)
' Loads an image with transparency to the ribbon
' Modul basGDIPlus is required

    Dim strPicturePath  As String
    Dim strPicture      As String

    strPicture = getTheValue(control.Tag, "CustomPicture")
    Set Image = getIconFromTable(strPicture)

End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
' Set the property "enabled" to a Ribbon Control
' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12

    Select Case control.id
        Case Else
            enabled = True
    End Select

End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
' To set the property "visible" to a Ribbon Control
' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12

    Select Case control.id
        Case Else
            visible = True
    End Select

End Sub

Public Sub GetLabel(control As IRibbonControl, ByRef label)
' Callbackname in XML File "getLabel"
' To set the property "label" to a Ribbon Control

    Select Case control.id
        Case Else
            label = "*getLabel*"
    End Select

End Sub

Public Sub GetScreentip(control As IRibbonControl, ByRef screentip)
' Callbackname in XML File "getScreentip"
' To set the property "screentip" to a Ribbon Control

    Select Case control.id
        Case Else
            screentip = "*getScreentip*"
    End Select

End Sub

Public Sub GetSupertip(control As IRibbonControl, ByRef supertip)
' Callbackname in XML File "getSupertip"
' To set the property "supertip" to a Ribbon Control

    Select Case control.id
        Case Else
            supertip = "*getSupertip*"
    End Select

End Sub

Public Sub GetDescription(control As IRibbonControl, ByRef description)
' Callbackname in XML File "getDescription"
' To set the property "description" to a Ribbon Control

    Select Case control.id
        Case Else
            description = "*getDescription*"
    End Select

End Sub

Public Sub GetTitle(control As IRibbonControl, ByRef title)
' Callbackname in XML File "getTitle"
' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.id
        Case Else
            title = "*getTitle*"
    End Select

End Sub

Public Sub OnActionButton(control As IRibbonControl)
' Callback for event button click
    
    Select Case control.id
        Case Else
            MsgBox "Button """ & control.id & """ clicked!", vbInformation
    End Select

End Sub

'' Command Button
'Public Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
'' Callbackname in XML File Command "onAction"
'' Callback for command event button click
'
'    MsgBox "Button ""Help"" clicked!", vbInformation
'    CancelDefault = True
'
'End Sub
'
'' CheckBox
'Public Sub OnActionCheckBox(control As IRibbonControl, ByRef pressed As Boolean)
'' Callbackname in XML File "OnActionCheckBox"
'' Callback for event checkbox click
'
'    Select Case control.id
'
'        Case Else
'            MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf & _
'                   "Der Wert der Checkbox """ & control.id & """ ist: " & pressed, _
'                   vbInformation
'
'    End Select
'
'End Sub
'
'Public Sub GetPressedCheckBox(control As IRibbonControl, ByRef blnReturn)
'' Callbackname in XML File "GetPressedCheckBox"
'' Callback for checkbox. Indicates how the control is displayed
'
'    Select Case control.id
'        Case Else
'            If getTheValue(control.Tag, "DefaultValue") = "1" Then
'                blnReturn = True
'            Else
'                blnReturn = False
'            End If
'    End Select
'
'End Sub
'
'' ToggleButton
'Public Sub OnActionTglButton(control As IRibbonControl, ByRef pressed As Boolean)
'' Callbackname in XML File "onAction"
'' Callback for a Toggle Buttons click event
'
'    Select Case control.id
'        Case Else
'            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed _
'                    & vbCrLf, vbInformation
'    End Select
'
'End Sub
'
'Public Sub GetPressedTglButton(control As IRibbonControl, ByRef pressed)
'' Callbackname in XML File "getPressed"
'' Callback for an Access ToogleButton Control. Indicates how the control is displayed
'
'    Select Case control.id
'        Case Else
'            If getTheValue(control.Tag, "DefaultValue") = "1" Then
'                pressed = True
'            Else
'                pressed = False
'            End If
'    End Select
'
'End Sub
'
'' EditBox
'Public Sub GetTextEditBox(control As IRibbonControl, ByRef strText)
'' Callbackname in XML File "GetTextEditBox"
'' Callback for an EditBox Control. Indicates which value is to set to the control
'
'    Select Case control.id
'        Case Else
'            strText = getTheValue(control.Tag, "DefaultValue")
'    End Select
'
'End Sub
'
'Public Sub OnChangeEditBox(control As IRibbonControl, strText As String)
'' Callbackname in XML File "OnChangeEditBox"
'' Callback Editbox: Return value of the Editbox
'
'    Select Case control.id
'        Case Else
'            MsgBox "The Value of the EditBox """ & control.id & """ is: " & _
'                        strText & vbCrLf & vbInformation
'    End Select
'
'End Sub
'
'' DropDown
'Public Sub OnActionDropDown(control As IRibbonControl, _
'                ByVal selectedId As String, ByVal selectedIndex As Integer)
'' Callbackname in XML File "OnActionDropDown"
'' Callback onAction (DropDown)
'
'    Select Case control.id
'        Case Else
'            Select Case selectedId
'                Case Else
'                    MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ _
'                            & selectedId & """" & vbCrLf, vbInformation
'            End Select
'    End Select
'
'End Sub
'
'Public Sub GetSelectedItemIndexDropDown(control As IRibbonControl, ByRef index)
'' Callbackname in XML File "GetSelectedItemIndexDropDown"
'' Callback getSelectedItemIndex
'
'    Dim varIndex As Variant
'    varIndex = getTheValue(control.Tag, "DefaultValue")
'
'    If IsNumeric(varIndex) Then
'        Select Case control.id
'            Case Else
'                index = getTheValue(control.Tag, "DefaultValue")
'        End Select
'    End If
'
'End Sub
'
'' Gallery
'Public Sub GetSelectedItemIndexGallery(control As IRibbonControl, ByRef index)
'' Callbackname in XML File "GetSelectedItemIndexGallery"
'' Callback GetSelectedItemIndexGallery
'
'    Dim varIndex As Variant
'    varIndex = getTheValue(control.Tag, "DefaultValue")
'
'    If IsNumeric(varIndex) Then
'        Select Case control.id
'            Case Else
'                index = varIndex
'        End Select
'    End If
'
'End Sub
'
'Public Sub OnActionGallery(control As IRibbonControl, _
'                     ByVal selectedId As String, ByVal selectedIndex As Integer)
'' Callbackname in XML File "OnActionGallery"
'' Callback onAction (Gallery)
'
'    Select Case control.id
'        Case Else
'            Select Case selectedId
'                Case Else
'                    MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & _
'                                selectedId & """" & vbCrLf, vbInformation
'            End Select
'    End Select
'
'End Sub
'
'' Combobox
'Public Sub GetTextComboBox(control As IRibbonControl, ByRef strText)
'' Callbackname im XML File "GetTextComboBox"
'' Callback getText (Combobox)
'
'    Select Case control.id
'        Case Else
'            strText = getTheValue(control.Tag, "DefaultValue")
'    End Select
'
'End Sub
'
'Public Sub OnChangeComboBox(control As IRibbonControl, strText As String)
'' Callbackname in XML File "OnChangeCombobox"
'' Callback onChange (Combobox)
'
'    Select Case control.id
'        Case Else
'            MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & _
'                        strText & """" & vbCrLf, vbInformation
'    End Select
'
'End Sub
'
'' DynamicMenu
'Public Sub GetContent(control As IRibbonControl, ByRef XMLString)
'' Sample for a Ribbon XML "getContent" Callback
'' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
''     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu
'
'    Select Case control.id
'        Case Else
'            XMLString = getXMLForDynamicMenu()
'    End Select
'
'End Sub
'
'' Helper Function
'Public Function getXMLForDynamicMenu() As String
'' Creates XML String for DynamicMenu CallBack - getContent
'
'    Dim lngDummy As Long
'    Dim strDummy As String
'    Dim strContent As String
'    Dim Items(4) As ItemsVal
'
'    Items(0).id = "btnDy1"
'    Items(0).label = "Item 1"
'    Items(0).imageMso = "_1"
'    Items(1).id = "btnDy2"
'    Items(1).label = "Item 2"
'    Items(1).imageMso = "_2"
'    Items(2).id = "btnDy3"
'    Items(2).label = "Item 3"
'    Items(2).imageMso = "_3"
'    Items(3).id = "btnDy4"
'    Items(3).label = "Item 4"
'    Items(3).imageMso = "_4"
'    Items(4).id = "btnDy5"
'    Items(4).label = "Item 5"
'    Items(4).imageMso = "_5"
'
'    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
'
'        For lngDummy = LBound(Items) To UBound(Items)
'            strContent = strContent & _
'                "<button id=""" & Items(lngDummy).id & """" & _
'                " label=""" & Items(lngDummy).label & """" & _
'                " imageMso=""" & Items(lngDummy).imageMso & """" & _
'                " onAction=""OnActionButton""/>" & vbCrLf
'        Next
'
'    strDummy = strDummy & strContent & "</menu>"
'    getXMLForDynamicMenu = strDummy
'
'End Function
'
Public Function getTheValue(strTag As String, strValue As String) As String

    On Error Resume Next

    Dim workTb() As String
    Dim Ele() As String
    Dim myVariabs() As String
    Dim i As Integer

    workTb = Split(strTag, ";")

    ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
    For i = LBound(workTb) To UBound(workTb)
        Ele = Split(workTb(i), ":=")
        myVariabs(i, 0) = Ele(0)
        If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
        End If
    Next

    For i = LBound(myVariabs) To UBound(myVariabs)
        If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
        End If
    Next

End Function

Private Function getIconFromTable(strFileName As String) As Picture

    Dim lngSize As Long
    Dim arrBin() As Byte
    Dim rst As DAO.Recordset

    On Error GoTo PROC_ERR

    Set rst = DBEngine(0)(0).OpenRecordset(mRIBBON_PIX, dbOpenDynaset)
    rst.FindFirst "[FileName]='" & strFileName & "'"
    If rst.NoMatch Then
        Set getIconFromTable = Nothing
    Else
        lngSize = rst.Fields("binary").FieldSize
        ReDim arrBin(lngSize)
        arrBin = rst.Fields("binary").GetChunk(0, lngSize)
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    rst.Close

PROC_EXIT:
    Reset
    Erase arrBin
    Set rst = Nothing
    Exit Function

PROC_ERR:
    Resume PROC_EXIT

End Function

Public Function getPic(strFullPath As String) As String
    Dim strResult As String

    If InStrRev(strFullPath, "\") > 0 Then
        strResult = Mid(strFullPath, InStrRev(strFullPath, "\") + 1)
    Else
        strResult = ""
    End If

    getPic = strResult
End Function