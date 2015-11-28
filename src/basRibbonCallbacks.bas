Option Compare Database
Option Explicit

Public gobjRibbon As IRibbonUI

Private Const mRIBBON_PIX As String = "tblRibbonPix"

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