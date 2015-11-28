Option Compare Database
Option Explicit

Public gobjRibbon As IRibbonUI
Private Const mRIBBON_PIX As String = "tblRibbonPix"
'

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Sub GetImages(control As IRibbonControl, ByRef Image)
    Dim strPicturePath  As String
    Dim strPicture      As String
    strPicture = getTheValue(control.Tag, "CustomPicture")
    Set Image = getIconFromTable(strPicture)
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
    Select Case control.id
        Case Else
            enabled = True
    End Select
End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
    Select Case control.id
        Case Else
            visible = True
    End Select
End Sub

Public Sub GetLabel(control As IRibbonControl, ByRef label)
    Select Case control.id
        Case Else
            label = "*getLabel*"
    End Select
End Sub

Public Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    Select Case control.id
        Case Else
            screentip = "*getScreentip*"
    End Select
End Sub

Public Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    Select Case control.id
        Case Else
            supertip = "*getSupertip*"
    End Select
End Sub

Public Sub GetDescription(control As IRibbonControl, ByRef description)
    Select Case control.id
        Case Else
            description = "*getDescription*"
    End Select
End Sub

Public Sub GetTitle(control As IRibbonControl, ByRef title)
    Select Case control.id
        Case Else
            title = "*getTitle*"
    End Select
End Sub

Public Sub OnActionButton(control As IRibbonControl)
    Select Case control.id
        Case Else
            MsgBox "Button """ & control.id & """ clicked!", vbInformation
    End Select
End Sub

Public Function getTheValue(strTag As String, strValue As String) As String

    On Error Resume Next

    Dim workTb() As String
    Dim Ele() As String
    Dim myVariables() As String
    Dim i As Integer

    workTb = Split(strTag, ";")

    ReDim myVariables(LBound(workTb) To UBound(workTb), 0 To 1)
    For i = LBound(workTb) To UBound(workTb)
        Ele = Split(workTb(i), ":=")
        myVariables(i, 0) = Ele(0)
        If UBound(Ele) = 1 Then
            myVariables(i, 1) = Ele(1)
        End If
    Next

    For i = LBound(myVariables) To UBound(myVariables)
        If strValue = myVariables(i, 0) Then
            getTheValue = myVariables(i, 1)
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