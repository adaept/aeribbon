Option Compare Database
Option Explicit

Public gobjRibbon As IRibbonUI
Private Const mRIBBON_TABLE As String = "tblRibbonPix"
Private Const mIMAGE_FIELD As String = "Image"
Private pixClass As aeGDayClass

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Function SetImage(ByVal rcontrol As IRibbonControl, ByRef pic As Variant) As Boolean
    'MsgBox "rcontrol.Id = " & rcontrol.Id, vbInformation, gconTHIS_APP_NAME
    On Error GoTo 0
    Set pixClass = New aeGDayClass
    Select Case rcontrol.Id
        Case "btn1"
            Set pic = pixClass.aeAttachmentToPicture(mRIBBON_TABLE, mIMAGE_FIELD, "adaept32only.ico")
        'Case "btn2"
            'Set pic = pixClass.aeAttachmentToPicture(mRIBBON_TABLE, mIMAGE_FIELD, "something.png")
        'Case "btn3"
            'Set pic = pixClass.aeAttachmentToPicture(mRIBBON_TABLE, mIMAGE_FIELD, "theotherthing.ico")
        Case Else
            MsgBox "Bad SetImage Case!", vbCritical, gconTHIS_APP_NAME
    End Select
End Function

Public Sub GetEnabled(ByVal rcontrol As IRibbonControl, ByRef enabled)
    Select Case rcontrol.Id
        Case Else
            enabled = True
    End Select
End Sub

Public Sub GetVisible(ByVal rcontrol As IRibbonControl, ByRef visible)
    Select Case rcontrol.Id
        Case Else
            visible = True
    End Select
End Sub

Public Sub GetLabel(ByVal rcontrol As IRibbonControl, ByRef label)
    Select Case rcontrol.Id
        Case Else
            label = "*getLabel*"
    End Select
End Sub

Public Sub GetScreentip(ByVal rcontrol As IRibbonControl, ByRef screentip)
    Select Case rcontrol.Id
        Case Else
            screentip = "*getScreentip*"
    End Select
End Sub

Public Sub GetSupertip(ByVal rcontrol As IRibbonControl, ByRef supertip)
    Select Case rcontrol.Id
        Case Else
            supertip = "*getSupertip*"
    End Select
End Sub

Public Sub GetDescription(ByVal rcontrol As IRibbonControl, ByRef description)
    Select Case rcontrol.Id
        Case Else
            description = "*getDescription*"
    End Select
End Sub

Public Sub GetTitle(ByVal rcontrol As IRibbonControl, ByRef title)
    Select Case rcontrol.Id
        Case Else
            title = "*getTitle*"
    End Select
End Sub

Private Function IsOpen(ByVal strFormName As String) As Boolean
    IsOpen = False
    ' Is form open?
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        ' If so make sure it is not in design view
        If Forms(strFormName).CurrentView <> 0 Then
            IsOpen = True
        End If
    End If
    Exit Function
 End Function

Public Sub OnActionButton(ByVal rcontrol As IRibbonControl)
    Select Case rcontrol.Id
        Case "btn1"
            DoEvents
            If IsOpen("frmSplash") Then
                MsgBox "frmSplash is already open!", vbCritical, gconTHIS_APP_NAME
            Else
                DoCmd.OpenForm "frmSplash"
            End If
        Case Else
            MsgBox "Button """ & rcontrol.Id & """ clicked!", vbInformation, gconTHIS_APP_NAME
    End Select
End Sub

Public Function getTheValue(ByVal strTag As String, ByVal strValue As String) As String

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