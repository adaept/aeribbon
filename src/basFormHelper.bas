Option Compare Database
Option Explicit

' Ref: https://misterslimm.wordpress.com/2007/12/13/microsoft-access-visual-basic-form-helper-setopacity/
' By MISTER SLIMM

Private Const Namespace$ = "basFormHelper"
' Used by BringToTop
Private Declare Function apiBringWindowToTop Lib "user32" Alias "BringWindowToTop" ( _
                         ByVal hWnd As Long) As Long
' Used by SetOpacity and SetTransparentColor
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const WS_EX_LAYERED = &H80000
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                         ByVal hWnd As Long, _
                         ByVal nIndex As Long) As Long
Private Declare Function apiSetLayeredWindowAttributes Lib "user32" Alias "SetLayeredWindowAttributes" ( _
                         ByVal hWnd As Long, _
                         ByVal color As Long, _
                         ByVal AlphaPercent As Byte, _
                         ByVal alpha As Long) As Boolean
Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                         ByVal hWnd As Long, _
                         ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long

Public Sub BringToTop(frm As Access.Form)
' Brings form to top and activates it
    apiBringWindowToTop frm.hWnd
End Sub

Public Function FormOpen(strFormName As String) As Boolean
' Returns True if the form is open in Form View, False if not
' strFormName - The name of the form whose open status we are querying.
' The SysCmd method is used in conjunction with the AllForms.CurrentView method
' so as to prevent the situation that can occur during development whereby my previous
' FormOpen function would return True when the form is open in Design View.

    Debug.Print Namespace$ & "." & "FormOpen"
    On Error GoTo 0

    ' Retrieve object state of form
    Dim varResult As Variant
    Dim blnResult As Boolean
    varResult = SysCmd(acSysCmdGetObjectState, acForm, strFormName)
    blnResult = CBool(varResult)
    Debug.Print , "varResult = " & varResult
    Debug.Print , "blnResult = " & blnResult
    Debug.Print , "acObjStateOpen = " & acObjStateOpen, CBool(acObjStateOpen)
    Debug.Print , "CBool(varResult) And CBool(acObjStateOpen) = " & CStr(CBool(varResult) And CBool(acObjStateOpen))
    If CBool(varResult) And acObjStateOpen Then
        ' Only set to True if the form is in Form View (not Design View)
        FormOpen = (CurrentProject.AllForms(strFormName).CurrentView = acCurViewFormBrowse)
    End If
    Debug.Print , "FormOpen = " & FormOpen

End Function
 
Public Sub SetOpacity(frm As Access.Form, AlphaPercent As Byte)
' Set opacity of form
' Setting AlphaPercent to zero makes the form fully transparent.
' Setting AlphaPercent to 100 makes the form fully opaque.
' This only has an affect on forms whose PopUp property is True.

    ' Perform checks on arguments
    ' Ensure frm is a PopUp form. Raise an error if it is not.
    If Not frm.PopUp Then
        Err.Raise 5, , "Invalid argument." & vbCrLf & "Cannot SetOpacity on form " & frm.Name & ". PopUp form required."
        Exit Sub
    End If

    ' Ensure AlphaPercent is between 0 and 100.
    ' Do not raise an error if out of range, simply fix it.
    If AlphaPercent < 0 Then
        AlphaPercent = 0
    ElseIf AlphaPercent > 100 Then
        AlphaPercent = 100
    End If

    ' --------------------------------------------------
    ' If we reach here, all arguments have been accepted
    ' --------------------------------------------------
    
    ' Convert supplied percentage value to one ranging between 0 and 255 for apiSetLayeredWindowAttributes
    Dim iAlpha As Integer
    iAlpha = (AlphaPercent / 100) * 255

    ' Get forms current extended attributes
    Dim attrib As Long
    attrib = apiGetWindowLong(frm.hWnd, GWL_EXSTYLE)
    ' Set form to have extended layered attribute
    apiSetWindowLong frm.hWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    ' Set opacity
    apiSetLayeredWindowAttributes frm.hWnd, RGB(0, 0, 0), iAlpha, LWA_ALPHA
End Sub

Public Function OpenForm( _
                FormName As String, _
                Optional View As AcFormView = acNormal, _
                Optional FilterName As String, _
                Optional WhereCondition As String, _
                Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
                Optional WindowMode As AcWindowMode = acWindowNormal, _
                Optional OpenArgs As String)
' Opens form or, if open already, brings form to the top and activates it. Is an exact replacement for DoCmd.OpenForm.
' There are some significant limitations on passing arguments to forms that are already open.
' View, FilterName, WindowMode and OpenArgs will all be ignored.
' If these arguments are supplied, the form will be closed and re-opened using DoCmd.OpenForm.
  
    ' Determine whether we must use DoCmd.OpenForm
    Dim DoCmdOpenFormRequired As Boolean
    DoCmdOpenFormRequired = (View <> acNormal) Or (FilterName <> "") Or (WindowMode <> acWindowNormal) Or (OpenArgs <> "")
  
    If FormOpen(FormName) And Not DoCmdOpenFormRequired Then
  
        ' Bring open form to top
        BringToTop Forms(FormName)
    
        With Forms(FormName)
            ' Update Filter property using WhereCondition
            If WhereCondition <> .Filter Then
                .Filter = WhereCondition
                .FilterOn = (WhereCondition <> "")
            End If
      
            Select Case DataMode
                Case acFormAdd: .AllowAdditions = True: .AllowDeletions = False: .AllowEdits = False
                Case acFormEdit: .AllowAdditions = True: .AllowDeletions = True: .AllowEdits = True
                Case acFormReadOnly: .AllowAdditions = False: .AllowDeletions = False: .AllowEdits = False
            End Select
      
        End With
    Else
        ' Open form in standard way using DoCmd.OpenForm
        DoCmd.OpenForm FormName, View, FilterName, WhereCondition, DataMode, WindowMode, OpenArgs
    End If

End Function