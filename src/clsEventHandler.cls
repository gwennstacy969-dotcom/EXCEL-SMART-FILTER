VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' CLASS        : clsEventHandler
' DESCRIPTION  : Event sink untuk kontrol yang dibuat secara dinamis.
'                Menghubungkan event klik, change, dan keydown ke UserForm.
' ==============================================================================

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Public WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1
Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1

Public EventKey As String
Public ParentForm As Object

Private Sub btn_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnCustomButtonClick EventKey
    End If
End Sub

Private Sub opt_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnCustomOptionClick EventKey
    End If
End Sub

Private Sub cbo_Change()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnCustomComboChange EventKey
    End If
End Sub

Private Sub lst_Change()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnCustomListChange EventKey
    End If
End Sub

Private Sub txt_Change()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnCustomTextChange EventKey
    End If
End Sub
