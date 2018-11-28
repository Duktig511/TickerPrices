VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserSelection 
   Caption         =   "Stock Price Analysis"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "UserSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Btn_OK_Click()

    Dim ProceedSelection As Boolean
    Dim ProcessSelection As Boolean
    
    ProceedSelection = Me.Ob_Yes.Value
    
    Select Case ProceedSelection
        Case True
            ProcessSelection = Me.ChBx_ProcessAll.Value
            Unload Me
            Call Main(ProcessSelection)
        Case Else
            Unload Me
        End Select
        
End Sub

Private Sub Ob_No_Click()
    
    If Me.ChBx_ProcessAll.Enabled = True Then
        Me.ChBx_ProcessAll.Value = False
        Me.ChBx_ProcessAll.Enabled = False
    End If
        
End Sub

Private Sub Ob_Yes_Click()
    Me.ChBx_ProcessAll.Enabled = True
End Sub


Private Sub UserForm_Initialize()

End Sub
