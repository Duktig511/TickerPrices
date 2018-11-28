Attribute VB_Name = "FrmRoutines"
Option Explicit

Sub ShowForm()
Attribute ShowForm.VB_ProcData.VB_Invoke_Func = "w\n14"
    
    Dim CurrentWS As String
    Dim FormPrompt As String
    
    FormPrompt = "You Are About to Process Summary Info for Stock Price Transactions." _
    & vbCrLf _
    & "Would You Like to Continue?"
    
    Load UserSelection
    
    With UserSelection
        .Lbl_ContinuePrompt.Caption = FormPrompt
        .Ob_Yes.Value = False
        .Ob_No = True
        .ChBx_ProcessAll.Enabled = False
        .Show
    End With

End Sub


Sub RemovePrTCollection()
    
    Set PriceTransactions = Nothing

End Sub
