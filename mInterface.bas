Attribute VB_Name = "mInterface"
'---------------------------------------------------------------------------------------
' Module    : mInterface
' DateTime  : 12/16/2004 21:24
' Author    : Shane Mulligan
' Purpose   : Implements the askbox
'---------------------------------------------------------------------------------------

Option Explicit

Public AskBoxReturnVal As Integer

Public Function AskBox(ByRef AskingForm As Form, ByVal aCaption As String, ByVal aTitle As String, Optional ByVal aStyle As VbMsgBoxStyle = vbOKOnly, Optional ByVal XEscape As Boolean = False) As VbMsgBoxResult
      
   AskingForm.Enabled = False
   
   With frmAskBox
   
      If Not XEscape Then
         .cmdExitEx.Visible = False
         .Shape1.Visible = False
         .Line1.Visible = False
         .Line2.Visible = False
      End If
      
      Select Case aStyle
      Case vbOKOnly
         .cmd(2).Visible = True
         .shp(2).Visible = True
      Case vbYesNo
         .cmd(0).Visible = True
         .cmd(1).Visible = True
         .shp(0).Visible = True
         .shp(1).Visible = True
      End Select
      
      .Header.Caption = aTitle
      .Header.Alignment = 2
      .lblMessage.Caption = aCaption
      .Show
      
      AskBoxReturnVal = -1
      Do
         DoEvents
      Loop Until AskBoxReturnVal <> -1 Or .Visible = False
      
      Select Case AskBoxReturnVal
      Case 0
         AskBox = vbNo
      Case 1
         AskBox = vbYes
      Case 2
         AskBox = vbOKOnly
      Case Else
         AskBox = vbAbort
      End Select
      
   End With
   
   AskingForm.Enabled = True
   
   Unload frmAskBox

End Function
