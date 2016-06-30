VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBLogin2 
   Caption         =   "Set Oracle Login Parameters Operating Listing"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDBLogin2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDBLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCancel_Click()
   'User clicked the [Cancel] button
   blnKeepProcessing = False
   Unload Me
End Sub

Private Sub cmdSaveLoginParams_Click()
   Dim strMsgText As String
   
   strMsgText = ""
   If tbxEBUsername.Text = "" Then
      strMsgText = strMsgText & "-Invalid or Missing Username." & vbCr
   End If
   
   If tbxEBPassword.Text = "" Then
      strMsgText = strMsgText & "-Missing Password." & vbCr
   End If
   
   If cboDate.Text = "" Then
      strMsgText = strMsgText & "-Missing Month." & vbCr
   End If
   
   If cboYear.Text = "" Then
      strMsgText = strMsgText & "-Missing Year." & vbCr
   End If
   
   If cboFY.Text = "" Then
      strMsgText = strMsgText & "-Missing Fiscal Year." & vbCr
   End If

   If Len(strMsgText) > 0 Then
      'User inputs are invalid and must be corrected
      strMsgText = "You must correct the following errors:" & vbCr & strMsgText
      MsgBox strMsgText
      blnKeepProcessing = False
      Exit Sub
   Else
      'Inputs are valid: Set global variables
      prmUsername = tbxEBUsername.Text
      prmPassword = tbxEBPassword.Text
      prmBeginningCurrentMonth = BeginDate(cboDate.Text, cboYear.Text)
      prmEndCurrentMonth = EndDate(cboDate.Text, cboYear.Text)
      prmFiscalYear = cboFY.Text
      MsgBox prmBeginningCurrentMonth
      MsgBox prmEndCurrentMonth
      MsgBox prmFiscalYear
      blnKeepProcessing = True
      Unload Me
   End If
End Sub
   
   Private Function BeginDate(Month As String, Year As Integer) As String
   
        Select Case Month
            Case "JAN"
                BeginDate = "1-JAN-" & Year
            Case "FEB"
                BeginDate = "1-FEB-" & Year
            Case "MAR"
                BeginDate = "1-MAR-" & Year
            Case "APR"
                BeginDate = "1-APR-" & Year
            Case "MAY"
                BeginDate = "1-MAY-" & Year
            Case "JUN"
                BeginDate = "1-JUN-" & Year
            Case "JUL"
                BeginDate = "1-JUL-" & Year
            Case "AUG"
                BeginDate = "1-AUG-" & Year
            Case "SEP"
                BeginDate = "1-SEP-" & Year
            Case "OCT"
                BeginDate = "1-OCT-" & Year
            Case "NOV"
                BeginDate = "1-NOV-" & Year
            Case "DEC"
                BeginDate = "1-DEC-" & Year
        End Select
   
   End Function
   Private Function EndDate(Month As String, Year As Integer) As String
   
        Select Case Month
            Case "JAN"
                EndDate = "31-JAN-" & Year
            Case "FEB"
                EndDate = "28-FEB-" & Year
            Case "MAR"
                EndDate = "31-MAR-" & Year
            Case "APR"
                EndDate = "30-APR-" & Year
            Case "MAY"
                EndDate = "31-MAY-" & Year
            Case "JUN"
                EndDate = "30-JUN-" & Year
            Case "JUL"
                EndDate = "31-JUL-" & Year
            Case "AUG"
                EndDate = "31-AUG-" & Year
            Case "SEP"
                EndDate = "30-SEP-" & Year
            Case "OCT"
                EndDate = "31-OCT-" & Year
            Case "NOV"
                EndDate = "30-NOV-" & Year
            Case "DEC"
                EndDate = "31-DEC-" & Year
        End Select
   
   End Function
   
   Public Sub UserForm_Initialize()
    With Me.cboDate
        .AddItem "JAN"
        .AddItem "FEB"
        .AddItem "MAR"
        .AddItem "APR"
        .AddItem "MAY"
        .AddItem "JUN"
        .AddItem "JUL"
        .AddItem "AUG"
        .AddItem "SEP"
        .AddItem "OCT"
        .AddItem "NOV"
        .AddItem "DEC"
    End With
    
    With Me.cboYear
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .AddItem "18"
        .AddItem "19"
    End With
    
    With Me.cboFY
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .AddItem "18"
        .AddItem "19"
    End With
   
End Sub

