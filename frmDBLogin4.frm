VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBLogin4 
   Caption         =   "Set General Ledger Parameters"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDBLogin4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDBLogin4"
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
      prmBeginningCurrentMonth = BeginDate(cboDate.Text, cboYear.Text)
      prmBeginningNextMonth = EndDate(cboDate.Text, cboYear.Text)
      prmFiscalYear = cboFY.Text
      MsgBox prmBeginningCurrentMonth
      MsgBox prmBeginningNextMonth
      MsgBox prmFiscalYear
      blnKeepProcessing = True
      Unload Me
   End If
End Sub
   
   Private Function BeginDate(Month As String, Year As Integer) As String
   
        Select Case Month
            Case "JAN"
                BeginDate = Year & "-01-01"
            Case "FEB"
                BeginDate = Year & "-02-01"
            Case "MAR"
                BeginDate = Year & "-03-01"
            Case "APR"
                BeginDate = Year & "-04-01"
            Case "MAY"
                BeginDate = Year & "-05-01"
            Case "JUN"
                BeginDate = Year & "-06-01"
            Case "JUL"
                BeginDate = Year & "-07-01"
            Case "AUG"
                BeginDate = Year & "-08-01"
            Case "SEP"
                BeginDate = Year & "-09-01"
            Case "OCT"
                BeginDate = Year & "-10-01"
            Case "NOV"
                BeginDate = Year & "-11-01"
            Case "DEC"
                BeginDate = Year & "-12-01"
        End Select
   
   End Function
   Private Function EndDate(Month As String, Year As Integer) As String
   
        Select Case Month
            Case "JAN"
                EndDate = Year & "-02-01"
            Case "FEB"
                EndDate = Year & "-03-01"
            Case "MAR"
                EndDate = Year & "-04-01"
            Case "APR"
                EndDate = Year & "-05-01"
            Case "MAY"
                EndDate = Year & "-06-01"
            Case "JUN"
                EndDate = Year & "-07-01"
            Case "JUL"
                EndDate = Year & "-08-01"
            Case "AUG"
                EndDate = Year & "-09-01"
            Case "SEP"
                EndDate = Year & "-10-01"
            Case "OCT"
                EndDate = Year & "-11-01"
            Case "NOV"
                EndDate = Year & "-12-01"
            Case "DEC"
                EndDate = (Year + 1) & "-01-01"
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
        .AddItem "2014"
        .AddItem "2015"
        .AddItem "2016"
        .AddItem "2017"
        .AddItem "2018"
        .AddItem "2019"
    End With
    
    With Me.cboFY
        .AddItem "2014"
        .AddItem "2015"
        .AddItem "2016"
        .AddItem "2017"
        .AddItem "2018"
        .AddItem "2019"
    End With
   
End Sub

