Attribute VB_Name = "Module1"
Public prmUsername As String
Public prmPassword As String
Public prmBeginningNextMonth As String
Public prmBeginningCurrentMonth As String
Public prmEndCurrentMonth As String
Public prmFiscalYear As String
Public blnKeepProcessing As Boolean



Private Sub FixedAssetsQT()

     Dim sConn As String
     Dim sSql As String
     Dim sSql_2 As String
     Dim oQt As QueryTable
     
'   Show Userform
     frmDBLogin.Show
     
'   Establish an ODBC connection with Oracle Database
     sConn = "ODBC;DSN=PROD;"
     sConn = sConn & "UID=" & prmUsername & ";PWD=" & prmPassword & ";"
     sConn = sConn & "SERVER=PROD"
        
     sSql = "with MAX_CHANGE_SEQ_NUM as "
     sSql = sSql & "(select ffrmasa_otag_code OTAG, max(ffrmasa_change_seq_num) CHANGE_SEQ_NUM from ffrmasa group by ffrmasa_otag_code) "
     sSql = sSql & "select ffrmasa_acct_code_asset ACCT,"
     sSql = sSql & "ffbmast_ptag_code PTAG,"
     sSql = sSql & "ffbmast_asty_code ASSET_TYPE,"
     sSql = sSql & "ffbmast_asset_desc ASSET_DESC,"
     sSql = sSql & "to_char(ffbmast_acqd_date,'MM/DD/YYYY') ACQ_DATE,"
     sSql = sSql & "ffrmasa_amt TOTAL_COST, "
     sSql = sSql & "ffrmasa_adj_amt ADJUSTMENT, "
     sSql = sSql & "(nvl(ffrmasa_adj_amt,0) + ffrmasa_amt) ADJ_COST, "
     sSql = sSql & "ffrmasa_depr_accum_amt ACCUM_DEPR "
     sSql = sSql & "from MAX_CHANGE_SEQ_NUM, "
     sSql = sSql & "ffbmast join ffrmasa on ffbmast_otag_code = ffrmasa_otag_code "
     sSql = sSql & "where (ffbmast_disp_date >= '" & prmBeginningNextMonth & "' "
     sSql = sSql & "or ffbmast_disp_date is null) and "
     sSql = sSql & "trunc(ffbmast_acqd_date) <= '" & prmEndCurrentMonth & "' and "
     sSql = sSql & "ffrmasa_otag_code = OTAG and "
     sSql = sSql & "ffrmasa_change_seq_num = CHANGE_SEQ_NUM and "
     sSql = sSql & "ffrmasa_acct_code_asset = '17400' and "
     sSql = sSql & "ffbmast_ptag_code is not null "
     sSql = sSql & "order by ffrmasa_acct_code_asset, "
     sSql = sSql & "ffbmast_ptag_code; "
     
    With ActiveSheet.QueryTables.Add( _
         Connection:=sConn, _
         Destination:=ActiveCell)
            .CommandText = sSql
            .Refresh BackgroundQuery:=False
    End With

 End Sub
 

Private Sub OperatingListingQT()


     Dim sConn As String
     Dim sSql As String
     Dim oQt As QueryTable
          
'   Show Userform
     frmDBLogin2.Show
     
'   Establish an ODBC connection with Oracle Database
     sConn = "ODBC;DSN=PROD;"
     sConn = sConn & "UID=" & prmUsername & ";PWD=" & prmPassword & ";"
     sConn = sConn & "SERVER=PROD"
     
'   Query
     sSql = "select to_char(fgbtrnh_trans_date, 'MM/DD/YYYY') Transaction_, "
     sSql = sSql & "fgbtrnh_rucl_code Rule_, "
     sSql = sSql & "fgbtrnh_doc_code Document_, "
     sSql = sSql & "fgbtrnh_acci_code Index_, "
     sSql = sSql & "fgbtrnh_fund_code Fund_, "
     sSql = sSql & "(CASE WHEN fgbtrnh_fund_code BETWEEN '10000' AND '19999' THEN '11' "
     sSql = sSql & "WHEN fgbtrnh_fund_code BETWEEN '20000' AND '29999' THEN '21' "
     sSql = sSql & "WHEN fgbtrnh_fund_code BETWEEN '30000' AND '39999' THEN '31' "
     sSql = sSql & "WHEN fgbtrnh_fund_code BETWEEN '91000' AND '92999' THEN '91' "
     sSql = sSql & "WHEN fgbtrnh_fund_code BETWEEN '93000' AND '94000' THEN '93' "
     sSql = sSql & "ELSE '??' "
     sSql = sSql & "END) Type_, "
     sSql = sSql & "fgbtrnh_acct_code Account_, "
     sSql = sSql & "fgbtrnh_dr_cr_ind DR_CR, "
     sSql = sSql & "(CASE WHEN fgbtrnh_dr_cr_ind = 'C' AND fgbtrnh_rucl_code IN ('XC5','ICEI','CNEI','INNC','INEC') THEN fgbtrnh_trans_amt "
     sSql = sSql & "WHEN fgbtrnh_dr_cr_ind != 'C' AND fgbtrnh_rucl_code IN ('XC5','ICEI','CNEI','INNC','INEC') THEN (fgbtrnh_trans_amt * -1) "
     sSql = sSql & "WHEN fgbtrnh_dr_cr_ind = 'C' AND fgbtrnh_rucl_code NOT IN ('XC5','ICEI','CNEI','INNC','INEC') THEN (fgbtrnh_trans_amt * -1) "
     sSql = sSql & "WHEN fgbtrnh_dr_cr_ind != 'C' AND fgbtrnh_rucl_code NOT IN ('XC5','ICEI','CNEI','INNC','INEC') THEN fgbtrnh_trans_amt "
     sSql = sSql & "ELSE 0 "
     sSql = sSql & "END) Amount_, "
     sSql = sSql & "fgbtrnh_trans_desc Description, "
     sSql = sSql & "fabinvh_pohd_code PO "
     sSql = sSql & "FROM fgbtrnh "
     sSql = sSql & "LEFT JOIN fabinvh "
     sSql = sSql & "ON fgbtrnh_doc_code=fabinvh_code "
     sSql = sSql & "where fgbtrnh_fsyr_code = '" & prmFiscalYear & "' and "
     sSql = sSql & "fgbtrnh_coas_code = 'M' and "
     sSql = sSql & "fgbtrnh_acct_code between '78110' AND '78190' AND "
     sSql = sSql & "trunc(fgbtrnh_trans_date) BETWEEN '" & prmBeginningCurrentMonth & "' AND '" & prmEndCurrentMonth & "' AND "
     sSql = sSql & "fgbtrnh_rucl_code IN ('INEI','ICEI','INNI','FT01','INNC','INCI','CNEI','X25','XCD','XC5','ICNI','INEC','XPC','ICEC') "
     sSql = sSql & "ORDER BY Transaction_;"

    With ActiveSheet.QueryTables.Add( _
         Connection:=sConn, _
         Destination:=ActiveCell)
            .CommandText = sSql
            .Refresh BackgroundQuery:=False
    End With
    
 End Sub
 
Private Sub GeneralLedgerListingQT()


     Dim sConn As String
     Dim sSql As String
     Dim oQt As QueryTable
     
'   Show Userform
     frmDBLogin4.Show
     
'   Establish an ODBC connection with SQL SERVER Database
     sConn = "ODBC;DSN=banner_report;UID=" & prmUsername & ";Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=D003310ACCTSVCS"
     
'   Query
'   (AF_TRANSACTION_DETAIL.DETAIL_RUCL_CODE In ('SCAP', 'SCAO', 'WRIT')) AND
     sSql = "SELECT convert(varchar(10),AF_TRANSACTION_DETAIL.TRANSACTION_DATE,101) TRANSACTION_DATE, AF_TRANSACTION_DETAIL.DETAIL_RUCL_CODE, AF_TRANSACTION_DETAIL.DOC_CODE_KEY, AF_TRANSACTION_DETAIL.FUND_CODE, (CASE WHEN AF_TRANSACTION_DETAIL.DEBIT_CREDIT_IND In ('C','-') THEN AF_TRANSACTION_DETAIL.TRANSACTION_AMOUNT*(-1) ELSE AF_TRANSACTION_DETAIL.TRANSACTION_AMOUNT END) TRANS_AMT, AF_TRANSACTION_DETAIL.TRANSACTION_DESC "
     sSql = sSql & "FROM DPDWMT.dbo.AF_TRANSACTION_DETAIL AF_TRANSACTION_DETAIL "
     sSql = sSql & "WHERE (AF_TRANSACTION_DETAIL.TRANSACTION_DATE>={ts '" & prmBeginningCurrentMonth & " 00:00:00'} And AF_TRANSACTION_DETAIL.TRANSACTION_DATE<={ts '" & prmBeginningNextMonth & " 00:00:00'}) AND (AF_TRANSACTION_DETAIL.FUND_CODE='974000') AND (AF_TRANSACTION_DETAIL.FSYR_CODE='" & prmFiscalYear & "') AND (AF_TRANSACTION_DETAIL.COAS_CODE='M') AND AF_TRANSACTION_DETAIL.ACCT_CODE In ('17400') "
     sSql = sSql & "ORDER BY AF_TRANSACTION_DETAIL.TRANSACTION_DATE"
     
    With ActiveSheet.QueryTables.Add( _
         Connection:=sConn, _
         Destination:=ActiveCell)
            .CommandText = sSql
            .Refresh BackgroundQuery:=False
    End With
    
 End Sub
 
 Private Sub SendAsPDF()
'   Uses early binding
'   Requires a reference to the Outlook Object Library
    Dim OutlookApp As Outlook.Application
    Dim MItem As Object
    Dim Recipient As String, Subj As String
    Dim Msg As String, Fname As String
            
'   Message details
    Recipient = "ben.jones@mtsu.edu"
    Subj = "Signed Equipment Reconciliation"
    Msg = "Hi Ben," & vbNewLine & vbNewLine & " here's the signed reconciliation for the month." & vbNewLine & _
    "See the folder Z:\Alex Simonian\FY15_Equipment\_Signed_Reports"
    Msg = Msg & vbNewLine & vbNewLine & "-Alex"
    Fname = Application.DefaultFilePath & "\" & _
      ActiveWorkbook.Name & ".pdf"
   
'   Create the attachment
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=Fname
    
'   Create Outlook object
    Set OutlookApp = New Outlook.Application
    
'   Create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(olMailItem)
    With MItem
      .To = Recipient
      .Subject = Subj
      .Body = Msg
      .Attachments.Add Fname
      .Save 'to Drafts folder
      '.Send
    End With
    Set OutlookApp = Nothing

'   Delete the file
    Kill Fname
End Sub

'Callback for customButton1 onAction
Sub Macro1(control As IRibbonControl)
    FixedAssetsQT
End Sub

'Callback for customButton2 onAction
Sub Macro2(control As IRibbonControl)
    OperatingListingQT
End Sub

'Callback for customButton3 onAction
Sub Macro3(control As IRibbonControl)
    GeneralLedgerListingQT
End Sub

'Callback for customButton4 onAction
Sub Macro4(control As IRibbonControl)
    SendAsPDF
End Sub

'Callback for customButton5 onAction
Sub Macro5(control As IRibbonControl)
    MsgBox "This is macro 5"
End Sub

'Callback for customButton6 onAction
Sub Macro6(control As IRibbonControl)
    MsgBox "This is macro 6"
End Sub

'Callback for customButton7 onAction
Sub Macro7(control As IRibbonControl)
    MsgBox "This is macro 7"
End Sub

'Callback for customButton8 onAction
Sub Macro8(control As IRibbonControl)
    MsgBox "This is macro 8"
End Sub

    
