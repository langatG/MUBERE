Attribute VB_Name = "DataBase"
Public Sub ModifyDatabase()
    Dim CnDatabase As New Connection
    Dim RsDatabase As New Recordset
    Dim rsLoanBal As New Recordset
    On Error GoTo ErrorTrap
    With CnDatabase
        If .State = adStateClosed Then
            .Open SelectedDsn, "atm"
        End If
        '// this can only be used in access data base
        .Execute "CREATE TABLE APPRAISAL(LoanNo TEXT(50) NULL,AppraisDate DATETIME," _
        & "Salary CURRENCY,Allowances CURRENCY,RepayMethod TEXT,[Co-opShares] CURRENCY," _
        & "[Co-opLoans] CURRENCY,Shares CURRENCY,Loans CURRENCY,Deductions CURRENCY," _
        & "AmtRecommended CURRENCY,Reason TEXT,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE LOANVAR (LoanNo TEXT(20),MemberNo TEXT(50),OldRate " _
        & "CURRENCY,NewRate CURRENCY,VarDate DATETIME,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE SAVINGS (MemberNo TEXT(50),TransDate DATETIME," _
        & "TransNo NUMBER,TransName TEXT,TransBy TEXT,ChequeNo TEXT,Amount CURRENCY," _
        & "ReceiptNo TEXT,IncTrans TEXT,Locked TEXT(3),AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE TEMPMEMBERSTATEMENT (MemberNo TEXT(50),RefNo " _
        & "TEXT,Description TEXT,TransCode NUMBER,Principal CURRENCY,Interest CURRENCY," _
        & "MonthlyContr CURRENCY,Total CURRENCY)"
    
        .Execute "CREATE TABLE TMPDIVIDENDPAYLIST (MemberNo TEXT(50)," _
        & "Current_Tot_Shares CURRENCY,Shares_As_At CURRENCY,Gross_Dividend " _
        & "CURRENCY,WithTax CURRENCY,Net_Dividend CURRENCY,BankName TEXT,AcctNo" _
        & " TEXT,CompanyName TEXT,ShareInterest NUMBER,WithHoldingTax NUMBER)"
    
        .Execute "CREATE TABLE TMPPERIODICTRANS (MemberNo TEXT(50)," _
        & "Names TEXT,OpeningBal CURRENCY,NewLoans CURRENCY,TotalLoanRepayment" _
        & " CURRENCY,TotalLoanInterest CURRENCY,LoanOutstanding CURRENCY," _
        & "TotMonthContr CURRENCY,SharesContr CURRENCY,SharesAsAt CURRENCY," _
        & "Period TEXT)"
        
        .Execute "CREATE TABLE TMPSTATEMENT (MemberNo TEXT(50),Period " _
        & "TEXT,OpeningBalance CURRENCY,NewLoan CURRENCY,LoanInterestPaid CURRENCY" _
        & ",LoanInterestCharged CURRENCY,LoanInterestOwing CURRENCY,LoanRepayment " _
        & "CURRENCY,OutstandingLoanBalance CURRENCY,OpeningShares CURRENCY," _
        & "SharesContributed CURRENCY,ClosingShares CURRENCY," _
        & "TotalMonthlyContribution CURRENCY)"
        
        .Execute "ALTER TABLE KIN ADD COLUMN Percentage NUMBER,Witness TEXT(50)"
        
        .Execute "ALTER TABLE LOANS ADD COLUMN jOBJRP TEXT(50),pupose TEXT(50)"
        
        .Execute "ALTER TABLE REPAY ADD COLUMN ReceiptNo TEXT(50)"
        
        .Execute "ALTER TABLE CONTRIB ADD COLUMN ReceiptNo TEXT(50)"
        
        .Execute "ALTER TABLE BFDEDUCT ADD COLUMN PostingID COUNTER"
        
        .Execute "ALTER TABLE LOANBAL ADD COLUMN LoanCode TEXT(10)"
        
        .Execute "ALTER TABLE USERGRPS ADD COLUMN RWOtherSchemes TEXT(5),MemReg TEXT(5),NOK TEXT(5)," _
        & "MemStatement TEXT(5),Contributions TEXT(5),ShareVar TEXT(5)," _
        & "LoanApplic TEXT(5),LoanEndorsement TEXT(5),ChequeEntry TEXT(5)," _
        & "LoanBal TEXT(5),LoanTrans TEXT(5),EffectRepayment TEXT(5)," _
        & "LoanGuarantors TEXT(5),Dividends TEXT(5),Deductions TEXT(5)," _
        & "PeriodicTrans TEXT(5),UtilStatements TEXT(5),MonthlyDeductions " _
        & "TEXT(5),ExportToGL TEXT(5),UtilGuarantor TEXT(5),Dormant TEXT(5)," _
        & "Archived TEXT(5),Withdrawn TEXT(5),BackUp TEXT(5),Savings TEXT(5)" _
        & ",BenevolentFund TEXT(5),Parametization TEXT(5),LoanTypes TEXT(5)," _
        & "CompanySetUp TEXT(5),BankSetUp TEXT(5),DatabaseSetUp TEXT(5)," _
        & "Activate TEXT(5),RejReasons TEXT(5),UserGrps TEXT(5),SysUsers " _
        & "TEXT(5),ClearLoanRecs TEXT(5),ClearMemRecs TEXT(5),ChangeMemNo TEXT(5)"
        
    End With
    Set RsDatabase = oSaccoMaster.GetRecordset("select LoanNo,LoanCode from " _
    & "Loans order by LoanNo")
    With RsDatabase
        If Not .EOF Then
            While Not .EOF
                Set rsLoanBal = oSaccoMaster.GetRecordset("select * from" _
                & " LOANBAL where LoanNo='" & RsDatabase!Loanno & "'")
                With rsLoanBal
                    If Not .EOF Then
                        !Loancode = RsDatabase!Loancode
                        .Update
                    End If
                End With
                .MoveNext
            Wend
        End If
    End With
    
    CnDatabase.Execute "Update LOANBAL set LoanCode='BAL' where isnull(LoanCode)"
    
    CnDatabase.Execute "INSERT INTO LOANTYPE (LoanCode,LoanType,RepayPeriod,Interest" _
    & ",MaxAmount,Guarantor,AuditID,AuditTime) VALUES ('BAL','BALANCE',60,12," _
    & "1000000,'No','ADMIN',#" & Now & "#)"
    
    Exit Sub
ErrorTrap:
    MsgBox err.description, , "Modifying Database"
End Sub
Public Sub CreateDatabase()
    Dim CnDatabase As New Connection
    Dim RsDatabase As New Recordset
    On Error GoTo ErrorTrap
    With CnDatabase
        If .State = adStateClosed Then
            .Open SelectedDsn, ""
        End If
        .Execute "CREATE TABLE APPRAISAL(LoanNo TEXT(50) NULL,AppraisDate DATETIME," _
        & "Salary CURRENCY,Allowances CURRENCY,RepayMethod TEXT,[Co-opShares] CURRENCY," _
        & "[Co-opLoans] CURRENCY,Shares CURRENCY,Loans CURRENCY,Deductions CURRENCY," _
        & "AmtRecommended CURRENCY,Reason TEXT,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE ARCHIVED(MemberNo TEXT,DateArchived DATETIME," _
        & "TotalShares CURRENCY,Reasons TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE BANKS (BankCode TEXT,BankName TEXT,BranchName " _
        & "TEXT,Address TEXT,Telephone TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE BFDEDUCT (PostingID COUNTER,FundID TEXT," _
        & "MemberNo TEXT,TransDate DATETIME,RefNo NUMBER,Amount CURRENCY,Posted TEXT," _
        & "Remarks TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE CHEQUES (LoanNo TEXT,ChequeNo TEXT,Amount " _
        & "CURRENCY, CollectorID TEXT,CollectName TEXT,DateIssued DATETIME,ClerkStaffNo" _
        & " TEXT,ClerkName TEXT,Status TEXT,Reasons TEXT,AuditID TEXT,AuditTime DATETIME," _
        & "Remarks TEXT)"
    
        .Execute "CREATE TABLE COMPANY (CompanyCode TEXT,CompanyName TEXT," _
        & "Telephone TEXT,Address TEXT,AccountNo TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE CONTRIB (MemberNo TEXT,Contrdate DATETIME,RefNo " _
        & "INTEGER,Amount CURRENCY,ShareBal CURRENCY,TransBy TEXT,ChequeNo TEXT,ReceiptNo" _
        & " TEXT,Locked TEXT,Posted TEXT,Remarks TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE DORMANT (MemberNo TEXT,Name TEXT,LastContrDate" _
        & " DATETIME,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE ENDMAIN (LoanNo TEXT,MinuteNo TEXT,MeetingDate" _
        & " DATETIME,AmtApproved CURRENCY,Accepted TEXT,ChairSigned TEXT,SecSigned TEXT," _
        & "MembSigned TEXT,Reasons TEXT,Remarks TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE GUARANTO (MemberNo TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE KIN (MemberNo TEXT(50),KinNames TEXT,KinNo TEXT," _
        & "Address TEXT,IDNo TEXT,Relationship TEXT,HomeTelNo TEXT,Witness TEXT,Percentage" _
        & " NUMBER,OfficeTelNo TEXT,SignDate DATETIME,KinSigned TEXT,AuditID TEXT," _
        & "AuditTime DATETIME)"
    
        .Execute "CREATE TABLE LOANBAL (LoanNo TEXT,LoanCode TEXT,MemberNo TEXT(50)," _
        & "Balance CURRENCY,FirstDate DATETIME,RepayRate CURRENCY,LastDate DATETIME,RepayMethod" _
        & " TEXT,Cleared TEXT,AutoCalc TEXT,IntrAmount CURRENCY,RepayPeriod NUMBER,Remarks TEXT," _
        & "AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE LOANGUAR (MemberNo TEXT(50),LoanNo TEXT,Amount CURRENCY," _
        & "Balance CURRENCY,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE LOANS (LoanNo TEXT(20),MemberNo TEXT(50),LoanCode TEXT," _
        & "ApplicDate DATETIME,LoanAmt CURRENCY,RepayPeriod NUMBER,WitMemberNo TEXT,WitSigned " _
        & "TEXT,SupMemberNo TEXT,SupSigned TEXT,PreparedBy TEXT,AddSecurity TEXT,Insurance CURRENCY," _
        & "InsPercent CURRENCY,InsCalcType NUMBER,Posted TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE LOANSCHD (MemberNo TEXT(50),Period NUMBER,Principal " _
        & "CURRENCY,Interest CURRENCY,Balance CURRENCY,Comments TEXT,FmtPer TEXT)"
    
        .Execute "CREATE TABLE LOANTYPE (LoanCode TEXT,LoanType TEXT,RepayPeriod NUMBER," _
        & "Interest CURRENCY,MaxAmount CURRENCY,Guarantor TEXT,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE LOANVAR (LoanNo TEXT(20),MemberNo TEXT(50),OldRate " _
        & "CURRENCY,NewRate CURRENCY,VarDate DATETIME,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE MEMBERS (MemberNo TEXT(50),StaffNo TEXT(50),IDNo TEXT(20)," _
        & "AccNo TEXT(50),SurName TEXT(50),OtherNames TEXT(50),Sex TEXT(10),DOB DATETIME," _
        & "Employer TEXT,Dept TEXT,Rank TEXT,Terms TEXT,PresentAddr TEXT,OfficeTelNo TEXT," _
        & "HomeAddr TEXT,HomeTelNo TEXT(50),RegFee CURRENCY,InitShares CURRENCY,AsAtDate DATETIME,MonthlyContr " _
        & "CURRENCY,ApplicDate DATETIME,EffectDate DATETIME,Signed TEXT,Accepted TEXT,Archived " _
        & "TEXT,Withdrawn TEXT,IsGuarantor TEXT,Province TEXT,District TEXT,Station TEXT," _
        & "CompanyCode TEXT,PIN TEXT,Photo LONGBINARY,ShareCap CURRENCY,BankCode TEXT,AuditId " _
        & "TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE PTRTPR (MemberNo TEXT,MemberNames TEXT,OpeningBal " _
        & "CURRENCY,NewLoans CURRENCY,LoanRepayment CURRENCY,LoanInterest CURRENCY," _
        & "LoanOutstanding CURRENCY,MonthContr CURRENCY,SharesContr CURRENCY,SharesAsAtEOM" _
        & " CURRENCY,CurrentShares CURRENCY)"
    
        .Execute "CREATE TABLE REASONS (ReasonID NUMBER,Description TEXT)"
        
        .Execute "CREATE TABLE REPAY (LoanNo TEXT(20),DateReceived DATETIME," _
        & "PaymentNo NUMBER,Amount CURRENCY,Principal CURRENCY,Interest CURRENCY," _
        & "IntrCharged CURRENCY,IntrOwed CURRENCY,LoanBalance CURRENCY,ReceiptNo TEXT," _
        & "Locked TEXT,Posted TEXT,Accrued TEXT,Remarks TEXT,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE SAVINGS (MemberNo TEXT(50),TransDate DATETIME," _
        & "TransNo NUMBER,TransName TEXT,TransBy TEXT,ChequeNo TEXT,Amount CURRENCY," _
        & "ReceiptNo TEXT,IncTrans TEXT,Locked TEXT(3),AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE SHARES (MemberNo TEXT(50),TotalShares CURRENCY," _
        & "TransDate DATETIME,LastDivDate DATETIME,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE SHRVAR (MemberNo TEXT(50),OldContr CURRENCY," _
        & "NewContr CURRENCY,VarDate DATETIME,AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE SYSPARAM (ShareInterest CURRENCY,LoanInterest " _
        & "CURRENCY,MinGuarantors NUMBER,MaxGuarantors NUMBER,LoanToShareRatio CURRENCY," _
        & "MinLoanPeriod NUMBER,MinTotShares CURRENCY,MaxContribPeriod NUMBER," _
        & "BankInterest CURRENCY,WithdrawalNotice NUMBER,MinDivPeriod NUMBER,DeductAmt" _
        & " CURRENCY,SelfGuar TEXT(3),GuarShareRatio CURRENCY,CompanyName TEXT," _
        & "AuditID TEXT,AuditTime DATETIME)"
    
        .Execute "CREATE TABLE TEMPMEMBERSTATEMENT (MemberNo TEXT(50),RefNo " _
        & "TEXT,Description TEXT,TransCode NUMBER,Principal CURRENCY,Interest CURRENCY," _
        & "MonthlyContr CURRENCY,Total CURRENCY)"
    
        .Execute "CREATE TABLE TMPDIVIDENDPAYLIST (MemberNo TEXT(50)," _
        & "Current_Tot_Shares CURRENCY,Shares_As_At CURRENCY,Gross_Dividend " _
        & "CURRENCY,WithTax CURRENCY,Net_Dividend CURRENCY,BankName TEXT,AcctNo" _
        & " TEXT,CompanyName TEXT,ShareInterest NUMBER,WithHoldingTax NUMBER)"
    
        .Execute "CREATE TABLE TMPPERIODICTRANS (MemberNo TEXT(50)," _
        & "Names TEXT,OpeningBal CURRENCY,NewLoans CURRENCY,TotalLoanRepayment" _
        & " CURRENCY,TotalLoanInterest CURRENCY,LoanOutstanding CURRENCY," _
        & "TotMonthContr CURRENCY,SharesContr CURRENCY,SharesAsAt CURRENCY," _
        & "Period TEXT)"
        
        .Execute "CREATE TABLE TMPSTATEMENT (MemberNo TEXT(50),Period " _
        & "TEXT,OpeningBalance CURRENCY,NewLoan CURRENCY,LoanInterestPaid CURRENCY" _
        & ",LoanInterestCharged CURRENCY,LoanInterestOwing CURRENCY,LoanRepayment " _
        & "CURRENCY,OutstandingLoanBalance CURRENCY,OpeningShares CURRENCY," _
        & "SharesContributed CURRENCY,ClosingShares CURRENCY," _
        & "TotalMonthlyContribution CURRENCY)"
        
        .Execute "CREATE TABLE USERGRPS (GroupID TEXT,Description TEXT," _
        & "RWMembers TEXT,MemReg TEXT,NOK TEXT,MemStatement TEXT,RWShares TEXT," _
        & "Contributions TEXT,ShareVar TEXT,RWLoans TEXT,LoanApplic TEXT," _
        & "LoanEndorsement TEXT,ChequeEntry TEXT,LoanBal TEXT,LoanTrans TEXT," _
        & "EffectRepayment TEXT,LoanGuarantors TEXT,RWUtilities TEXT,Calc TEXT," _
        & "Dividends TEXT,PeriodTran TEXT,UtilStatements TEXT,MonthlyDeductions " _
        & "TEXT,ExportToGL TEXT,UtilGuarantor TEXT,Dormant TEXT,Archived TEXT," _
        & "Withdrawn TEXT,BackUp TEXT,RWBanking TEXT,RWOtherSchemes TEXT," _
        & "Savings TEXT,BenevolentFund TEXT,RWSetup TEXT,Parametization TEXT," _
        & "LoanTypes TEXT,CompanySetup TEXT,DataBaseSetup TEXT,Activate TEXT," _
        & "RejReasons TEXT,UserGrps TEXT,SysUsers TEXT,ClearLoanRecs TEXT," _
        & "ClearMemRecs TEXT,ChangeMemNo TEXT,AuditID TEXT,AuditTime DATETIME)"
        
        .Execute "CREATE TABLE USERS (UserID TEXT,GroupID TEXT," _
        & "UserPassword TEXT,IsSuperUser TEXT)"
    
        .Execute "CREATE TABLE WITHDRAWN (MemberNo TEXT," _
        & "DateWithdrawn DATETIME,TotalShares CURRENCY,AuditID TEXT," _
        & "AuditTime DATETIME)"
        
    End With
    Exit Sub
ErrorTrap:
    MsgBox err.description, , "CREATING DATABASE"
End Sub

