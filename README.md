- 👋 Hi, I’m @JAUBOTI
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
JAUBOTI/JAUBOTI is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
Option Compare Database

Private Sub Detail_Click()
DoCmd.ShowToolbar "Ribbon", acToolbarNo
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
DoCmd.ShowToolbar "Ribbon", acToolbarYes
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim isaved As VbMsgBoxResult
isaved = MsgBox("click yes to save or No to discard changes", vbInformation + vbYesNo, "Loan management system")
If saved = vbNo Then
DoCmd.RunCommand AccUndo
Cancel = l
End If
List153.Requery
End Sub
'------------------------------------------------------------
' CmdSave_Click
'
'------------------------------------------------------------
Private Sub CmdSave_Click()
On Error GoTo CmdSave_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    List153.Requery
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If
    MsgBox "Loan record saved", vbInformation, "Loan Management system"


CmdSave_Click_Exit:
    Exit Sub

CmdSave_Click_Err:
    MsgBox Error$
    Resume CmdSave_Click_Exit

End Sub


'------------------------------------------------------------
' CmdNext_Click
'
'------------------------------------------------------------
Private Sub CmdNext_Click()
On Error GoTo CmdNext_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="CmdSave" Event="OnClick" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="OnError"/><Action Name="SaveRecord"/
    ' _AXL:><ConditionalBlock><If><Condition>[MacroError]&lt;&gt;0</Condition><Statements><Action Name="MessageBox"><Argument Name="Message">=[MacroError].[Description]</Argument></Action></Statements></If></ConditionalBlock><Action Name="MessageBox"><Argumen
    ' _AXL:t Name="Message">Loan record saved</Argument><Argument Name="Beep">No</Argument><Argument Name="Type">Information</Argument><Argument Name="Title">Loan Management system</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.FindNext


CmdNext_Click_Exit:
    Exit Sub

CmdNext_Click_Err:
    MsgBox Error$
    Resume CmdNext_Click_Exit

End Sub


'------------------------------------------------------------
' CmdPrevious_Click
'
'------------------------------------------------------------
Private Sub Cmdprevious_Click()
On Error GoTo Cmdprevious_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="CmdNext" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="FindNextRecord"/></Statements></UserInterfaceMacro>
    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


Cmdprevious_Click_Exit:
    Exit Sub

Cmdprevious_Click_Err:
    MsgBox Error$
    Resume Cmdprevious_Click_Exit

End Sub


'------------------------------------------------------------
' CmdPrint_Click
'
'------------------------------------------------------------
Private Sub Cmdprint_Click()
On Error GoTo Cmdprint_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="CmdPrevious" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="OnError"/><Action Name="GoToRecord"><Argument Na
    ' _AXL:me="Record">Previous</Argument></Action><ConditionalBlock><If><Condition>[MacroError]&lt;&gt;0</Condition><Statements><Action Name="MessageBox"><Argument Name="Message">=[MacroError].[Description]</Argument></Action></Statements></If></ConditionalB
    ' _AXL:lock></Statements></UserInterfaceMacro>
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdPrintSelection


Cmdprint_Click_Exit:
    Exit Sub

Cmdprint_Click_Err:
    MsgBox Error$
    Resume Cmdprint_Click_Exit

End Sub


'------------------------------------------------------------
' CmdClose_Click
'
'------------------------------------------------------------
Private Sub Cmdclose_Click()
On Error GoTo Cmdclose_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="CmdPrint" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="RunMenuCommand"><Argument Name="Command">SelectReco
    ' _AXL:rd</Argument></Action><Action Name="RunMenuCommand"><Argument Name="Command">PrintSelection</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.Close , ""


Cmdclose_Click_Exit:
    Exit Sub

Cmdclose_Click_Err:
    MsgBox Error$
    Resume Cmdclose_Click_Exit

Dim tblpersonalledger As DAO.Recordset

Set tblCustomers = CurrentDb.OpenRecordset("SELECT * FROM [tblpersonalledger]")
tblpersonalledger.AddNew
tblpersonalledger![loan_reff_ID] = loanreffID.Value
tblpersonalledger![Amount_of_loan] = Amountofloan.Value
tblpersonalledger![number_of_fortnights] = numberoffortnights.Value
tblpersonalledger![inerest_rate] = interestrate.Value
tblpersonalledger![fortnightinstallments] = Amount_of_loan * interest / number_of_fortnights.Value
tblpersonalledger![Total_Payment] = PVal = ("## number of fortnights") + ("##interestrate") = ("####.##totalpayment")
FPR = ("number_of_fortnights " & _
      "interest_ rate")
tblpersonalledger![Full_Name] = FullName.Value
tblpersonalledger![Organisation_Department] = Organisation_Department.Value
tblpersonalledger![File_Number] = File_Number.Value
tblpersonalledger![Address] = Address.Value
tblpersonalledger![Phone_Number] = phonenumber.Value
tblpesonalledger![Bank_branch] = bankbranch.Value
tblpersonalledger![Bank_Account_number] = bankaccountnumber.Value
tblpersonalledger![Batch_Number] = batchnumber.Value
tblpersonalledger![Batch_Date] = batchdate.Value
tblpersonalledger![Standing_order_number] = (Standing_order_number.Value)
End Sub
Private Sub cmdSHOWdefaultcharg_AFTERUPDATE_AFTER_28_DAYS()
Dim intTotalPAYMENT As Double '\\if i use interger i get "overflow" error(Cancel As Integer)
tblpersonalledger![Default_Charge] = ("28_days")
Dim Fmt, FVal, PVal, FPR, TotPmts, PayType, Payment
' When payments are made.
Const ENDPERIOD = ("14 days"), commencementdate = 1
If FPR = ("14DAYS") And FPR = ("000.00") Then comencementdate = 1
Fmt = "###,###,##0.00" 'money format=(K).
FVal = (0) ' Usually 0 for a loan.
End Sub
Private Sub CMDSHOWDEFAULTCHARGE_POP_UP_AFTER_28_DAYS()
Const inputbox = ("####.##"), commencementdate = 1
Fmt = "###,###,##0.00" 'money format=()
FPR = ("number of fortnights " & _
      "interest rate")
If FPR > ("14DAYS") Then FPR = (10 / 100) '+ fortnight installments = totalpayment.
TotPmts = ("##number of fortnights") = 14
PayType = MsgBox("Do you make payments " & _
          "at the end of 14 days", vbYesNo)
If PayType = vbNo Then
    PayType = comencementdate = 0
    Else: PayType = ENDPERIOD
End If

PayType = (Rate(FPR / 14, TotPmts, -PVal, FVal, PayType, GUESS) * 14) * 100
MsgBox "totalpayment " & _
    Format(Payment, Fmt) & " per fortnight."

tblpersonalledger.Update
End Sub


