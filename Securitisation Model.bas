Attribute VB_Name = "Module1"
Option Explicit
Public nobs As Long, nc As Integer, nodef As Integer, defaultmonth As Integer, i As Long, msgans As Integer, x As Variant
Public uniques As Variant, lgdfix As Variant, pdfix As Variant, yr_mn_confirm As Variant, raw_data As Variant, starttime As Long
Public cumulative_arr As Variant, pd_Xead As Variant, pd_Xead_sum_by_month As Variant

Sub EAD()

'The following line counts the number of observations in our dataset and then Number of defaults we want to simulate
nobs = Sheet2.Range("T1")
nodef = Sheet9.Range("U4") ' Number of defaults

Application.ScreenUpdating = False

starttime = Timer

raw_data = Sheet2.Range("A2:J34144")

' Making sure there are no errors for nodef
If nodef <= 0 Or nodef >= nobs Then
    MsgBox "Number of defaults is empty. Please select a value from the drop down list"
    Exit Sub
End If


' Making sure with the user that he knows whether he wants default month values constant or random using msgbox inside an if statement
If Sheet9.Range("U5") = "" Then
    msgans = MsgBox("Do you want the month of defaults to be constant for all loans?", vbYesNo + vbQuestion)
        If msgans = vbNo Then
            'User wants the random values for year of default so we directly go to the EAD computation.
             Sheet9.Range("U5").Clear
             Call Cal_EAD
             
        ElseIf msgans = vbYes Then
            ' User wants the month of default to be constant
            If Sheet9.Range("U5").Value > 360 Or IsEmpty(Sheet9.Range("U5")) Then
                Sheet9.Range("U5").Clear
                 x = InputBox("Please enter a value for the month of default between month 3 and 360")
                 ' Now we are checking if the user entered a valid range
                If x = "" Then
                    MsgBox "Invalid Input"
                    Exit Sub
                ElseIf IsNumeric(x) And x >= 3 And x <= 360 Then ' If everything is fine, paste the user value to U5.
                    Sheet9.Range("U5") = x
                    Call Cal_EAD
                ElseIf IsNumeric(x) And x <= 2 Then
                    MsgBox "Please enter a valid month. Enter a value between 3 and 360"
                    Sheet9.Range("W2:AB" & nobs).Clear
                    Exit Sub
                End If
            ElseIf Sheet9.Range("U5").Value < 3 Then
                MsgBox "Please enter a valid month. Enter a value between 3 and 360"
                Exit Sub
            ElseIf Sheet9.Range("U5").Value >= 3 And Sheet9.Range("U5").Value <= 360 Then
                defaultmonth = Sheet9.Range("U5")
                Call Cal_EAD
            End If
        End If
Else
    Call Cal_EAD
End If

End Sub


Sub Cal_EAD()

' THis Sub is the first step, it sees loan ID's based on year of default then fetches the term, credit rating and the EAD values.

Dim zeroEADflag As Boolean, loan_origterm As Integer, loan_ID As Variant, check_termerr As Variant, check_EADval As Variant
Dim iy1 As Double, iy2 As Double, iy As Double, outstanding_balance As Double, pmt1 As Double, pmt2 As Double, pmt As Double

zeroEADflag = False

Application.ScreenUpdating = True


Sheet9.Range("W2:AD" & nobs).Clear
Sheet9.Range("AF2:AM" & nobs).Clear
Sheet9.Range("K17:O376").Clear

Application.ScreenUpdating = False

'                        ----- ADDING CODE TO GET UNIQUE RANDOM NUMBERS in Column W-------

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim randID As Long
Dim rowIndex As Long
Dim arr As Variant
ReDim arr(1 To nodef, 1 To 8) As Variant

ReDim pd_Xead(1 To nodef, 1 To 1)


rowIndex = 1

Do While dict.Count < nodef
    randID = WorksheetFunction.RandBetween(1, nobs)
    
    If Not dict.exists(randID) Then
        dict.Add randID, True
        arr(rowIndex, 1) = randID
        rowIndex = rowIndex + 1
    End If
Loop
                              
                                ' The main calculations starts from now
                
 For i = 1 To nodef
    
    'The random values are already. A new line of code is no longer needed as we have dealt with it in the above Do.while loop to get unique random  numbers.
    
     ' Getting the Loan ID's from the "Data for Securitization Sheet" based on the random loan serial numbers from above and putting it in Column "X".
      
      loan_ID = Application.VLookup(arr(i, 1), raw_data, 3, False)
        
        If IsError(loan_ID) Then
          
          MsgBox "Loan ID lookup failed for row " & i + 1, vbExclamation
          Exit Sub
        
        Else
          
          arr(i, 2) = Application.VLookup(arr(i, 1), raw_data, 3, False)
        
        End If
      
     ' Getting the term of the laons (maturity)  from the raw_data array
       check_termerr = Application.VLookup(arr(i, 1), raw_data, 6, False)
       
       If IsError(check_termerr) Then
           MsgBox "Loan ID lookup failed for row " & i + 1, vbExclamation
           Exit Sub
       Else
       
       ' Getting the term of the loans
       arr(i, 3) = check_termerr
       
        ' The range for vlookup in the above code is hardcoded, after everything see if you could make it dynamic for any loan numbers.
       End If
      
      If IsEmpty(Sheet9.Range("U5").Value) Then
            'Randomly assigning month of defaults. Loan_origterm is in terms of months
            
            loan_origterm = Int(arr(i, 3))
             
             If loan_origterm < 1 Or loan_origterm <= 0 Then ' This ensures term < 12 months doesn't crash RandBetween() for new data in future.
                 arr(i, 4) = 1
             Else
                ' assigning random month of default.
                 arr(i, 4) = WorksheetFunction.RandBetween(3, loan_origterm)
            End If
            
      ElseIf Sheet9.Range("U5").Value >= 1 And Sheet9.Range("U5").Value <= 360 Then
             'Setting the month for the default = U5
              arr(i, 4) = Sheet9.Range("U5")
      End If

      ' Assigning each Loan ID to the Amortization schedule one by one
       Sheet9.Range("C2") = arr(i, 2)
      
      'Fetching the outstanding balance from the amortization schedule based on the (month of default) using If statements
        
        'The following IF statement sets EAD=0 if mod > term of the loan (in months)
        If arr(i, 4) > arr(i, 3) Then
            arr(i, 6) = 0
        
        Else
            check_EADval = Application.VLookup(arr(i, 4), Sheet9.Range("B10:F370"), 5, False)
            
            If IsError(check_EADval) Then
                arr(i, 6) = 0
             Else
                'Adding Principal EAD ie., t-3 months o/s balance for waterfall structure-  used later in the code, we also have the cumulative column of it **
                
                ' Here we are getting the outstanding balance for the month of default
                outstanding_balance = Sheet9.Application.VLookup((arr(i, 4) - 3), Sheet9.Range("B10:F370"), 5, False)
                
                ' Computing EAD = principal EAD + interest component for last 3 months
                
                ' t month interest component:
                iy = Sheet9.Application.VLookup(arr(i, 4), Sheet9.Range("B10:F370"), 4, False)
                ' t-1 month Interest component :
                iy1 = Sheet9.Application.VLookup(arr(i, 4) - 1, Sheet9.Range("B10:F370"), 4, False)
                ' t-2 month Interest component:
                iy2 = Sheet9.Application.VLookup(arr(i, 4) - 2, Sheet9.Range("B10:F370"), 4, False)
                
                'EAD:
                
                 arr(i, 6) = outstanding_balance + iy1 + iy2 + iy
                              
                ' Adding the Cumulative PMT (Cash flow/ EMI) not received column for the last 3 months - Used for Waterfall Structure.
                 pmt = Sheet9.Application.VLookup(arr(i, 4), Sheet9.Range("B10:F370"), 2, False)
                 
                 arr(i, 7) = (pmt * 3)
                           
                'Adding borrower creditscore value as well for future
                
                'Adding Principal EAD ie., t-3 months o/s balance for waterfall structure-  used later in the code, we also have the cumulative column of it **
                arr(i, 8) = Sheet9.Application.VLookup((arr(i, 4) - 3), Sheet9.Range("B10:F370"), 5, False)
                
                'Borrower Credit rating Value
                arr(i, 5) = Application.VLookup(arr(i, 2), raw_data, 10)
                
                pd_Xead(i, 1) = Application.XLookup(arr(i, 5), Sheet2.Range("P30:P34"), Sheet2.Range("R30:R34").Value, , 0) * arr(i, 6)
                
            End If
        End If
        
 Next
 
 Sheet9.Range("W2:AD" & nodef + 1) = arr
 
 
 
 Sheet9.Range("AB2:AD" & nodef).NumberFormat = "#,##0.00" ' Adding number format to EAD column
 
 'The following For loop and if statement is for visuals and editing such that any value that is 0 or close to 0 to be highlighted.
 
 For i = 1 To nodef
    If Abs(Sheet9.Range("AB" & i + 1).Value) < 0.01 Then
       Sheet9.Range("AB" & i + 1).Interior.ColorIndex = 50
       zeroEADflag = True
    End If
Next
 If zeroEADflag = True Then
    MsgBox ("The highlighted cells in the EAD column have a value of zero because either the loan term = default year or the term of the loan is less than the default year.")
 End If

Call Cumu_EAD

End Sub

Sub Cumu_EAD()
' This Sub Calculates the Cumulative EAD values for each year of default and puts them in Range("AF:AG")

Application.ScreenUpdating = False

Sheet9.Range("AF2:AM" & nobs).Clear

Dim dict As Object, i As Long, val As Variant, msgans As Variant, lgdans As Variant, lgdpdans As Variant, pdans As Variant, pdfix As Variant

'Getting the unique values from our Z range and sorting it in ascending order using bubble sort
Set dict = CreateObject("Scripting.Dictionary")

' Collecting unique values from Column Z
Dim lastrow As Long
lastrow = Sheet9.Cells(Sheet9.Rows.Count, "Z").End(xlUp).Row

For i = 2 To lastrow

    val = Sheet9.Range("Z" & i)
    If Not dict.exists(val) And Not IsEmpty(val) Then
            dict.Add val, val
    End If
Next

'Transferring keys and values to an array and sorting them by calling the Bubblesort sub
uniques = dict.items
Call Bubblesort(uniques)

'This writes the sorted, unique values from your uniques array into Column AF, starting from row 2 downward.

ReDim cumulative_arr(UBound(uniques), 1 To 8)

ReDim pd_Xead_sum_by_month(UBound(uniques), 1 To 1)

For i = 0 To UBound(uniques) ' We did not use nodef because nodef can iterate to 3000 or more but we want the sum of EAD's which at max can go to 360
    cumulative_arr(i, 1) = uniques(i)
    
    'In the following line of code, we are finding the cumulative EAD values
    cumulative_arr(i, 2) = Sheet9.Application.WorksheetFunction.SumIfs(Range("AB2:AB" & nodef + 1), Range("Z2:Z" & nodef + 1), cumulative_arr(i, 1))
    If cumulative_arr(i, 2) < 0.01 Then
        cumulative_arr(i, 2) = 0
    End If
        
    'In the following line of code, we are finding the cumulative 1-month PMT values for each month
    cumulative_arr(i, 7) = Application.WorksheetFunction.SumIfs(Range("AC2:AC" & nodef + 1), Range("Z2:Z" & nodef + 1), cumulative_arr(i, 1)) / 3
    
    ' In the following line of code we are finding the Cumulative Outstanding balance values -> used for the waterfall structure
    cumulative_arr(i, 8) = Sheet9.Application.WorksheetFunction.SumIfs(Range("AD2:AD" & nodef + 1), Range("Z2:Z" & nodef + 1), cumulative_arr(i, 1))
    
    If cumulative_arr(i, 8) < 0.01 Then cumulative_arr(i, 8) = 0
    
    pd_Xead_sum_by_month(i, 1) = Application.SumIfs(Range("AN2:AN" & nodef + 1), Range("Z2:Z" & nodef + 1), cumulative_arr(i, 1))
     
Next
Sheet9.Range("AF2:AM" & UBound(uniques) + 2) = cumulative_arr

Sheet9.Range("AG1:AG" & UBound(uniques) + 2).NumberFormat = "#,##0.00"

                                            '     Now we are moving to EL calcuations

msgans = MsgBox("Do you want to see the actual defaults (standing at year =30) or Expected defaults (Standing at Year=1)? Click Yes for actual defaults and No for Expected defaults", 4)
      If msgans = vbYes Then
          ' The user wants actual defaults, so PD =1, he is standing at Year =30
          lgdans = MsgBox("Do you want LGD to be predefined at 45%?", 4)
                  If lgdans = vbYes Then
                          Call Calculate_Actual_prefixed_LGD
                          
                  ElseIf lgdans = vbNo Then
                          ' User Wants his own value for LGD
                          lgdfix = InputBox("Please enter a value for LGD :")
                          If IsNumeric(lgdfix) Then
                              If lgdfix >= 1 And lgdfix < 100 Then
                                  lgdfix = lgdfix / 100
                                  Call Calculate_Actual_UserLGD(lgdfix)
                              ElseIf lgdfix >= 0 And lgdfix <= 1 Then
                                  Call Calculate_Actual_UserLGD(lgdfix)
                              ElseIf lgdfix > 100 Or lgdfix <= 0 Then
                                  MsgBox "Please Enter a number between 1 and 100"
                                  Exit Sub
                              End If
                          Else
                              MsgBox "Please enter a numeric value"
                          End If
                  End If
                  
      ElseIf msgans = vbNo Then
          ' The user wants to see the Expected defaults, he is standing at Year =1
          lgdpdans = MsgBox("Do you want to see the pre-defined PD based on rating grade and LGD = 45% for all loans?" & _
                            "If you wish to update the PD for each rating grade, please do so by changing the values in the ""Raw data for Waterfall"" sheet directly and press cancel", vbYesNoCancel)
          
          If lgdpdans = vbCancel Then
                
                Exit Sub
          
          ElseIf lgdpdans = vbYes Then
              
              Call Calculate_Expected_Random
    
          ElseIf lgdpdans = vbNo Then
              ' The user wants his own values of PD and LGD
                lgdfix = InputBox("Please enter a value for LGD :")
                If IsNumeric(lgdfix) Then
                  If (lgdfix >= 1 And lgdfix < 100) Then
                      lgdfix = lgdfix / 100
                      Call Calculate_Expected_Fixed(lgdfix)
                  ElseIf lgdfix < 1 And lgdfix > 0 Then
                      Call Calculate_Expected_Fixed(lgdfix)
                  ElseIf lgdfix < 0 Then
                      MsgBox "Please enter a valid numeric value"
                      Exit Sub
                  End If
                End If
          End If
      End If
End Sub
Sub Calculate_Actual_prefixed_LGD()

For i = 0 To UBound(uniques)

    ' Assigning LGD = 45% for all observations in Column AH
     cumulative_arr(i, 3) = 0.45
    
    ' Assigning PD =1 for all observations in Column AI
    cumulative_arr(i, 4) = 1
    
    'Computing Expected loss
    cumulative_arr(i, 5) = cumulative_arr(i, 3) * cumulative_arr(i, 2)
    
    ' Computing Recoveries:
    cumulative_arr(i, 6) = cumulative_arr(i, 2) * (1 - cumulative_arr(i, 3))

    'To check for errors
    If cumulative_arr(i, 2) < 0.01 Then cumulative_arr(i, 2) = 0
Next

Sheet9.Range("AF2:AM" & UBound(uniques) + 2) = cumulative_arr

'Formatting :

Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
Sheet9.Range("AH2:AI" & UBound(uniques) + 2).NumberFormat = "#0.00%"


Call Waterfall

End Sub

Sub Calculate_Actual_UserLGD(lgdfix As Variant)

For i = 0 To UBound(uniques)

    ' Assigning PD =1 for all in Column AI
    cumulative_arr(i, 4) = 1
    
    ' Assigning LGD = User specific
    cumulative_arr(i, 3) = lgdfix
    
    ' Computing Expected Loss
    cumulative_arr(i, 5) = cumulative_arr(i, 2) * cumulative_arr(i, 3)
    
    ' Computing Recoveries :
    cumulative_arr(i, 6) = cumulative_arr(i, 2) * (1 - cumulative_arr(i, 3))
        
    'To check for errors
    If cumulative_arr(i, 2) < 0.01 Then cumulative_arr(i, 2) = 0
    
Next

Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
Sheet9.Range("U5").Clear

Sheet9.Range("AF2:AM" & UBound(uniques) + 2) = cumulative_arr

' Formatting :
Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
Sheet9.Range("AH2:AI" & UBound(uniques) + 2).NumberFormat = "#0.00%"


Call Waterfall ' Calling the Waterall Subroutine

End Sub


Sub Calculate_Expected_Random()

For i = 0 To UBound(uniques)
    
    ' The following IF...elseif statemens assigns PD values
    
    cumulative_arr(i, 4) = pd_Xead_sum_by_month(i, 1) / cumulative_arr(i, 2)
    
    ' Assigning LGD = pre defined value
    cumulative_arr(i, 3) = 0.45
    'Sheet9.Range("AH" & i + 2) = 0.45
    
    'Computing EL for each observation and putting the values in cell.
    cumulative_arr(i, 5) = cumulative_arr(i, 3) * cumulative_arr(i, 4) * cumulative_arr(i, 2)
    'Sheet9.Range("AJ" & i + 2) = Sheet9.Range("AI" & i + 2).Value * Sheet9.Range("AH" & i + 2) * Sheet9.Range("AG" & i + 2)
    
    ' Computing Recoveries
    cumulative_arr(i, 6) = cumulative_arr(i, 2) * (1 - cumulative_arr(i, 3))
    
    'To check for errors
    If cumulative_arr(i, 2) < 0.01 Then cumulative_arr(i, 2) = 0
        
Next
  ' Number Formatting
    Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
    Sheet9.Range("AH2:AI" & UBound(uniques) + 2).NumberFormat = "#0.0%"

'REMOVE THE VALUES IN AN and AO column:


Sheet9.Range("U5").Clear
Sheet9.Range("AF2:AM" & UBound(uniques) + 2) = cumulative_arr

'Formatting :

Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
Sheet9.Range("AH2:AI" & UBound(uniques) + 2).NumberFormat = "#0.00%"


Call Waterfall

End Sub

Sub Calculate_Expected_Fixed(lgdfix As Variant)

Application.ScreenUpdating = False

' LGD and PD values comes from user

For i = 0 To UBound(uniques)
    
    ' Assuming the user changed the PD for each rating grade in the specified sheet before the start of the simulation:
    cumulative_arr(i, 4) = pd_Xead_sum_by_month(i, 1) / cumulative_arr(i, 2)
    
    'Assigning LGD = User specified for all in Column AH
    
    cumulative_arr(i, 3) = lgdfix
    
    'Computing EL for each observation and putting the values in cell.
    cumulative_arr(i, 5) = cumulative_arr(i, 3) * cumulative_arr(i, 4) * cumulative_arr(i, 2)
    
    cumulative_arr(i, 6) = cumulative_arr(i, 2) * (1 - lgdfix)
    
    'To check for errors
    If cumulative_arr(i, 2) < 0.01 Then cumulative_arr(i, 2) = 0
    
Next
Sheet9.Range("U5").Clear

Sheet9.Range("AF2:AM" & UBound(uniques) + 2) = cumulative_arr

'Formatting :

Sheet9.Range("AJ2:AM" & UBound(uniques) + 2).NumberFormat = "#,##0.00"
Sheet9.Range("AH2:AI" & UBound(uniques) + 2).NumberFormat = "#0.00%"


Call Waterfall

Application.ScreenUpdating = True

End Sub

Sub Waterfall()

Application.ScreenUpdating = False

Dim Gross_cash_flow As Variant, net_cash_flow As Variant, month_of_defaults As Variant, recovery_at_i1 As Variant, missed_pmt_at_t As Variant, missed_pmt_at_tp1 As Variant, missed_pmt_at_tp2 As Variant

Dim cumulative_principal_ead As Variant, month_i As Double, equity_cap_value As Double, equity_cap_remaining As Double, senior_value As Double, mezz_value As Double, equity_value As Double

Dim senior_ipmt As Double, senior_ppmt As Double, mezz_ipmt As Double, mezz_ppmt As Double, equity_ipmt As Double, equity_ppmt As Double

Dim net_cash_flow_after_senior As Double, net_cash_flow_after_mezz As Double, shortfall_mezz As Double, shortfall_equity As Double, net_cash_flow_after_equity As Double

Dim reserve_balance As Double, final_shortfall As Double, shortfall_senior As Double, net_cash_flow_after_senior_ipmt As Double, total_amount_to_write_off As Double

Dim lastrow As Long, unique_months As Variant, net_cash_flow_after_mezz_ipmt As Double, net_cash_flow_after_equity_ipmt As Double

Dim senior_face_value As Double, mezz_face_value As Double, equity_face_value As Double, pool_face_value As Double

Dim amount_left_after_equity_write_off As Double

' Dimming new variables: Need to initialize them****

Dim principalEAD_writedown As Double, writedown_from_reserves As Double, writedown_from_equity As Double, writedown_from_mezz As Double, writedown_from_senior As Double

Dim senior_pmt As Double, mezz_pmt As Double, equity_pmt As Double, total_obligation As Double

Dim value_to_deduct_from_equity As Double, value_to_deduct_from_mezz As Double, value_to_deduct_from_senior As Double, amount_taken_from_reserve As Double

lastrow = Sheet9.Cells(Rows.Count, "AF:AF").End(xlUp).Row

unique_months = Sheet9.Range("AF2:AF" & lastrow)

equity_cap_value = Sheet9.Range("M4").Value   ' A lifetime cap on equity, they cannot be paid more than this amount

equity_cap_remaining = equity_cap_value

Dim waterfall_arr As Variant

ReDim waterfall_arr(1 To 360, 1 To 5)

For month_i = 1 To 360 ' We are running the laon loop for 360 months- fixed
  
  ' Initializing the missed pmt values for t, t-1 and t-2 :
    missed_pmt_at_t = 0:   missed_pmt_at_tp1 = 0: missed_pmt_at_tp2 = 0: recovery_at_i1 = 0
    
    ' Initialzing Tranche values:
    senior_value = 0: mezz_value = 0: equity_value = 0
    
    'Initializing shortfall variables:
    final_shortfall = 0: shortfall_senior = 0: shortfall_mezz = 0
    
    ' Initializing net cash flow values:
    net_cash_flow_after_senior_ipmt = 0:  net_cash_flow_after_senior = 0
    net_cash_flow_after_mezz_ipmt = 0:    net_cash_flow_after_mezz = 0
    net_cash_flow_after_equity_ipmt = 0:  net_cash_flow_after_equity = 0
    Gross_cash_flow = 0:                  net_cash_flow = 0
    
    ' Initializing pmts, ipmts, recoveries, face values and ead values:
    pool_face_value = 0
    senior_ipmt = 0:     senior_ppmt = 0: senior_face_value = 0
    mezz_ipmt = 0:       mezz_ppmt = 0: mezz_face_value = 0
    equity_ppmt = 0:     equity_ipmt = 0: equity_face_value = 0
    reserve_balance = 0
    month_of_defaults = 0: cumulative_principal_ead = 0
    
    'Before anything I need the Gross Cash flow for each month from the pool amortization table
    'Getting the gross (expected) cash flow for month_i:
    Gross_cash_flow = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("B11:F371"), 2, False)
    
    ' Cheking uisng if statement for defaults in month i, i+1 and i+2
     
    ' Getting the missed PMT for defaults happening at time t
        If IsInArray(month_i, uniques) Then
            ' Getting the cumulative 1-month PMT values to be dedcuted from gross cash flow
             missed_pmt_at_t = Application.WorksheetFunction.VLookup(month_i, Sheet9.Range("AF2:AM" & lastrow), 7, False)
            Else
            Debug.Print month_i & " NOT found in unique_months"
            missed_pmt_at_t = 0
        End If
        
     ' Getting the missed PMT for defaults happening at t + 1
        If IsInArray(month_i + 1, uniques) Then
             missed_pmt_at_tp1 = Application.WorksheetFunction.VLookup(month_i + 1, Sheet9.Range("AF2:AM" & lastrow), 7, False)
        Else
            Debug.Print month_i & " NOT found in unique_months"
            missed_pmt_at_tp1 = 0
        End If
     
     ' ' Getting the missed PMT for defaults happening at t +2
        
        If IsInArray(month_i + 2, uniques) Then
            missed_pmt_at_tp2 = Application.WorksheetFunction.VLookup(month_i + 2, Sheet9.Range("AF2:AM" & lastrow), 7, False)
        Else
            missed_pmt_at_tp2 = 0
        End If
        
    ' Checking for recoveries happening at month >3, if yes, add then to next month cash flows
    If month_i > 3 Then
        If IsInArray(month_i - 1, uniques) Then  ' If this condition is true, it means default has happened at previous month. So, we add recoveries
            recovery_at_i1 = Application.WorksheetFunction.VLookup(month_i - 1, Sheet9.Range("AF2:AM" & lastrow), 6, False)
        End If
    End If
  
    ' Now, we are getting the net cash flow
    net_cash_flow = Gross_cash_flow - missed_pmt_at_t - missed_pmt_at_tp1 - missed_pmt_at_tp2 + recovery_at_i1
    
    ' Accounting for defaults and shortfalls into the net_cash flow and writting off any defaults from tranches:
    
    ' Step 1: Accounting for defaults (EAD) and writeoffs follow: reserves--> equity principal --> mezzanine --> senior principal
    
    If IsInArray(month_i, uniques) Then
        principalEAD_writedown = Application.WorksheetFunction.VLookup(month_i, cumulative_arr, 8, False)
    End If
    
    ' After getting the principal EAD, we will first reduce the amount from the pool balance to reflect change in the gross cash flow and then
    ' we will absorb the losses starting from reserves.
    
    '#1 Reducing pool face amount:
    pool_face_value = Sheet10.Range("C4").Value
    pool_face_value = pool_face_value - principalEAD_writedown
    Sheet10.Range("C4") = pool_face_value

    ' #2 Absorbing losses:
    
    'Taking previous month reserve balance:
    If month_i = 1 Then
        reserve_balance = 0
    Else
        reserve_balance = waterfall_arr(month_i - 1, 5)
    End If
    
    ' Checking if reserve balance can absorb losses:
    
    If reserve_balance > 0 Then
        writedown_from_reserves = Application.WorksheetFunction.Min(reserve_balance, principalEAD_writedown)
        principalEAD_writedown = principalEAD_writedown - writedown_from_reserves
    End If
    
    If principalEAD_writedown > 0 Then
        equity_face_value = Sheet10.Range("AD4")
        writedown_from_equity = Application.WorksheetFunction.Min(equity_face_value, principalEAD_writedown)
        principalEAD_writedown = principalEAD_writedown - writedown_from_equity
        equity_face_value = equity_face_value - writedown_from_equity
        Sheet10.Range("AD4") = equity_face_value
    End If
    
    If principalEAD_writedown > 0 Then
        mezz_face_value = Sheet10.Range("U4")
        writedown_from_mezz = Application.WorksheetFunction.Min(principalEAD_writedown, mezz_face_value)
        principalEAD_writedown = principalEAD_writedown - writedown_from_mezz
        mezz_face_value = mezz_face_value - writedown_from_mezz
        Sheet10.Range("U4") = mezz_face_value
    End If
    
    If principalEAD_writedown > 0 Then
        senior_face_value = Sheet10.Range("L4")
        writedown_from_senior = Application.WorksheetFunction.Min(principalEAD_writedown, senior_face_value)
        principalEAD_writedown = principalEAD_writedown - writedown_from_senior
        senior_face_value = senior_face_value - writedown_from_senior
        Sheet10.Range("L4") = senior_face_value
    End If
    
    If principalEAD_writedown > 0 Then
        MsgBox ("Excess defaults than anticipated caused senior tranche to be wiped off")
        Exit For
    End If
    
    ' Step 2: After adjusting for defaulted accounts, we now move on to check if the net cash flow is sufficient to pay all obligations,
    ' if not, we have a shortfall and proceed to writeoffs starting from equity --> mezzanine --> senior
    
    senior_pmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("K11:O371"), 2, False)
    mezz_pmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("T11:X371"), 2, False)
    equity_pmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("AC11:AG371"), 2, False)
    
    total_obligation = senior_pmt + mezz_pmt + equity_pmt
    
    If net_cash_flow >= total_obligation Then
        net_cash_flow = net_cash_flow
    Else ' Checking if reserve balance + net cash flow is sufficient to pay total obligations:
        amount_taken_from_reserve = Application.WorksheetFunction.Min(reserve_balance, total_obligation - net_cash_flow)
        net_cash_flow = net_cash_flow + amount_taken_from_reserve
        If net_cash_flow > total_obligation Then
            net_cash_flow = net_cash_flow
            reserve_balance = reserve_balance - amount_taken_from_reserve
                
        Else
            final_shortfall = total_obligation - net_cash_flow
        End If
    End If
    
    If final_shortfall > 0 Then
        equity_face_value = Sheet10.Range("AD4")
        If equity_face_value > 0 Then
            value_to_deduct_from_equity = Application.WorksheetFunction.Min(equity_face_value, final_shortfall)
            equity_face_value = equity_face_value - value_to_deduct_from_equity
            Sheet10.Range("AD4") = equity_face_value
            final_shortfall = final_shortfall - value_to_deduct_from_equity
        End If
        ' If we still have any shortfal left, we write off mezzanine:
        If final_shortfall > 0 Then
            mezz_face_value = Sheet10.Range("U4")
            If mezz_face_value > 0 Then
                value_to_deduct_from_mezz = Application.WorksheetFunction.Min(mezz_face_value, final_shortfall)
                final_shortfall = final_shortfall - value_to_deduct_from_mezz
                mezz_face_value = mezz_face_value - value_to_deduct_from_mezz
                Sheet10.Range("U4") = mezz_face_value
            End If
        ' If we still have any shortfall left, after writing-off mezzanine, we proceed to write off senior's:
            If final_shortfall > 0 Then
                senior_face_value = Sheet10.Range("L4")
                If senior_face_value > 0 Then
                    value_to_deduct_from_senior = Application.WorksheetFunction.Min(senior_face_value, final_shortfall)
                    final_shortfall = final_shortfall - value_to_deduct_from_senior
                    senior_face_value = senior_face_value - value_to_deduct_from_senior
                    Sheet10.Range("L4") = senior_face_value
                End If
            End If
        End If
    End If

    ' Now, we are pasting the Net Cash Flow value in our Waterfall table
    
     waterfall_arr(month_i, 1) = net_cash_flow
    'Sheet9.Range("K" & month_i + 16) = net_cash_flow --> uncommented the follwowing to see the value in the worksheet.
    
    ' Now I need each tranche values based on net cash flows and ALSO MAKE SURE FIRST IPMT IS PAID TO EACH TRANCHE THEN PPMT AND IF LOSS OCCURED, I DEDUCT PRINCIPAL O/S FROM RESPECTIVE TRANCHE AND THE POOL O/S BALANCE.
        
    'Getting senior tranche interest and principal, respectively, payments seperately:
    
    senior_ipmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("K11:O371"), 4, False)
    senior_ppmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("K11:O371"), 3, False)
    
    'Getting mezzanine tranche interest and principal, respectively, payments seperately :
    
    mezz_ipmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("T11:X371"), 4, False)
    mezz_ppmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("T11:X371"), 3, False)
    
    ' Getting this month's reserve balance, because we have already updated this month's reserve balance for defaults:
    
    ' Checking if Senior face value is > 0 and then we check if Senior ipmt +ppmt < net cash flow
    If Sheet10.Range("L4") > 0 Then
        ' Since senior Face value > 0, we proceed with payments
        If net_cash_flow > senior_ipmt + senior_ppmt Then
            senior_value = senior_ipmt + senior_ppmt
            net_cash_flow_after_senior = net_cash_flow - senior_ipmt - senior_ppmt
        Else
            If net_cash_flow > senior_ipmt Then
                senior_value = senior_ipmt
                net_cash_flow_after_senior_ipmt = net_cash_flow - senior_ipmt
            ElseIf net_cash_flow < senior_ipmt Then
                shortfall_senior = senior_ipmt - net_cash_flow
                If reserve_balance >= shortfall_senior Then
                    senior_value = senior_ipmt
                    reserve_balance = reserve_balance - shortfall_senior
                    net_cash_flow_after_senior = 0 ' Since net cash flow < senior_ipmt, net cash flow after paying senior's ipmt becomes 0
                ElseIf reserve_balance < shortfall_senior Then
                    senior_value = reserve_balance
                    shortfall_senior = shortfall_senior - reserve_balance
                    reserve_balance = 0
                    net_cash_flow_after_senior = 0 ' Since net cash flow < senior_ipmt, net cash flow after senior's ipmt becomes 0
                End If
            End If
            
            ' Now, we are checking that after paying senior IPMT, do we have enough cash to also pay PPMT?
            
            If net_cash_flow_after_senior_ipmt > senior_ppmt Then
                senior_value = senior_value + senior_ppmt
                net_cash_flow_after_senior = net_cash_flow_after_senior_ipmt - senior_ppmt
            ElseIf net_cash_flow_after_senior_ipmt < senior_ppmt Then
                If reserve_balance >= senior_ppmt Then
                    senior_value = senior_value + senior_ppmt
                    reserve_balance = reserve_balance - senior_ppmt
                    net_cash_flow_after_senior = 0 ' Since net cash flow after senior ipmt < senior_ppmt, net cash flow after senior becomes 0 ie., there is nothing left to pay others.
                ElseIf reserve_balance < senior_ppmt Then
                    senior_value = senior_value + reserve_balance
                    reserve_balance = 0
                    net_cash_flow_after_senior = 0 ' Since net cash flow after senior ipmt < senior_ppmt, net cash flow after senior becomes 0 ie., there is nothing left to pay others.
                End If
            End If
            
            ' Now, if above 2 if statements are not run, it means our cash flow is not sufficient
            'to pay both IPMT and PPMT and thus we need to check reserve balances.
        End If
        
    ElseIf Sheet10.Range("L4") = 0 Or Sheet10.Range("L4") < 0 Then ' If Face value is less than 0 or =0, we pay nothing to senior tranche.
        senior_value = 0
    End If
    
    
    ' Outputing the senior cash flow to Waterfall table:
    
    waterfall_arr(month_i, 2) = senior_value
    'Sheet9.Range("L" & month_i + 16) = senior_value --> uncommented the follwowing to see the value in the worksheet.
    
    ' COMPUTING MEZZANINE CASH FLOW
    
    ' Checking if mezzanine face value >0?
    If Sheet10.Range("U4") > 0 And net_cash_flow_after_senior > 0 Then
        If net_cash_flow_after_senior > mezz_ipmt + mezz_ppmt Then ' We have enough cash flow to pay all the dues of mezzanine
            mezz_value = mezz_ipmt + mezz_ppmt
            net_cash_flow_after_mezz = net_cash_flow_after_senior - mezz_ipmt - mezz_ppmt
        Else ' Here we do not have enough cash flows to pay all dues so we first check our capacity to pay IPMT only and then we check or capacity to pay PPMT.
            If net_cash_flow_after_senior > mezz_ipmt Then
                mezz_value = mezz_value + mezz_ipmt
                net_cash_flow_after_mezz_ipmt = net_cash_flow_after_senior - mezz_ipmt ' deducting the ipmt value paid to mezzanine, then we use this value to see if we can pay PPMT.
                
            ElseIf net_cash_flow_after_senior < mezz_ipmt Then ' Here, we do not have enough cash flow to pay IPMT, so we check if we have any reserve balance to pay the shortfall.
                shortfall_mezz = mezz_ipmt - net_cash_flow_after_senior
                If reserve_balance >= shortfall_mezz Then
                    mezz_value = mezz_ipmt
                    reserve_balance = reserve_balance - shortfall_mezz
                    net_cash_flow_after_mezz_ipmt = 0
                ElseIf reserve_balance < shortfall_mezz Then ' In this case, we do have enough reserves to pay in full, so we pay all that we have
                    mezz_value = reserve_balance + net_cash_flow_after_senior
                    shortfall_mezz = shortfall_mezz - reserve_balance
                    reserve_balance = 0
                    net_cash_flow_after_mezz_ipmt = 0
                End If
            End If
            
            ' Now, after checking mezzanine IPMT, we are moving to mezzaninr PPMT.
            If net_cash_flow_after_mezz_ipmt > mezz_ppmt Then
                mezz_value = mezz_value + mezz_ppmt
                net_cash_flow_after_mezz = net_cash_flow_after_mezz_ipmt - mezz_ppmt
            ElseIf net_cash_flow_after_mezz_ipmt < mezz_ppmt Then ' Here, our current cash flow is not enough to pay mezz ppmt, so we check if we have any reserves.
                shortfall_mezz = shortfall_mezz + (mezz_ppmt - net_cash_flow_after_mezz_ipmt)
                If reserve_balance >= shortfall_mezz Then
                    mezz_value = mezz_value + mezz_ppmt
                    reserve_balance = reserve_balance - mezz_ppmt
                    net_cash_flow_after_mezz = 0
                ElseIf reserve_balance < mezz_ppmt Then ' We do have enough in reserves to pay so we pay whatever we can and reserve baance and cash flow becomes zero.
                    mezz_value = mezz_value + reserve_balance + net_cash_flow_after_mezz_ipmt
                    reserve_balance = 0
                    net_cash_flow_after_mezz = 0
                End If
            End If
        End If
    Else
        If Sheet10.Range("U4") = 0 Or Sheet10.Range("U4") < 0 Then
            mezz_value = 0 ' There are no dues to be paid
            reserve_balance = reserve_balance + net_cash_flow_after_senior
        End If
        If net_cash_flow_after_senior = 0 Then
            mezz_value = 0
            shortfall_mezz = mezz_ipmt + mezz_ppmt
        End If
    End If
    
    ' Outputting the mezzanine cash flow value to the table
    
    waterfall_arr(month_i, 3) = mezz_value
    'Sheet9.Range("M" & month_i + 16) = mezz_value --> uncommented the follwowing to see the value in the worksheet.
     
     
    equity_ipmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("AC11:AG371"), 4, False)
    equity_ppmt = Application.WorksheetFunction.VLookup(month_i, Sheet10.Range("AC11:AG371"), 3, False)
    
    If equity_ppmt < 0 Then ' Bounding the negative ppmt values of equity
        equity_ppmt = 0
    End If
    
    ' COMPUTING EQUITY CASH FLOW
    
    If Sheet10.Range("AD4") > 0 And net_cash_flow_after_mezz > 0 Then
        If equity_cap_remaining > 0 Then
            If net_cash_flow_after_mezz > equity_ipmt + equity_ppmt Then
                equity_value = Application.WorksheetFunction.Min(equity_ipmt + equity_ppmt, equity_cap_remaining)
                net_cash_flow_after_equity = net_cash_flow_after_mezz - equity_value
                reserve_balance = reserve_balance + net_cash_flow_after_equity
                net_cash_flow_after_equity = 0
                equity_cap_remaining = equity_cap_remaining - equity_ppmt - equity_ipmt
            Else
                If net_cash_flow_after_mezz > equity_ipmt Then ' Checking Equity IPMT value
                    equity_value = Application.WorksheetFunction.Min(equity_ipmt, equity_cap_remaining)
                    net_cash_flow_after_equity_ipmt = net_cash_flow_after_mezz - equity_value
                Else ' Since we do not have enough to pay in full all the amount towards equity ipmt, this shows an early sign of low cash flows coming in than expected, so we pay them only 80% and rest goes to reserves
                    equity_value = Application.WorksheetFunction.Min(net_cash_flow_after_mezz, equity_cap_remaining) * 0.8
                    reserve_balance = reserve_balance + (net_cash_flow_after_mezz - equity_value)
                    net_cash_flow_after_equity_ipmt = 0
                    equity_cap_remaining = equity_cap_remaining - equity_value
                End If
                
                ' Now that we have paid equity ipmt, we check if we have to pay equity ppmt
                If net_cash_flow_after_equity_ipmt > equity_ppmt Then
                    equity_value = equity_value + Application.WorksheetFunction.Min(equity_ppmt, equity_cap_remaining)
                    reserve_balance = reserve_balance + (net_cash_flow_after_equity_ipmt - equity_value)
                    net_cash_flow_after_equity = 0
                    equity_cap_remaining = equity_cap_remaining - equity_value
                ElseIf net_cash_flow_after_equity_ipmt <= equity_ppmt Then
                    equity_value = equity_value + Application.WorksheetFunction.Min(net_cash_flow_after_equity_ipmt, equity_cap_remaining)
                    net_cash_flow_after_equity = 0
                    equity_cap_remaining = equity_cap_remaining - equity_value
                End If
            End If
        ElseIf equity_cap_remaining < 0 Or equity_cap_remaining = 0 Then ' If equity cap has been reached, then equity gets nothing, all the balance goes to reserves.
                equity_value = 0
                reserve_balance = reserve_balance + net_cash_flow_after_mezz
        End If
    Else ' Since Equity face value has become 0, there is no payment made to equity holders.
        equity_value = 0
        reserve_balance = reserve_balance + net_cash_flow_after_mezz
    End If
    
    ' Just a safety check
    If net_cash_flow_after_equity > 0 Then
        reserve_balance = reserve_balance + net_cash_flow_after_equity
    End If
    
    ' Outputting equity cash flow to the table
    
    waterfall_arr(month_i, 4) = equity_value
    
    ' Outputtig the final reserve balance
    waterfall_arr(month_i, 5) = reserve_balance
    'Sheet9.Range("O" & month_i + 16) = reserve_balance --> uncommented the follwowing to see the value in the worksheet.
         
Next

Sheet9.Range("K17:O376") = waterfall_arr

Sheet9.Range("K17:O376").NumberFormat = "#,##0.00"
Sheet9.Range("O16:O376").Borders(xlEdgeRight).LineStyle = xlContinuous
Sheet9.Range("J376:O376").Borders(xlEdgeBottom).LineStyle = xlContinuous

Application.ScreenUpdating = True

End Sub

Function IsInArray(val As Variant, arr As Variant) As Boolean
 On Error GoTo ErrHandler
 Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If CLng(arr(i)) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
ErrHandler:
    IsInArray = False
End Function

Sub Bubblesort(arr As Variant)
Dim i As Long, j As Long, temp As Variant
For i = LBound(arr) To UBound(arr)
    For j = i + 1 To UBound(arr)
        If arr(i) > arr(j) Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next j
Next i
End Sub

Sub Solve()
Application.ScreenUpdating = False

Dim sen_weight As Variant, mezz_weight As Variant, eq_weight As Variant, total As Variant, targetIy As Variant, WAC As Variant

sen_weight = Sheet9.Range("L2").Value
mezz_weight = Sheet9.Range("L3").Value
eq_weight = Sheet9.Range("L4").Value

Sheet9.Range("L5").Formula = "=L2+L3+L4"

total = Sheet9.Range("L5").Value

' Assigning sumprodct formula to cell K8, Target I/Y:
Sheet9.Range("K8").Formula = "=SUMPRODUCT(K2:K3, L2:L3)"
'Sheet9.Range("K8").Formula = "=K6-K7" ' "=SUMPRODUCT(K2:K4, L2:L4)"
' Assigning WAC as a safety measure
Sheet9.Range("K6") = 6.65370837038667E-02

If IsNumeric(sen_weight) And IsNumeric(mezz_weight) And IsNumeric(eq_weight) Then
    If sen_weight > mezz_weight And sen_weight > eq_weight Then
        If total = 1 Then
            Solverreset
            ' Maximizing the target I/Y by changing Interest rate cells for each tranche.
            SolverOk SetCell:=Sheet9.Range("K8"), MaxMinVal:=1, byChange:="$K$2:$K$4"
            ' Adding constraint that Senior I/Y < Mezzanine I/Y
            SolverAdd CellRef:=Sheet9.Range("K2"), Relation:=1, FormulaText:=Sheet9.Range("K3")
            ' Adding constraint for Mezzanine Tranche's I/Y < Equity tranche I/Y :
            SolverAdd CellRef:=Sheet9.Range("K3"), Relation:=1, FormulaText:=Sheet9.Range("K4")
            ' Adding constraint that target I/Y cannot exceed WAC -soft credit enhancement
            SolverAdd CellRef:=Sheet9.Range("K8"), Relation:=1, FormulaText:=Sheet9.Range("K6").Value - Sheet9.Range("K7").Value
            
            SolverSolve UserFinish:=True
            SolverFinish KeepFinal:=1
            'Sheet9.Range("K7").NumberFormat = "0.00%"
            
            ' Post Solver Check
            targetIy = Sheet9.Range("K8").Value  '-> K7 has target I/Y
            WAC = Sheet9.Range("K6").Value     ' -> K9 stores WAC
            If targetIy > WAC Then
                targetIy = 0.0565
                'Sheet9.Range("K7").NumberFormat = "0.00%"
                Sheet9.Range("L2") = 0.5
                Sheet9.Range("L3") = 0.45
                Sheet9.Range("L4") = 0.05
                Sheet9.Range("L2:L4").NumberFormat = "0.00%"
                
                MsgBox ("Target I/Y computed was greater than WAC, thus the value has been reset to default")
            End If
        Else
            MsgBox ("Sum of tranche allocation must be 100%")
        End If
    Else
        MsgBox ("Senior Tranche should always have the highest allocation")
    End If
Else
   MsgBox ("Operation Failed. Please make sure Senior tranche % allocation is greater than that of mezzanine and the sum of the three tranches equals 100%")

End If

Sheet9.Range("K8").Formula = "=K6-K7"

Application.ScreenUpdating = True

End Sub

Sub reset_solver_values()

Application.ScreenUpdating = False

Sheet9.Range("L2") = 0.5
Sheet9.Range("L3") = 0.45
Sheet9.Range("L4") = 0.05

' Setting equity value = 1- senior-mezzanie
Sheet9.Range("L4") = 1 - Sheet9.Range("L2") - Sheet9.Range("L3")
Sheet9.Range("L2:L4").NumberFormat = "0.00%"
Sheet9.Range("K2") = 3.38306075424368E-02 ' Senior tranche I/Y
Sheet9.Range("K3") = 7.79415715944189E-02 ' Mezzanine Tranche I/Y
Sheet9.Range("K4") = 9.02197802256271E-02 ' Equity tranche I/Y
Sheet9.Range("K7") = 0.01 ' Soft credit enhancement
Sheet9.Range("K6") = 6.65370837038667E-02 ' Resetting WAC
' Assigning sumprodct formula to cell K7
Sheet9.Range("K8").Formula = "=SUMPRODUCT(K2:K4, L2:L4)"

' Resetting Pool value after hard credit enhancement
Sheet9.Range("K11").Formula = "=K10*(1-K9)"

' Resetting Face values
Sheet9.Range("M2").Formula = "=L2*K11"
Sheet9.Range("M3").Formula = "=L3*K11"
Sheet9.Range("M4").Formula = "=L4*K11"

Sheet9.Range("M5").Formula = "=SUM(M2:M4)"

Sheet9.Range("L5").Formula = "=SUM(L2:L4)"

Sheet9.Range("M2:M5").NumberFormat = "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "

' Data_Val Macro
    
    With Sheet9.Range("K9")
        .Value = 0 'Default value
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="0,0.05,0.1,0.15,0.2,0.25,0.3,0.35,0.4,0.45,0.5,0.55"
        .Validation.IgnoreBlank = False
        .Validation.InCellDropdown = True
        .Validation.InputTitle = "Hard credit enhancement"
        .Validation.ErrorTitle = ""
        .Validation.InputMessage = "A hard credit enhancement is the percentage of the pool value that you want to keep as a buffer to protect investors"
        .Validation.ErrorMessage = ""
        .Validation.ShowInput = True
        .Validation.ShowError = True
    End With
    
    With Sheet9.Range("K7")
        .Value = 0#   'Default value
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="0,0.005,0.01,0.015,0.02,0.03,0.04,0.05"
        .Validation.IgnoreBlank = False
        .Validation.InCellDropdown = True
        .Validation.InputTitle = "Soft credit enhancement"
        .Validation.ErrorTitle = ""
        .Validation.InputMessage = "A soft credit enhancement is the difference that you want to keep between the WAC and the weighted interest rate given to tranche investors"
        .Validation.ErrorMessage = ""
        .Validation.ShowInput = True
        .Validation.ShowError = True
    End With
    
Sheet9.Range("K10") = 10393354106#
Sheet9.Range("K10").NumberFormat = _
        "_-[$$-en-US]* #,##0.00_ ;_-[$$-en-US]* -#,##0.00 ;_-[$$-en-US]* ""-""??_ ;_-@_ "

Application.ScreenUpdating = True

End Sub

Sub Information_tab()

Application.ScreenUpdating = False

Worksheets("Model Information").Visible = True
Sheet10.Visible = True
Sheet10.Activate

Worksheets("Model Information").Activate
Sheet9.Visible = False

Application.ScreenUpdating = True

End Sub

Sub Closeinfo()

Application.ScreenUpdating = False

Worksheets("Model Information").Visible = xlVeryHidden
Sheet9.Visible = True
Sheet9.Activate

Sheet10.Visible = xlVeryHidden
Worksheets("Rough | backend").Visible = xlVeryHidden
Sheet2.Visible = xlVeryHidden

Application.ScreenUpdating = True

End Sub

Sub Clearall()

Application.ScreenUpdating = False

nobs = 40000
Sheet9.Range("W2:AD" & nobs).Clear
Sheet9.Range("AF2:AM" & nobs).Clear
Sheet9.Range("K17:O376").Clear

' Resetting Tranche_Amortization table:
Sheet10.Range("C4").Value = "='Waterfall Structure'!K10"     ' Portfolio's
Sheet10.Range("L4").Formula = "='Waterfall Structure'!M2"    ' Senior tranche
Sheet10.Range("U4").Formula = "='Waterfall Structure'!M3"    ' Mezzanine tranche
Sheet10.Range("AD4").Formula = "='Waterfall Structure'!M4"   ' Equity tranche

' Resetting Tranche's interest rates:
Sheet10.Range("C5").Formula = "='Waterfall Structure'!K6/12"   ' Portfolio's
Sheet10.Range("L5").Formula = "='Waterfall Structure'!K2/12"    ' Senior's
Sheet10.Range("U5").Formula = "='Waterfall Structure'!K3/12"      ' Mezzanine's
Sheet10.Range("AD5").Formula = "='Waterfall Structure'!K4/12"     ' Equity's


Sheet10.Range("AD4").NumberFormat = "#,##0"
Sheet10.Range("L4").NumberFormat = "#,##0"
Sheet10.Range("C4").NumberFormat = "#,##0"
Sheet10.Range("U4").NumberFormat = "#,##0"

Sheet9.Range("O16:O376").Borders(xlEdgeRight).LineStyle = xlContinuous
Sheet9.Range("J376:O376").Borders(xlEdgeBottom).LineStyle = xlContinuous

Call reset_solver_values
Call Closeinfo

Application.ScreenUpdating = True

End Sub

Sub tranche_reset()

Sheet10.Range("C4").Value = "='Waterfall Structure'!K10"     ' Portfolio's
Sheet10.Range("L4").Formula = "='Waterfall Structure'!M2"    ' Senior tranche
Sheet10.Range("U4").Formula = "='Waterfall Structure'!M3"    ' Mezzanine tranche
Sheet10.Range("AD4").Formula = "='Waterfall Structure'!M4"   ' Equity tranche

' Resetting Tranche's interest rates:
Sheet10.Range("C5").Formula = "='Waterfall Structure'!K6/12"   ' Portfolio's
Sheet10.Range("L5").Formula = "='Waterfall Structure'!K2/12"    ' Senior's
Sheet10.Range("U5").Formula = "='Waterfall Structure'!K3/12"      ' Mezzanine's
Sheet10.Range("AD5").Formula = "='Waterfall Structure'!K4/12"     ' Equity's


Sheet10.Range("AD4").NumberFormat = "#,##0"
Sheet10.Range("L4").NumberFormat = "#,##0"
Sheet10.Range("C4").NumberFormat = "#,##0"
Sheet10.Range("U4").NumberFormat = "#,##0"

Application.ScreenUpdating = True

End Sub

Sub Show_Calculation_sheet()
Application.ScreenUpdating = False

Sheet2.Visible = True
Sheet2.Activate
Sheet9.Visible = False
Sheet10.Visible = True

Application.ScreenUpdating = True

End Sub

Sub unhiding()

Sheet2.Visible = True
Sheet5.Visible = True
Sheet10.Visible = True

End Sub
