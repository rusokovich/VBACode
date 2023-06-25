
' This sub is to build the agents schedule

Public Function ScheduleArray() As Variant

Dim WeekDaysName() As Variant
Dim Schedule() As Variant
Dim ScheduleDays() As Variant
Dim i, J As Long
Dim a As String


J = 0
WeekDaysName() = InputSheet.Range("WeekDays")
Schedule() = InputSheet.Range("Schedule")


For i = 1 To 7

    If Not IsEmpty(Schedule(1, i)) Then
    ReDim Preserve ScheduleDays(J)
    ScheduleDays(J) = i
    J = J + 1
    End If

Next i

ScheduleArray = ScheduleDays


End Function


Public Function BenefitCalculation( _
ByVal PaymentCadence As String, ByVal Salary As Double _
, ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal StandardHours As Double = 0) As Double

Dim DaysCalculated, DaysScheduled, DailyHours As Double
Dim DaysArray As Variant

Dim Day As Variant
 
 
TotalWorkedDays = 0
AmountOfWorkedDays = 0

DaysCalculated = EndDate - StartDate

On Error Resume Next

DaysArray = ScheduleArray()
DaysScheduled = UBound(DaysArray) - LBound(DaysArray) + 1


DailyHours = StandardHours / DaysScheduled



If DaysCalculated < 0 Then

Debug.Print "Error is here ( Benefit )"

MsgBox "Check Dates"
Exit Function

End If


        Select Case PaymentCadence
        
            Case "Monthly"
            
                For Each Day In DaysArray
                
                AmountOfWorkedDays = HowManyDays(Month(StartDate), Year(StartDate), Day)
                
                TotalWorkedDays = AmountOfWorkedDays + TotalWorkedDays
                Next Day
                
                WritingTest ("Total month worked days: " & TotalWorkedDays)
                
            
                DailySalary = Salary / TotalWorkedDays
                
                
                
                DaysCalculated = 0
                 
                For Each Day In DaysArray
     
                
                AmountOfWorkedDays = WorkingDays(StartDate, EndDate - 1, Day)
                
                DaysCalculated = AmountOfWorkedDays + DaysCalculated
                Next Day
                
                                
                BenefitCalculation = DailySalary * DaysCalculated
                
                
                WritingTest ("Amount of worked days: " & DaysCalculated & " Daily Salary: " & DailySalary & " Benefit Calculation: " & BenefitCalculation)
                
                

            
            Case "Weekly"
            
                DailySalary = Salary * DailyHours
                
                For Each Day In DaysArray
                
                AmountOfWorkedDays = WorkingDays(StartDate, EndDate, Day)
                
                TotalWorkedDays = AmountOfWorkedDays + TotalWorkedDays
                Next Day
                
                                
                DaysCalculated = 0
                 
                For Each Day In DaysArray
                
                
                AmountOfWorkedDays = WorkingDays(StartDate, EndDate - 1, Day)
                
                DaysCalculated = AmountOfWorkedDays + DaysCalculated
                Next Day

                
                BenefitCalculation = DailySalary * DaysCalculated
                
                 WritingTest ("Amount of worked days: " & DaysCalculated & " Daily Salary: " & DailySalary & " Benefit Calculation: " & BenefitCalculation)
                
                

                
                
            Case "BiWeekly"
            
                
            
                DailySalary = Salary * DailyHours
                
                For Each Day In DaysArray
                
                AmountOfWorkedDays = WorkingDays(StartDate, EndDate, Day)
                
                TotalWorkedDays = AmountOfWorkedDays + TotalWorkedDays
                Next Day
                
                                
                DaysCalculated = 0
                 
                For Each Day In DaysArray
                
                
                AmountOfWorkedDays = WorkingDays(StartDate, EndDate - 1, Day)
                
                DaysCalculated = AmountOfWorkedDays + DaysCalculated
                Next Day

                
                BenefitCalculation = DailySalary * DaysCalculated
                
                
                WritingTest ("Amount of worked days: " & DaysCalculated & " Daily Salary: " & DailySalary & " Benefit Calculation: " & BenefitCalculation)
                

            Case Else
            
              MsgBox "Check Cadence"
              Exit Function
        
        End Select


                
End Function

Public Function DeductionCalculation(ByVal LeaveType As String _
, DeductionStartdate As Date, DeductionEndDate As Date, WeeklyDeduction As Double)

Dim DailyDeduction As Double
Dim DaysCalculated As Double
Dim AmountOfWorkedDays As Integer
Dim TotalWorkedDays As Integer

DaysCalculated = DeductionEndDate - DeductionStartdate



WritingTest ("Start Date Deduction: " & DeductionStartdate & " End Date Deduction: " & DeductionEndDate & " Distance Between Dates: " & DaysCalculated)

If DaysCalculated < 0 Then

WritingTest ("Error is here ( Deduction)")

DeductionCalculation = 0
Exit Function

End If


DailyDeduction = WeeklyDeduction / 7


Select Case LeaveType

    Case "Pre Partum"
    
        DeductionCalculation = DailyDeduction * DaysCalculated

    Case "Post Partum"
    
        DeductionCalculation = DailyDeduction * DaysCalculated

    Case "Parental"
        
        DeductionCalculation = DailyDeduction * DaysCalculated

    Case Else

        MsgBox "Check Leave"
        Exit Function

End Select

WritingTest ("Deduction Calculation " & DeductionCalculation & " Daily deduction: " & DailyDeduction & " Days Calculated: " & DaysCalculated)

End Function

'Method1: Next Sunday Date using Excel VBA Functions
Public Function NextSundayDate(ByVal InputDate As Date)

    
    NextSundayDate = DateAdd("d", -Weekday(InputDate) + 8, InputDate)


End Function


Public Function LastSundayDate(ByVal InputDate As Date)
    
    LastSundayDate = DateAdd("d", -Weekday(InputDate) + 1, InputDate)

End Function

Public Function GrossCalculation(ByVal Benefit As Double, ByVal Deduction As Double) As Double

GrossCalculation = Benefit - Deduction

WritingTest (" ")
WritingTest ("Deduction: " & Deduction)
WritingTest ("Benefit: " & Benefit)
'WritingTest ( GrossCalculation)

End Function

Public Function MinimumBenefitCalculation(ByVal LeaveType As String _
, MinimalBenefitStartdate As Date, MinimalBenefitEndDate As Date, MinimalBenefit As Double) As Double

Dim DailyMinimalBenefit As Double
Dim DaysCalculated As Double

DaysCalculated = MinimalBenefitEndDate - MinimalBenefitStartdate

If DaysCalculated < 0 Then

'MsgBox "Check Dates"
MinimumBenefitCalculation = 0
Exit Function

End If


DailyMinimalBenefit = MinimalBenefit / 7




Select Case LeaveType

    Case "Pre Partum"
    
        MinimumBenefitCalculation = DailyMinimalBenefit * DaysCalculated

    Case "Post Partum"
    
        MinimumBenefitCalculation = DailyMinimalBenefit * DaysCalculated


    Case "Parental"
        
        MinimumBenefitCalculation = DailyMinimalBenefit * DaysCalculated

    Case Else

        MsgBox "Check Leave"
        Exit Function

End Select


WritingTest ("Minimal Benefit: " & MinimalBenefit & " Start Date Minimal Benefit: " & MinimalBenefitStartdate & " End Date Minimal Benefit: " & MinimalBenefitEndDate & " Distance Between Dates: " & DaysCalculated & " Daily Minimal Benefit: " & DailyMinimalBenefit & " Minimum Benefit Calculation: " & MinimumBenefitCalculation)



End Function

Public Function HowManyDays(ByVal MonthNumber As Integer, ByVal YearNumber As Integer, ByVal Wday As Integer) As Integer

Dim i
Dim MyCount As Long
Dim StartDate As Date
Dim EndDate As Date


StartDate = DateSerial(YearNumber, MonthNumber, 1)

EndDate = CDate(Excel.Application.WorksheetFunction.EoMonth(StartDate, 0))

For i = StartDate To EndDate
    If Weekday(i, vbSunday) = Wday Then MyCount = MyCount + 1
Next i


HowManyDays = MyCount


End Function

Public Function WorkingDays(ByVal StartDate As Date, ByVal EndDate As Date, ByVal Wday As Integer) As Integer

Dim i
Dim MyCount As Long


For i = StartDate To EndDate
    If Weekday(i, vbSunday) = Wday Then MyCount = MyCount + 1
Next i


WorkingDays = MyCount

End Function


' Source: https://stackoverflow.com/questions/7015486/write-contents-of-immediate-window-to-a-text-file

Public Function WritingTest(ByVal phrase As String)

Dim FileName As String


FileName = "\Audit.txt"

n = FreeFile()

Open Application.ActiveWorkbook.Path & FileName For Append As #n

Print #n, phrase

Close #n

End Function



