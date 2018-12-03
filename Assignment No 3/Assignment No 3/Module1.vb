Module Module1
    ' Business View
    ' This is a payroll program for the Gleaming Crystall Company. It adjusts the pay rates by the employee's years of experience. For example one year of experience increases
    ' the pay by two percent. It only selects Glass Buyers, Blowers, and Cutters using the first digit of their employee number.

    ' Classroom View 
    ' Nested If
    ' If CONDITION Then
    '   Console.WriteLine("True")
    ' Else
    '     If OTHERCONDITION Then
    '           Console.WriteLine(1)
    '     End If
    ' End If

    ' Switch or a Case Statement
    ' Select Case VariableInteger
    '      Case 1
    '           Console.WriteLine("True")
    '      Case Else
    '           Console.WriteLine("False")
    ' End Select
    ' Data File (Input Variables)
    Private EmpHoursWorkedDecimal As Decimal

    Private EmpNumberInteger As Integer
    Private EmpYearsWorkedInteger As Integer
    Private ExemptionCountInteger As Integer

    Private EmpNameString As String

    ' Printer Spacing Chart (OUTPUT Variables)
    Private BasePayRateDecimal As Decimal
    Private ActualPayRateDecimal As Decimal
    Private GrossPayDecimal As Decimal
    Private ExemptionAmountDeductionDecimal As Decimal
    Private TaxableIncomeDecimal As Decimal
    Private TaxAmountDecimal As Decimal
    Private NetPayDecimal As Decimal
    Private RegPayDecimal As Decimal
    Private OTPayDecimal As Decimal

    Private PayRatePercentIncreaseInteger As Integer

    Private JobNameString As String

    Private Const NORMAL_HOURS_Integer As Integer = 40


    ' Pay Rate Percent    YOE = Years Of Experience
    Private Const YOE_TIER_ONE_Decimal As Decimal = 0.02
    Private Const YOE_TIER_TWO_Decimal As Decimal = 0.05
    Private Const YOE_TIER_THREE_Decimal As Decimal = 0.1
    Private Const YOE_TIER_FOUR_Decimal As Decimal = 0.15
    Private Const YOE_TIER_FIVE_Decimal As Decimal = 0.2
    Private Const YOE_TIER_SIX_Decimal As Decimal = 0.25
    Private Const YOE_TIER_SEVEN_Decimal As Decimal = 0.3

    ' Job Names
    Private Const JOB_NAME_BUYER_String As String = "Buyer"
    Private Const JOB_NAME_MELTER_String As String = "Melter"
    Private Const JOB_NAME_BLOWER_String As String = "Blower"
    Private Const JOB_NAME_MOLDER_String As String = "Molder"
    Private Const JOB_NAME_CUTTER_String As String = "Cutter"

    ' Job Pay Rate
    Private Const BASE_PAY_RATE_ONE_Decimal As Decimal = 15.5
    Private Const BASE_PAY_RATE_TWO_Decimal As Decimal = 14.9
    Private Const BASE_PAY_RATE_THREE_Decimal As Decimal = 19.0
    Private Const BASE_PAY_RATE_FOUR_Decimal As Decimal = 16.0
    Private Const BASE_PAY_RATE_FIVE_Decimal As Decimal = 25.5


    ' Accumulated Variables
    Private AccumNumOfEmployeesInteger As Integer = 0
    Private AccumGrossPayDecimal As Decimal = 0
    Private AccumTaxesDecimal As Decimal = 0
    Private AccumNetPayDecimal As Decimal = 0

    Private AccumPageCounterInteger As Integer = 1
    Private AccumLineCounterInteger As Integer = 20

    Private CurrentRecord() As String

    Private ValidRecordBoolean As Boolean

    Private CompanyFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("CRYSTALF18.TXT")

    Sub Main()
        Call HouseKeeping()
        Do While Not CompanyFile.EndOfData()
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub


    ' Tier Two
    Sub HouseKeeping()
        Call SetFileDelimiter()
    End Sub


    Sub ProcessRecords()
        Call ReadFile()
        Call RecordSelection()
        If ValidRecordBoolean = True Then
            Call DetailCalculation()
            Call AccumulateTotals()
            Call WriteDetailLines()
        End If
    End Sub


    Sub EndOfJob()
        Call SummaryOutput()
        Call CloseFile()
    End Sub


    ' Tier Three
    Sub SetFileDelimiter()
        CompanyFile.TextFieldType = FileIO.FieldType.Delimited
        CompanyFile.SetDelimiters(",")
    End Sub


    Sub ReadFile()
        ' Current Record Read for Info
        CurrentRecord = CompanyFile.ReadFields()

        ' Employee Number
        EmpNumberInteger = CurrentRecord(0)

        ' Employee Name
        EmpNameString = CurrentRecord(1)

        ' Years Worked by Employee
        EmpYearsWorkedInteger = CurrentRecord(2)

        ' Count of Employee Exemptions
        ExemptionCountInteger = CurrentRecord(3)

        ' Hours Worked by Employee
        EmpHoursWorkedDecimal = CurrentRecord(4)
    End Sub

    Sub RecordSelection()
        ValidRecordBoolean = False
        If EmpNumberInteger.ToString().Chars(0) = "5" Then
            JobNameString = JOB_NAME_CUTTER_String
            BasePayRateDecimal = BASE_PAY_RATE_FIVE_Decimal
            ValidRecordBoolean = True
        Else
            If EmpNumberInteger.ToString().Chars(0) = "3" Then
                JobNameString = JOB_NAME_BLOWER_String
                BasePayRateDecimal = BASE_PAY_RATE_THREE_Decimal
                ValidRecordBoolean = True
            Else
                If EmpNumberInteger.ToString().Chars(0) = "1" Then
                    JobNameString = JOB_NAME_BUYER_String
                    BasePayRateDecimal = BASE_PAY_RATE_ONE_Decimal
                    ValidRecordBoolean = True
                End If
            End If
        End If
    End Sub

    Sub DetailCalculation()

        Call PercentModifier()
        Call OvertimeModifier()

        GrossPayDecimal = RegPayDecimal + OTPayDecimal

        ExemptionAmountDeductionDecimal = ExemptionCountInteger * 20.0

        TaxableIncomeDecimal = GrossPayDecimal - ExemptionAmountDeductionDecimal

        TaxAmountDecimal = TaxableIncomeDecimal * 0.1

        NetPayDecimal = GrossPayDecimal - TaxAmountDecimal
    End Sub

    Sub AccumulateTotals()

        ' Counts Lines Until Reset
        AccumLineCounterInteger += 1

        ' Counts Number Of Employees Printed
        AccumNumOfEmployeesInteger += 1

        ' Adds Employee Gross Pay to Total Gross Pay
        AccumGrossPayDecimal += GrossPayDecimal

        ' Adds Employee Taxes to Total Taxes
        AccumTaxesDecimal += TaxAmountDecimal

        ' Adds Employee Net Pay to Total Net Pay
        AccumNetPayDecimal += NetPayDecimal

    End Sub

    Sub WriteDetailLines()
        If AccumLineCounterInteger >= 15 Then
            Call Paginate()
        End If

        Console.WriteLine()
        Console.WriteLine(EmpNumberInteger & Space(1) & JobNameString.PadRight(6) & Space(1) & EmpHoursWorkedDecimal.ToString.PadLeft(4) &
                          Space(2) & EmpYearsWorkedInteger.ToString().PadLeft(2) & Space(1) & BasePayRateDecimal.ToString("N2").PadLeft(5) &
                          Space(1) & PayRatePercentIncreaseInteger.ToString().PadLeft(2) & Space(2) & ActualPayRateDecimal.ToString("N2").PadLeft(5) &
                          Space(2) & GrossPayDecimal.ToString("N2").PadLeft(8) & Space(1) & ExemptionAmountDeductionDecimal.ToString("N2").PadLeft(6) &
                          Space(2) & TaxableIncomeDecimal.ToString("N2").PadLeft(8) & Space(1) & TaxAmountDecimal.ToString("N2").PadLeft(6) &
                          Space(2) & NetPayDecimal.ToString("N2").PadLeft(8))
        Console.WriteLine(Space(37) & "(Reg Pay: " & RegPayDecimal.ToString("N2").PadLeft(8) & "  OT Pay: " & OTPayDecimal.ToString("N2").PadLeft(8) & ")")

    End Sub

    Sub SummaryOutput()

        ' Printer Spacing Chart Line 14
        Console.WriteLine("FINAL TOTALS:")

        ' Printer Spacing Chart Line 15
        Console.WriteLine(Space(5) & "Number Of Employees" & Space(16) & AccumNumOfEmployeesInteger.ToString().PadLeft(2))

        ' Printer Spacing Chart Line 16
        Console.WriteLine(Space(5) & "Total Gross Pay" & Space(14) & AccumGrossPayDecimal.ToString("C").PadLeft(11))

        ' Printer Spacing Chart Line 17
        Console.WriteLine(Space(5) & "Total Taxes" & Space(19) & AccumTaxesDecimal.ToString("C").PadLeft(10))

        ' Printer Spacing Chart Line 18
        Console.WriteLine(Space(5) & "Total Net Pay" & Space(16) & AccumNetPayDecimal.ToString("C").PadLeft(11))

    End Sub

    Sub CloseFile()
        Console.WriteLine()
        Console.WriteLine(Space(1) & "Press -ENTER- To Exit")
        Console.WriteLine()
        Console.ReadLine()
    End Sub



    Sub PercentModifier()
        ActualPayRateDecimal = BasePayRateDecimal
        Select Case EmpYearsWorkedInteger
            Case 1
                PayRatePercentIncreaseInteger = 2
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case 2 To 4
                PayRatePercentIncreaseInteger = 5
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case 5 To 6
                PayRatePercentIncreaseInteger = 10
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case 7 To 10
                PayRatePercentIncreaseInteger = 15
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case 11 To 18
                PayRatePercentIncreaseInteger = 20
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case 19 To 25
                PayRatePercentIncreaseInteger = 25
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case > 25
                PayRatePercentIncreaseInteger = 30
                ActualPayRateDecimal += BasePayRateDecimal * (PayRatePercentIncreaseInteger / 100)
            Case Else
                PayRatePercentIncreaseInteger = 0
        End Select
    End Sub

    Sub OvertimeModifier()
        RegPayDecimal = 0
        OTPayDecimal = 0
        If EmpHoursWorkedDecimal <= 40.0 Then
            OTPayDecimal = 0
            RegPayDecimal = ActualPayRateDecimal * EmpHoursWorkedDecimal
        Else
            If EmpHoursWorkedDecimal > 40.0 Then
                OTPayDecimal = 1.5 * ((EmpHoursWorkedDecimal - NORMAL_HOURS_Integer) * ActualPayRateDecimal)
                RegPayDecimal = NORMAL_HOURS_Integer * ActualPayRateDecimal
            End If
        End If
    End Sub

    Sub Paginate()
        AccumLineCounterInteger = 0
        If AccumPageCounterInteger > 1 Then
            Call AddSpacing()
        End If
        Call WriteHeadings()
        AccumPageCounterInteger += 1
    End Sub

    Sub WriteHeadings()

        ' Printer Spacing Chart Line 01
        Console.WriteLine(Space(28) & "Gleaming Crystals Company" & Space(19) & "Page: " & AccumPageCounterInteger)

        ' Printer Spacing Chart Line 02
        Console.WriteLine(Space(26) & "Payroll Report By: David Rees")

        ' Printer Spacing Chart Line 03
        Console.WriteLine()

        ' Printer Spacing Chart Line 04   Add Spaces
        Console.WriteLine("------Employee------" & Space(2) & "---Pay Rates--" & Space(5) &
                          "Gross" & Space(2) & "Exmpt" & Space(3) &
                          "Taxable" & Space(4) & "Tax" & Space(7) &
                          "Net")

        ' Printer Spacing Chart Line 05   Add Spaces
        Console.WriteLine(Space(2) & "#" & Space(2) & "Job" &
                          Space(3) & "Hours" & Space(1) & "Yrs" &
                          Space(2) & "Base" & Space(2) & "%" &
                          Space(1) & "Actual" & Space(7) & "Pay" &
                          Space(4) & "Amt" & Space(4) & "Income" &
                          Space(4) & "Amt" & Space(7) & "Pay")

        ' Printer Spacing Chart Line 06
        Console.WriteLine()

    End Sub

    Sub AddSpacing()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
    End Sub
End Module
