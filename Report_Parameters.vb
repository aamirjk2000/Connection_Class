Imports System.Collections
Public Class Report_Parameters

    Public Const TestMessage As String = "10.200.2"

    '----------------------------------------------------- Report's Parameters
    Public Property CompanyName As String = ""                                      ' Compnay Name display at the top of the report.
    Public Property CompanyID As Integer = 0                                        ' Company ID, for Filter in SQL query
    Public Property CompanyAddress As String = ""                                   ' compnay Address display at the top area of Report
    Public Property ReportHeading As String = ""                                    ' Report Heading display at the top area of the report.

    Public Property SupplierID1 As Integer = 0                                      ' Supplier / Vender / Client ID start from 
    Public Property SupplierID2 As Integer = 0                                      ' Supplier / Vender / Client ID end at
    Public Property SupplierQueue As New Queue

    Public Property ProjectID1 As Integer = 0                                       ' Project ID start from 
    Public Property ProjectID2 As Integer = 0                                       ' Project ID End at
    Public Property ProjectQueue As New Queue                                       ' Selected Projects IDs, if 

    Public Property EmployeeID1 As Integer = 0                                      ' Employee ID start from 
    Public Property EmployeeID2 As Integer = 0                                      ' Employee ID End at
    Public Property EmployeeQueue As New Queue                                      ' Selected Projects IDs, if 

    Public Property COAID1 As Integer = 0                                           ' Chart of Accounts start from
    Public Property COAID2 As Integer = 0                                           ' Chart of Accounts end at
    Public Property COAQueue As New Queue                                           ' Selected Chart of Accounts IDs.

    Public Property DateFrom As Date = Now                                          ' Report Date start from
    Public Property DateTo As Date = Now                                            ' Report Date end at

    Public Property All_Projects As Boolean = True                                  ' Show All Projects in the Report
    Public Property All_Supplier As Boolean = True                                  ' show All Supplier / Vender / Client in the report
    Public Property All_Employee As Boolean = True                                  ' show All Emoployees in the report
    Public Property Ref_No As String = ""                                           ' Reference Number.
    Public Property ChequeNo As String = ""                                         ' Cheque number
    Public Property DateFormatText As String = ""                                   ' Date format
    Public Property AmountFormat As String = ""                                     ' Amount format
    Public Property ReportSort As String = ""                                       ' Report Sorting Order 

    '----------------------------------------------------------------------------   SQL

    Public Property SQLConnection As Data.SqlClient.SqlConnection
    Public Property SQLCommand As Data.SqlClient.SqlCommand                         ' SQL Parameters for SQL Quer
    Public Property SQLProcedure As String
    Public Property SQLAdapter As SqlClient.SqlDataAdapter
    Public Property SQLDataSet As DataSet

    'Public Property ReportProcedure As String                                       ' Name of SQL Users Procudure
    '---------------------------------------------------------- Report Property
    Public Property ReportName As String                                            ' Report's Name
    Public Property ReportPageAll As String                                         ' Reports Pages All
    Public Property ReportPagesQueue As Queue                                       ' Selected Pages Shows
    Public Property ReportFilter As String                                          ' Report Filter

    Public Property ShowProperties As Boolean                                       ' Show Reports Paramters beofre the report print.
    Public Property ShowMessages As Boolean                                         ' Show Messages at the end of the functions / Procedures.
    '---------------------------------------------------------- Printer Property
    Public Property PrinterName As String                                           ' Printer Name for use report printing.
    Public Property PrinterPort As String                                           ' Port of Printer
    Public Property ShowZeros As Boolean                                            ' Show Zero Amount record in report

    Public Sub New()

    End Sub


    Public ReadOnly Property FormatDate(_Date As DateTime) As String
        Get
            Return Format(_Date, DateFormatText)
        End Get
    End Property

    Public Function ShowReportParameters()

        Dim _Parameter As New ArrayList

        _Parameter.Add("Company Name      = " & CompanyName.ToString)
        _Parameter.Add("Company Address   = " & CompanyAddress.ToString)
        _Parameter.Add("Report Heading    = " & ReportHeading.ToString)
        _Parameter.Add("Company ID        = " & CompanyID.ToString)
        _Parameter.Add("Supplier ID 1     = " & SupplierID1.ToString)
        _Parameter.Add("Supplier ID 2     = " & SupplierID2.ToString)
        _Parameter.Add("Project ID 1      = " & ProjectID1.ToString)
        _Parameter.Add("Project ID 2      = " & ProjectID2.ToString)
        _Parameter.Add("Chart of Acc ID 1 = " & COAID1.ToString)
        _Parameter.Add("Chart of Acc ID 2 = " & COAID2.ToString)
        _Parameter.Add("Report Date From  = " & DateFrom.ToLongDateString)
        _Parameter.Add("Report Date To    = " & DateTo.ToLongDateString)

        _Parameter.Add("All Projects      = " & All_Projects.ToString)
        _Parameter.Add("All Suppliers     = " & All_Supplier.ToString)
        _Parameter.Add("Reference No      = " & Ref_No.ToString)
        _Parameter.Add("Cheque No.        = " & ChequeNo.ToString)
        _Parameter.Add("Date Format       = " & DateFormatText.ToString)
        _Parameter.Add("Amount Format     = " & AmountFormat.ToString)
        _Parameter.Add("Report Sort       = " & ReportSort.ToString)

        Return _Parameter

    End Function

End Class
