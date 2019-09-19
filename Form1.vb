Imports System
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Configuration

Public Class Form1
    Public conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Month As String = ComboBox1.Text
        Dim Year As String = ComboBox2.Text
        Dim JDate As Date = DateTimePicker1.Value

        Dim MonthVal As String
        MonthVal = 0

        If Month = "January" Then
            MonthVal = 1
        ElseIf Month = "February" Then
            MonthVal = 2
        ElseIf Month = "March" Then
            MonthVal = 3
        ElseIf Month = "April" Then
            MonthVal = 4
        ElseIf Month = "May" Then
            MonthVal = 5
        ElseIf Month = "June" Then
            MonthVal = 6
        ElseIf Month = "July" Then
            MonthVal = 7
        ElseIf Month = "August" Then
            MonthVal = 8
        ElseIf Month = "September" Then
            MonthVal = 9
        ElseIf Month = "October" Then
            MonthVal = 10
        ElseIf Month = "November" Then
            MonthVal = 11
        ElseIf Month = "December" Then
            MonthVal = 12
        End If

        ProgressBar1.Value = 10

        My.Computer.FileSystem.CreateDirectory("D:\Payroll " + Month + " " + Year)

        Dim Query As String = "SELECT ISNULL(SUM(dbo.Salary_Generation.Net_Amount),0), " & _
        "ISNULL(SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount),0), " & _
        "ISNULL(SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount),0), " & _
        "ISNULL(SUM(dbo.Salary_Generation.Goods_Amount),0), " & _
        "ISNULL(SUM(dbo.Salary_Generation.LST_Amount),0) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                      "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                      "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                      "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Research') AND " & _
                      "(dbo.Branch_Master.Branch_Name = 'ACBF')"

        Dim QueryCmd As New SqlCommand(Query, conn)
        Dim dataadapter As New SqlDataAdapter(QueryCmd)
        Dim ds As New DataSet()
        conn.Open()
        dataadapter.Fill(ds, "tblSalary")
        conn.Close()

        ProgressBar1.Value = 20

        Dim GOUR As String = "SELECT SUM(Net_Amount), SUM([5%]), SUM([10%]), SUM(PAYE), SUM(LST_Amount) " & _
"FROM dbo.ppiGOU_Research WHERE (Month = " + MonthVal + ") AND (Year = " + Year + ")"

        Dim GOURCmd As New SqlCommand(GOUR, conn)
        Dim GOURda As New SqlDataAdapter(GOURCmd)
        Dim GOURds As New DataSet()
        conn.Open()
        GOURda.Fill(ds, "tblSalaryGOUR")
        conn.Close()

        ProgressBar1.Value = 30


        Dim TTIR As String = "SELECT SUM(Net_Amount), SUM([5%]), SUM([10%]), SUM(PAYE), SUM(LST_Amount) " & _
"FROM dbo.ppiTTI_Research WHERE (Month = " + MonthVal + ") AND (Year = " + Year + ")"

        Dim TTIRCmd As New SqlCommand(TTIR, conn)
        Dim TTIRda As New SqlDataAdapter(TTIRCmd)
        Dim TTIRds As New DataSet()
        conn.Open()
        TTIRda.Fill(ds, "tblSalaryTTIR")
        conn.Close()

        ProgressBar1.Value = 40


        Dim ACBFS As String = "SELECT ISNULL(SUM(dbo.Salary_Generation.Net_Amount),0)," & _
        "ISNULL(SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount),0)," & _
        "ISNULL(SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount),0)," & _
        "ISNULL(SUM(dbo.Salary_Generation.Goods_Amount),0), ISNULL(SUM(dbo.Salary_Generation.LST_Amount),0) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                    "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                    "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                    "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Support') AND " & _
                    "(dbo.Branch_Master.Branch_Name = 'ACBF')"

        Dim ACBFSCmd As New SqlCommand(ACBFS, conn)
        Dim ACBFSda As New SqlDataAdapter(ACBFSCmd)
        Dim ACBFSds As New DataSet()
        conn.Open()
        ACBFSda.Fill(ds, "tblSalaryACBFS")
        conn.Close()


        ProgressBar1.Value = 50

        Dim GOUS As String = "SELECT SUM(dbo.Salary_Generation.Net_Amount), SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount), SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount), SUM(dbo.Salary_Generation.Goods_Amount), SUM(dbo.Salary_Generation.LST_Amount) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                  "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                  "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                  "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Support') AND " & _
                  "(dbo.Branch_Master.Branch_Name = 'GOU')"

        Dim GOUSCmd As New SqlCommand(GOUS, conn)
        Dim GOUSda As New SqlDataAdapter(GOUSCmd)
        Dim GOUSds As New DataSet()
        conn.Open()
        GOUSda.Fill(ds, "tblSalaryGOUS")
        conn.Close()


        ProgressBar1.Value = 60

        Dim TTIS As String = "SELECT SUM(dbo.Salary_Generation.Net_Amount), SUM(0.05 * dbo.Salary_Generation.Total_Earn_Amount), SUM(0.1 * dbo.Salary_Generation.Total_Earn_Amount), SUM(dbo.Salary_Generation.Goods_Amount), SUM(dbo.Salary_Generation.LST_Amount) " & _
"FROM dbo.Branch_Master INNER JOIN " & _
                     "dbo.Increment ON dbo.Branch_Master.Branch_ID = dbo.Increment.Branch_ID INNER JOIN " & _
                     "dbo.Salary_Generation ON dbo.Increment.Increment_Id = dbo.Salary_Generation.Increment_ID INNER JOIN " & _
                     "dbo.PRODUCT_MASTER ON dbo.Increment.Product_ID = dbo.PRODUCT_MASTER.Product_ID " & _
"WHERE (dbo.Salary_Generation.Month = " + MonthVal + ") AND (dbo.Salary_Generation.Year = " + Year + ") AND (dbo.PRODUCT_MASTER.Product_Name = 'Support') AND " & _
                     "(dbo.Branch_Master.Branch_Name = 'TTI')"

        Dim TTISCmd As New SqlCommand(TTIS, conn)
        Dim TTISda As New SqlDataAdapter(TTISCmd)
        Dim TTISds As New DataSet()
        conn.Open()
        TTISda.Fill(ds, "tblSalaryTTIS")
        conn.Close()


        ProgressBar1.Value = 70

        Dim Gratuity As String = "SELECT round(SUM(0.11 * dbo.GRATUITY.Basic),0)" & _
"FROM dbo.GRATUITY WHERE (dbo.GRATUITY.Month = " + MonthVal + ") AND (dbo.GRATUITY.Year = " + Year + ")"

        Dim GratuityCmd As New SqlCommand(Gratuity, conn)
        Dim GratuityDa As New SqlDataAdapter(GratuityCmd)
        Dim GratuityDs As New DataSet()
        conn.Open()
        GratuityDa.Fill(GratuityDs, "tblGratuity")
        conn.Close()

        Dim EDLoanDeduction As String = "SELECT 0.5 * dbo.Emp_Loan_Payment_Detail.Amount AS HalfLoan " & _
"FROM dbo.Emp_Loan_Payment INNER JOIN " & _
"dbo.Emp_Loan_Payment_Detail ON dbo.Emp_Loan_Payment.Loan_Payment_Id = dbo.Emp_Loan_Payment_Detail.Loan_Payment_Id " & _
"WHERE (dbo.Emp_Loan_Payment.Emp_Id = 1) AND (DATEPART(MM, dbo.Emp_Loan_Payment.Payment_Date) = " + MonthVal + ")"

        Dim EDLoanDeductionCmd As New SqlCommand(EDLoanDeduction, conn)
        Dim dsEDLoanDedu As New DataSet()
        Dim daEDLoanDedu As New SqlDataAdapter(EDLoanDeductionCmd)
        conn.Open()
        daEDLoanDedu.Fill(dsEDLoanDedu, "tblEDLoan")
        conn.Close()

        Dim EDHalfLoan As Decimal
        Dim recordCount = dsEDLoanDedu.Tables("tblEDLoan").Rows.Count
        If recordCount = 0 Then
            EDHalfLoan = 0
        Else
            EDHalfLoan = dsEDLoanDedu.Tables("tblEDLoan").Rows(0).Item(0)
        End If

        '-----------------------NET PAY-------------------------------

        Dim NETpath As String = "D:\Payroll " + Month + " " + Year + "\Net Pay.txt"
        Dim fs As FileStream = File.Create(NETpath)

        Dim NetLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "Net Pay For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        fs.Write(NetLine1, 0, NetLine1.Length)

        Dim NetLine2 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + ds.Tables("tblSalary").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine2, 0, NetLine2.Length)

        Dim NetLine3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine3, 0, NetLine3.Length)

        Dim GOURNet As Decimal = ds.Tables("tblSalaryGOUR").Rows(0).Item(0)
        Dim GOURNetNew As Decimal = (GOURNet - EDHalfLoan).ToString


        Dim NetLine4 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + GOURNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine4, 0, NetLine4.Length)

        Dim NetLine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + GOURNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine5, 0, NetLine5.Length)

        Dim TTIRNet As Decimal = ds.Tables("tblSalaryTTIR").Rows(0).Item(0)
        Dim TTIRNetNew As Decimal = (TTIRNet + EDHalfLoan).ToString


        Dim NetLine6 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + TTIRNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine6, 0, NetLine6.Length)

        Dim NetLine7 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + TTIRNetNew.ToString + Environment.NewLine)
        fs.Write(NetLine7, 0, NetLine7.Length)


        Dim NetLine8 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + ds.Tables("tblSalaryACBFS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine8, 0, NetLine8.Length)

        Dim NetLine9 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryACBFS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine9, 0, NetLine9.Length)


        Dim NetLine10 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine10, 0, NetLine10.Length)

        Dim NetLine11 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine11, 0, NetLine11.Length)


        Dim NetLine12 As Byte() = New UTF8Encoding(True).GetBytes("92000000," + "-" + ds.Tables("tblSalaryTTIS").Rows(0).Item(0).ToString + Environment.NewLine)
        fs.Write(NetLine12, 0, NetLine12.Length)

        Dim NetLine13 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryTTIS").Rows(0).Item(0).ToString)
        fs.Write(NetLine13, 0, NetLine13.Length)
        fs.Close()

        '----------------------PAYE----------------------------


        Dim PAYEPath As String = "D:\Payroll " + Month + " " + Year + "\PAYE.txt"
        Dim PAYEFs As FileStream = File.Create(PAYEPath)

        Dim PAYELine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "PAYE For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        PAYEFs.Write(PAYELine1, 0, PAYELine1.Length)

        Dim PAYELine2 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalary").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine2, 0, PAYELine2.Length)

        Dim PAYELine3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine3, 0, PAYELine3.Length)


        Dim PAYELine4 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine4, 0, PAYELine4.Length)

        Dim PAYELine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine5, 0, PAYELine5.Length)


        Dim PAYELine6 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryTTIR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine6, 0, PAYELine6.Length)

        Dim PAYELine7 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryTTIR").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine7, 0, PAYELine7.Length)


        Dim PAYELine8 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryACBFS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine8, 0, PAYELine8.Length)

        Dim PAYELine9 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryACBFS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine9, 0, PAYELine9.Length)


        Dim PAYELine10 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine10, 0, PAYELine10.Length)

        Dim PAYELine11 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine11, 0, PAYELine11.Length)


        Dim PAYELine12 As Byte() = New UTF8Encoding(True).GetBytes("93000001," + "-" + ds.Tables("tblSalaryTTIS").Rows(0).Item(3).ToString + Environment.NewLine)
        PAYEFs.Write(PAYELine12, 0, PAYELine12.Length)

        Dim PAYELine13 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryTTIS").Rows(0).Item(3).ToString)
        PAYEFs.Write(PAYELine13, 0, PAYELine13.Length)
        PAYEFs.Close()

        ProgressBar1.Value = 80
        '-------------------------------NSSF---------------------------

        Dim NSSFPath As String = "D:\Payroll " + Month + " " + Year + "\NSSF.txt"
        Dim NSSFFs As FileStream = File.Create(NSSFPath)

        Dim NSSFLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "NSSF For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        NSSFFs.Write(NSSFLine1, 0, NSSFLine1.Length)

        Dim NSSF5Line2 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalary").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line2, 0, NSSF5Line2.Length)

        Dim NSSF5Line3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line3, 0, NSSF5Line3.Length)


        Dim NSSF10Line4 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalary").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line4, 0, NSSF10Line4.Length)

        Dim NSSF10Line5 As Byte() = New UTF8Encoding(True).GetBytes("20000007," + ds.Tables("tblSalary").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line5, 0, NSSF10Line5.Length)


        Dim NSSF5Line6 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line6, 0, NSSF5Line6.Length)

        Dim NSSF5Line7 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUR").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line7, 0, NSSF5Line7.Length)


        Dim NSSF10Line8 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line8, 0, NSSF10Line8.Length)

        Dim NSSF10Line9 As Byte() = New UTF8Encoding(True).GetBytes("20000007," + ds.Tables("tblSalaryGOUR").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line9, 0, NSSF10Line9.Length)


        Dim NSSF5Line10 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryTTIR").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line10, 0, NSSF5Line10.Length)

        Dim NSSF5Line11 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryTTIR").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line11, 0, NSSF5Line11.Length)


        Dim NSSF10Line12 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryTTIR").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line12, 0, NSSF10Line12.Length)

        Dim NSSF10Line13 As Byte() = New UTF8Encoding(True).GetBytes("20000007," + ds.Tables("tblSalaryTTIR").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line13, 0, NSSF10Line13.Length)


        Dim NSSF5Line14 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryACBFS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line14, 0, NSSF5Line14.Length)

        Dim NSSF5Line15 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryACBFS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line15, 0, NSSF5Line15.Length)


        Dim NSSF10Line16 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryACBFS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line16, 0, NSSF10Line16.Length)

        Dim NSSF10Line17 As Byte() = New UTF8Encoding(True).GetBytes("24000007," + ds.Tables("tblSalaryACBFS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line17, 0, NSSF10Line17.Length)


        Dim NSSF5Line18 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line18, 0, NSSF5Line18.Length)

        Dim NSSF5Line19 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line19, 0, NSSF5Line19.Length)


        Dim NSSF10Line20 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line20, 0, NSSF10Line20.Length)

        Dim NSSF10Line21 As Byte() = New UTF8Encoding(True).GetBytes("24000007," + ds.Tables("tblSalaryGOUS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line21, 0, NSSF10Line21.Length)


        Dim NSSF5Line22 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryTTIS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line22, 0, NSSF5Line22.Length)

        Dim NSSF5Line23 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryTTIS").Rows(0).Item(1).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF5Line23, 0, NSSF5Line23.Length)


        Dim NSSF10Line24 As Byte() = New UTF8Encoding(True).GetBytes("93000002," + "-" + ds.Tables("tblSalaryTTIS").Rows(0).Item(2).ToString + Environment.NewLine)
        NSSFFs.Write(NSSF10Line24, 0, NSSF10Line24.Length)

        Dim NSSF10Line25 As Byte() = New UTF8Encoding(True).GetBytes("24000007," + ds.Tables("tblSalaryTTIS").Rows(0).Item(2).ToString)
        NSSFFs.Write(NSSF10Line25, 0, NSSF10Line25.Length)
        NSSFFs.Close()

        '---------------------------LST-----------------------------------

        ProgressBar1.Value = 90

        Dim LSTPath As String = "D:\Payroll " + Month + " " + Year + "\LST.txt"
        Dim LSTFs As FileStream = File.Create(LSTPath)

        Dim LSTLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "LST For " + Month + " " + Year + Chr(34) + Environment.NewLine)
        LSTFs.Write(LSTLine1, 0, LSTLine1.Length)

        Dim LSTLine2 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalary").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine2, 0, LSTLine2.Length)

        Dim LSTLine3 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalary").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine3, 0, LSTLine3.Length)


        Dim LSTLine4 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryGOUR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine4, 0, LSTLine4.Length)

        Dim LSTLine5 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryGOUR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine5, 0, LSTLine5.Length)


        Dim LSTLine6 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryTTIR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine6, 0, LSTLine6.Length)

        Dim LSTLine7 As Byte() = New UTF8Encoding(True).GetBytes("20000005," + ds.Tables("tblSalaryTTIR").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine7, 0, LSTLine7.Length)


        Dim LSTLine8 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryACBFS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine8, 0, LSTLine8.Length)

        Dim LSTLine9 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryACBFS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine9, 0, LSTLine9.Length)


        Dim LSTLine10 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryGOUS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine10, 0, LSTLine10.Length)

        Dim LSTLine11 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryGOUS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine11, 0, LSTLine11.Length)


        Dim LSTLine12 As Byte() = New UTF8Encoding(True).GetBytes("93000003," + "-" + ds.Tables("tblSalaryTTIS").Rows(0).Item(4).ToString + Environment.NewLine)
        LSTFs.Write(LSTLine12, 0, LSTLine12.Length)

        Dim LSTLine13 As Byte() = New UTF8Encoding(True).GetBytes("24000005," + ds.Tables("tblSalaryTTIS").Rows(0).Item(4).ToString)
        LSTFs.Write(LSTLine13, 0, LSTLine13.Length)
        LSTFs.Close()

        ProgressBar1.Value = 100


        '---------------------------Gratuity--------------------------


        Dim GratuityPath As String = "D:\Payroll " + Month + " " + Year + "\Gratuity.txt"
        Dim GratuityFs As FileStream = File.Create(GratuityPath)

        Dim GratuityLine1 As Byte() = New UTF8Encoding(True).GetBytes(JDate.ToString("MM-dd-yy") + "," + Chr(34) + "7" + Chr(34) + "," + Chr(34) + "Gratuity " + Month + " " + Year + Chr(34) + Environment.NewLine)
        GratuityFs.Write(GratuityLine1, 0, GratuityLine1.Length)

        Dim GratuityLine2 As Byte() = New UTF8Encoding(True).GetBytes("94000000," + "-" + GratuityDs.Tables("tblGratuity").Rows(0).Item(0).ToString + Environment.NewLine)
        GratuityFs.Write(GratuityLine2, 0, GratuityLine2.Length)

        Dim GratuityLine3 As Byte() = New UTF8Encoding(True).GetBytes("24000010," + GratuityDs.Tables("tblGratuity").Rows(0).Item(0).ToString)
        GratuityFs.Write(GratuityLine3, 0, GratuityLine3.Length)
        GratuityFs.Close()

        Application.Exit()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0

        ComboBox1.Text = "April"
        ComboBox2.Text = "2014"
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class
