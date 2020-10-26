Imports System
Imports System.IO
Imports System.Xml
Imports Microsoft.Office.Interop
Imports System.Threading
'Imports System.Timers
Imports System.Diagnostics
Imports System.Data.SqlClient

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
    Sub scheduler()
        If GlobalVariables.Runtype = True Then
            Dim ci As Integer = 0
            For ci = 1 To 10000
                Me.ListBox1.Items.Insert(0, "Starting Auto Loop")
                Call ProcessDCR()
                Me.ListBox1.Items.Insert(0, "Waiting for next run:" + DateAdd(DateInterval.Hour, 1, Now()).ToString)
                Me.ListBox1.Refresh()
                Thread.Sleep(3600000)
                'Thread.Sleep(10000)
                Me.Refresh()
                ci = ci + 1
            Next
        Else
            Call ProcessDCR()

        End If
    End Sub
    Private Sub ProcessDCR()
        ' End_Excel_App()
        Dim oBook As Excel.Workbook
        Dim oApp As Excel.Application
        Dim oSheets As Excel.Worksheet
        Dim oRange As Excel.Range
        oApp = CreateObject("Excel.Application")
        Dim vC1 As String
        Dim vSheetCount, ci_dol, ci2, ci_paint, ci_RM, ci_cap As Integer
        Dim DOPLOG, DOPLOG_SUM, MH, PB, AB, LST, DRT, DR1, DR2, DR3, DR4, DR5, DR6, DR7, DR8, DR9, DR10, DR11, DR12, DR13, DR14, TDC, HAZ, RMDRT, RMDRT_CUM, PDRT, PDRT_CUM As Integer
        Dim AOLV, ABV As Integer
        Dim vFile, vFile2, vPath As String
        Dim strDate As String
        Dim regDate As Date
        Dim par1, par2 As Decimal
        'Dim timer As System.Timers.Timer = New System.Timers.Timer(200)
        '
        '

        ' Bring files local
        Dim tempPath As String = System.IO.Path.GetTempPath()
        Dim di0 As New IO.DirectoryInfo("\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR")
        Dim aryFi0 As IO.FileInfo() = di0.GetFiles("*.xlsm")
        Dim fi0 As IO.FileInfo

        Dim vtempfolder As String = tempPath + "DCR\"
        If (Not System.IO.Directory.Exists(vtempfolder)) Then
            System.IO.Directory.CreateDirectory(vtempfolder)
        End If

        For Each fi0 In aryFi0

            IO.File.Move(fi0.FullName, Replace(fi0.FullName, "\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR\", tempPath + "DCR\"))

        Next
        'End Bring Files Local
        '



        Me.ListBox1.Items.Insert(0, "Starting")

        Dim di As New IO.DirectoryInfo(tempPath + "DCR\")
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xlsm")
        Dim fi As IO.FileInfo
        'Me.ListBox1.Items.Insert(0, "Here0")
        strDate = Now().ToString("ddMMMyyyy")
        Dim strDate2 As String = Now().ToString("yyyy_MM_dd_hhmm")
        Dim vfolder As String = "\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR\Processed\" + strDate2
        '  Dim sqlCon = New SqlConnection(strConn)

        'Me.ListBox1.Items.Insert(0, "Here")
        If aryFi.Length > 0 Then
            If (Not System.IO.Directory.Exists(vfolder)) Then
                System.IO.Directory.CreateDirectory(vfolder)
            End If

        End If
        'Me.ListBox1.Items.Insert(0, "Here2")

        For Each fi In aryFi ' This loops through the Files in the target folder with .xlsm extension.


            Console.WriteLine("File Name: {0}", fi.Name)
            Console.WriteLine("File Full Name: {0}", fi.FullName)
            Me.ListBox1.Items.Insert(0, "Creating XML")

            Dim writer As New XmlTextWriter("\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR\XML\" + Replace(fi.Name, ".xlsm", "") + "-" + strDate + ".xml", System.Text.Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            vFile = fi.FullName
            oBook = oApp.Workbooks.Open(vFile)

            ' Me.ListBox1.Items.Insert(0, "Pausing 15s")
            ' Thread.Sleep(15000)

            Me.ListBox1.Items.Insert(0, vFile)
            '   Label1.Text = vFile
            Console.WriteLine(vFile)
            vSheetCount = oBook.Sheets.Count
            Me.ListBox1.Items.Insert(0, "Total Sheets: " + Trim(vSheetCount))
            'Starting the XML File
            If Not IsNothing(oApp.Sheets("Lists").cells(4, 13).Value) Then
                Console.WriteLine("Good")
                Me.ListBox1.Items.Insert(0, "ID is defined!")

                writer.WriteStartDocument()
                writer.WriteStartElement("DATAROOT")
                For ci As Integer = 4 To vSheetCount
                    oSheets = oApp.Sheets(ci)
                    If oSheets.Range("A1").Value <> "DAILY CONSTRUCTION REPORT" Then
                        Continue For
                    End If

                    ci2 = oSheets.Cells.Find(What:="DAILY OPERATIONS SUMMARY (Hrs)”).Row + 2
                    par1 = oSheets.Range("I" + Trim(Str(ci2))).Value + oSheets.Range("O" + Trim(Str(ci2))).Value + oSheets.Range("T" + Trim(Str(ci2))).Value + oSheets.Range("Y" + Trim(Str(ci2))).Value + oSheets.Range("AD" + Trim(Str(ci2))).Value
                    TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                    par2 = oSheets.Range("Y" + Trim(Str(TDC))).Value
                    'This will check SQl for matching records and skip
                    If GlobalVariables.OnlyNew = True Then

                        If CheckDailyCost_Hours(oApp.Sheets("Lists").cells(4, 13).Value.ToString(), oSheets.Range("AE102").Value, par1, par2) = True Then
                            Continue For
                        End If
                    End If

                    Me.ListBox1.Items.Insert(0, "Sheet: " + Trim(ci))
                    writer.WriteStartElement("DRT")
                    writer.WriteStartElement("Project")
                    If oSheets.Range("AB102").Value = "Report Date:" Then
                        '

                        'writer.WriteElementString("DCRLink", vfolder + "\" + fi.Name)
                        writer.WriteElementString("DCRLink", strDate2 + "/" + Replace(fi.Name, " ", "%20"))
                        writer.WriteElementString("ID", oApp.Sheets("Lists").cells(4, 13).Value.ToString())
                        writer.WriteElementString("Facility", oSheets.Range("N102").Value)
                        writer.WriteElementString("Lease", oSheets.Range("S102").Value)
                        writer.WriteElementString("AFE", oSheets.Range("X102").Value)
                        writer.WriteElementString("Date", oSheets.Range("AE102").Value)
                        writer.WriteElementString("ReportNumber", oSheets.Range("AF103").Value)
                        writer.WriteElementString("ConstructionEngineer", oSheets.Range("G104").Value)
                        writer.WriteElementString("ConstructionSupervisor", oSheets.Range("T104").Value)
                        writer.WriteElementString("WorkingHours", oSheets.Range("X103").Value)
                        writer.WriteElementString("ProjectCumCost", oSheets.Range("AE104").Value)
                        writer.WriteElementString("Type", Replace(oSheets.Range("P103").Value, "&", ""))
                        ci2 = oSheets.Cells.Find(What:="Man Hour Tracker”).Row
                        writer.WriteElementString("Budget", oSheets.Range("G" + Trim(Str(ci2))).Value)
                        writer.WriteElementString("Remaining", oSheets.Range("AD" + Trim(Str(ci2))).Value)
                        writer.WriteElementString("ProjectDescription", Replace(Replace(oSheets.Range("F105").Value, "R&M", "RM"), "&", "and"))
                        If (oSheets.Range("A106").Value = "Safety Issues:") Then
                            writer.WriteElementString("SafetyIssues", Replace(oSheets.Range("F106").Value, "&", "and"))
                            writer.WriteElementString("CurrentOperations", Replace(oSheets.Range("F107").Value, "&", "and"))
                            writer.WriteElementString("Forecast", Replace(oSheets.Range("F108").Value, "&", "and"))
                        Else
                            writer.WriteElementString("CurrentOperations", Replace(oSheets.Range("F106").Value, "&", "and"))
                            writer.WriteElementString("Forecast", Replace(oSheets.Range("F107").Value, "&", "and"))

                        End If
                        ci2 = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                            writer.WriteElementString("DailyTicketCost", oSheets.Range("Y" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("CumTicketCost", oSheets.Range("AD" + Trim(Str(ci2))).Value)
                            writer.WriteEndElement()
                            writer.WriteStartElement("Summary")
                            writer.WriteStartElement("Daily")
                            ci2 = oSheets.Cells.Find(What:="DAILY OPERATIONS SUMMARY (Hrs)”).Row + 2
                            writer.WriteElementString("StandardOperations", oSheets.Range("I" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("WeatherDowntime", oSheets.Range("O" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("ExtraWork", oSheets.Range("T" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("OtherDowntime", oSheets.Range("Y" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("SWP", oSheets.Range("AD" + Trim(Str(ci2))).Value)
                            ' writer.WriteElementString("", oSheets.Range("" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("TotalHours", oSheets.Range("I" + Trim(Str(ci2))).Value + oSheets.Range("O" + Trim(Str(ci2))).Value + oSheets.Range("T" + Trim(Str(ci2))).Value + oSheets.Range("Y" + Trim(Str(ci2))).Value + oSheets.Range("AD" + Trim(Str(ci2))).Value)
                            writer.WriteEndElement()
                            writer.WriteStartElement("Cumulative")
                            ci2 = oSheets.Cells.Find(What:="DAILY OPERATIONS SUMMARY (Hrs)”).Row + 3
                            writer.WriteElementString("StandardOperations", oSheets.Range("I" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("WeatherDowntime", oSheets.Range("O" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("ExtraWork", oSheets.Range("T" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("OtherDowntime", oSheets.Range("Y" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("SWP", oSheets.Range("AD" + Trim(Str(ci2))).Value)
                            writer.WriteElementString("TotalHours", oSheets.Range("I" + Trim(Str(ci2))).Value + oSheets.Range("O" + Trim(Str(ci2))).Value + oSheets.Range("T" + Trim(Str(ci2))).Value + oSheets.Range("Y" + Trim(Str(ci2))).Value + oSheets.Range("AD" + Trim(Str(ci2))).Value)
                            writer.WriteEndElement() ' End Cumulative
                            writer.WriteEndElement() ' End Summary

                            ' SP Declarations
                            'sqlComm.Parameters.AddWithValue("AIMID", )
                            'sqlComm.Parameters.AddWithValue("Facility", )
                            'sqlComm.Parameters.AddWithValue("Area", )
                            'sqlComm.Parameters.AddWithValue("PM", )
                            'sqlComm.Parameters.AddWithValue("AFE", )
                            'sqlComm.Parameters.AddWithValue("BudgetType", )
                            'sqlComm.Parameters.AddWithValue("BudgetCategory", )
                            'sqlComm.Parameters.AddWithValue("AIMProjectName", )
                            'sqlComm.Parameters.AddWithValue("ReportDate", )
                            'sqlComm.Parameters.AddWithValue("Lease", )
                            'sqlComm.Parameters.AddWithValue("ReportNumber", )
                            'sqlComm.Parameters.AddWithValue("ConstrEngr", )
                            'sqlComm.Parameters.AddWithValue("ConstrSuper", )
                            'sqlComm.Parameters.AddWithValue("WorkHourType", )
                            'sqlComm.Parameters.AddWithValue("DCRType", )
                            'sqlComm.Parameters.AddWithValue("Budget", )
                            'sqlComm.Parameters.AddWithValue("DailyCost", )
                            'sqlComm.Parameters.AddWithValue("CumCost", )
                            'sqlComm.Parameters.AddWithValue("Project Description", )
                            'sqlComm.Parameters.AddWithValue("CurrentOperations", )
                            'sqlComm.Parameters.AddWithValue("Forecast", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursStandardOperations", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursWeatherDowntime", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursExtraWork", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursOtherDowntime", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursSWP", )
                            'sqlComm.Parameters.AddWithValue("DailyHoursTotal", )
                            'sqlComm.Parameters.AddWithValue("CumHoursStandardOperations", )
                            'sqlComm.Parameters.AddWithValue("CumHoursWeatherDowntime", )
                            'sqlComm.Parameters.AddWithValue("CumHoursExtraWork", )
                            'sqlComm.Parameters.AddWithValue("CumHoursOtherDowntime", )
                            'sqlComm.Parameters.AddWithValue("CumHoursSWP", )
                            'sqlComm.Parameters.AddWithValue("CumHoursTotal", )
                            'sqlComm.Parameters.AddWithValue("RemainingBudget", )
                            'sqlComm.Parameters.AddWithValue("DCRLink", )


                            ' sqlComm.Parameters.AddWithValue("", )



                            'Daily Operational Log
                            DOPLOG = oSheets.Cells.Find(What:="DAILY OPERATIONS LOG - (00:00 - 24:00)”).Row
                            DOPLOG_SUM = oSheets.Cells.Find(What:="DAILY OPERATIONS SUMMARY (Hrs)”).Row
                            writer.WriteStartElement("DailyOperationsLog")
                            '
                            Me.ListBox1.Items.Add("Starting Daily Log")
                            For ci_dol = DOPLOG + 1 To DOPLOG_SUM - 1
                                If Not IsNothing(oSheets.Range("I" + Trim(Str(ci_dol))).Value) Then
                                    writer.WriteStartElement("LogEntry")
                                    writer.WriteElementString("TimeRow", Trim(Str(ci_dol)))
                                    writer.WriteElementString("TimeFrom", TimeSpan.FromHours(oSheets.Range("A" + Trim(Str(ci_dol))).Value * 24).ToString)
                                    writer.WriteElementString("TimeTo", TimeSpan.FromHours(oSheets.Range("C" + Trim(Str(ci_dol))).Value * 24).ToString)
                                    writer.WriteElementString("TimeHours", oSheets.Range("E" + Trim(Str(ci_dol))).Value)
                                    writer.WriteElementString("TimeCode", oSheets.Range("G" + Trim(Str(ci_dol))).Value)
                                    writer.WriteElementString("TimeDescription", Replace(oSheets.Range("I" + Trim(Str(ci_dol))).Value, "&", "and"))
                                    'writer.WriteElementString("", oSheets.Range("" + Trim(Str(ci2))).Value)
                                    writer.WriteEndElement() 'End LogEntry

                                End If
                            Next
                            writer.WriteEndElement() 'DOL




                            Select Case oSheets.Range("P103").Value
                            Case Is = "ABN"
                                'Vendor and Hours
                                AOLV = oSheets.Cells.Find(What:="AOL Platform Based Vendor”).Row
                                    ABV = oSheets.Cells.Find(What:="Autonomous Based Vendor”).Row
                                    AB = oSheets.Cells.Find(What:="Autonomous Subtotal:”).Row
                                    'AOL PF Based Left Column
                                    For ci_cap = AOLV + 1 To ABV - 2
                                        If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                            writer.WriteEndElement() 'End Entry

                                        End If
                                    Next
                                    'AOL PF Based Right Column
                                    For ci_cap = AOLV + 1 To ABV - 2
                                        If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                            writer.WriteEndElement() 'End Entry

                                        End If
                                    Next
                                    'Autonomous Based
                                    For ci_cap = ABV + 1 To AB - 1
                                        If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                            writer.WriteEndElement() 'End Entry

                                        End If
                                    Next
                                    'AOL PF Based Right Column
                                    For ci_cap = ABV + 1 To AB - 1
                                        If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                            writer.WriteEndElement() 'End Entry

                                        End If
                                    Next



                                    TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                                    'Offshore Cum
                                    writer.WriteElementString("OffshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                    'Onshore Cum
                                    writer.WriteElementString("OnshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                    'Daily Totals
                                    writer.WriteElementString("TotalDailyCost", oSheets.Range("Y" + Trim(Str(TDC))).Value)
                                    writer.WriteElementString("TotalCumCost", oSheets.Range("AD" + Trim(Str(TDC))).Value)
                                Case Is = "Expense (AFE)"
                                'Vendor and Hours
                                AOLV = oSheets.Cells.Find(What:="AOL Platform Based Vendor”).Row
                                ABV = oSheets.Cells.Find(What:="Autonomous Based Vendor”).Row
                                AB = oSheets.Cells.Find(What:="Autonomous Subtotal:”).Row
                                'AOL PF Based Left Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'Autonomous Based
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next



                                TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                                'Offshore Cum
                                writer.WriteElementString("OffshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Onshore Cum
                                writer.WriteElementString("OnshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Daily Totals
                                writer.WriteElementString("TotalDailyCost", oSheets.Range("Y" + Trim(Str(TDC))).Value)
                                writer.WriteElementString("TotalCumCost", oSheets.Range("AD" + Trim(Str(TDC))).Value)
                            Case Is = "Capital (AFE)"
                                'Vendor and Hours
                                AOLV = oSheets.Cells.Find(What:="AOL Platform Based Vendor”).Row
                                ABV = oSheets.Cells.Find(What:="Autonomous Based Vendor”).Row
                                AB = oSheets.Cells.Find(What:="Autonomous Subtotal:”).Row
                                'AOL PF Based Left Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'Autonomous Based
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next



                                TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                                'Offshore Cum
                                writer.WriteElementString("OffshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Onshore Cum
                                writer.WriteElementString("OnshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Daily Totals
                                writer.WriteElementString("TotalDailyCost", oSheets.Range("Y" + Trim(Str(TDC))).Value)
                                writer.WriteElementString("TotalCumCost", oSheets.Range("AD" + Trim(Str(TDC))).Value)
                            Case Is = "Expense (AFE)"
                                'Vendor and Hours
                                AOLV = oSheets.Cells.Find(What:="AOL Platform Based Vendor”).Row
                                ABV = oSheets.Cells.Find(What:="Autonomous Based Vendor”).Row
                                AB = oSheets.Cells.Find(What:="Autonomous Subtotal:”).Row
                                'AOL PF Based Left Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = AOLV + 1 To ABV - 2
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'Autonomous Based
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("A" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("A" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("L" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("N" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("P" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next
                                'AOL PF Based Right Column
                                For ci_cap = ABV + 1 To AB - 1
                                    If Not IsNothing(oSheets.Range("R" + Trim(Str(ci_cap))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("Vendor", oSheets.Range("R" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("AC" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("AE" + Trim(Str(ci_cap))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("AG" + Trim(Str(ci_cap))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next



                                TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                                'Offshore Cum
                                writer.WriteElementString("OffshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Onshore Cum
                                writer.WriteElementString("OnshoreCumCost", oSheets.Range("Y" + Trim(Str(TDC - 2))).Value)
                                'Daily Totals
                                writer.WriteElementString("TotalDailyCost", oSheets.Range("Y" + Trim(Str(TDC))).Value)
                                writer.WriteElementString("TotalCumCost", oSheets.Range("AD" + Trim(Str(TDC))).Value)
                            Case Is = "P & A (AFE)"
                                '   Me.ListBox1.Items.Insert(0, "Setting up an R&M DCR")
                                RMDRT = oSheets.Cells.Find(What:="R&M DAY RATE COST TRACKER & POB”).Row
                                Try
                                    RMDRT_CUM = oSheets.Cells.Find(What:="ABN Offshore / Cumulative”).Row
                                Catch ex As Exception
                                    RMDRT_CUM = oSheets.Cells.Find(What:="R&M Offshore / Cumulative”).Row
                                End Try
                                PDRT_CUM = oSheets.Cells.Find(What:="PAINT Offshore / Cumulative”).Row


                                For ci_RM = RMDRT + 3 To RMDRT_CUM - 1
                                    If Not IsNothing(oSheets.Range("D" + Trim(Str(ci_RM))).Value) Then
                                        writer.WriteStartElement("Entry")
                                        writer.WriteElementString("CostCode", oSheets.Range("A" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("Vendor", oSheets.Range("D" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("POB", oSheets.Range("Q" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("DailyHours", oSheets.Range("S" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("CumHours", oSheets.Range("V" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("DailyCost", oSheets.Range("Y" + Trim(Str(ci_RM))).Value)
                                        writer.WriteElementString("CumCost", oSheets.Range("AD" + Trim(Str(ci_RM))).Value)
                                        'writer.WriteElementString("", oSheets.Range("" + Trim(Str(ci2))).Value)
                                        writer.WriteEndElement() 'End Entry

                                    End If
                                Next

                            Case Is = "R&M"
                                    '   Me.ListBox1.Items.Insert(0, "Setting up an R&M DCR")
                                    RMDRT = oSheets.Cells.Find(What:="R&M DAY RATE COST TRACKER & POB”).Row
                                    RMDRT_CUM = oSheets.Cells.Find(What:="R&M Offshore / Cumulative”).Row
                                    PDRT_CUM = oSheets.Cells.Find(What:="PAINT Offshore / Cumulative”).Row


                                    For ci_RM = RMDRT + 3 To RMDRT_CUM - 1
                                        If Not IsNothing(oSheets.Range("D" + Trim(Str(ci_RM))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("CostCode", oSheets.Range("A" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("Vendor", oSheets.Range("D" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("Q" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("S" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("V" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("DailyCost", oSheets.Range("Y" + Trim(Str(ci_RM))).Value)
                                            writer.WriteElementString("CumCost", oSheets.Range("AD" + Trim(Str(ci_RM))).Value)
                                            'writer.WriteElementString("", oSheets.Range("" + Trim(Str(ci2))).Value)
                                            writer.WriteEndElement() 'End Entry

                                        End If
                                    Next

                                Case Is = "Paint"
                                    ' Me.ListBox1.Items.Insert(0, "Setting up an Paint DCR")
                                    PDRT = oSheets.Cells.Find(What:="PAINT DAY RATE COST TRACKER & POB").Row
                                    PDRT_CUM = oSheets.Cells.Find(What:="PAINT Offshore / Cumulative”).Row

                                    For ci_paint = PDRT + 3 To PDRT_CUM - 1
                                        If Not IsNothing(oSheets.Range("D" + Trim(Str(ci_paint))).Value) Then
                                            writer.WriteStartElement("Entry")
                                            writer.WriteElementString("CostCode", oSheets.Range("A" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("Vendor", oSheets.Range("D" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("POB", oSheets.Range("Q" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("DailyHours", oSheets.Range("S" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("CumHours", oSheets.Range("V" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("DailyCost", oSheets.Range("Y" + Trim(Str(ci_paint))).Value)
                                            writer.WriteElementString("CumCost", oSheets.Range("AD" + Trim(Str(ci_paint))).Value)
                                            writer.WriteEndElement() 'End Entry
                                        End If
                                    Next 'Write Paint Totals
                                    writer.WriteElementString("TotalDailyCost", oSheets.Range("Y" + Trim(Str(PDRT_CUM))).Value)
                                    writer.WriteElementString("TotalCumCost", oSheets.Range("AD" + Trim(Str(PDRT_CUM))).Value)
                                    ' Write Paint Metrics
                                    writer.WriteStartElement("Metrics")
                                    writer.WriteStartElement("Daily")
                                    writer.WriteElementString("Progress", oSheets.Range("Q" + Trim(Str(PDRT_CUM + 2))).Value)
                                    writer.WriteElementString("SurfacePrep", oSheets.Range("Q" + Trim(Str(PDRT_CUM + 3))).Value)
                                    writer.WriteElementString("IntermediateCoat", oSheets.Range("Q" + Trim(Str(PDRT_CUM + 4))).Value)
                                    writer.WriteElementString("TopCoat", oSheets.Range("Q" + Trim(Str(PDRT_CUM + 5))).Value)
                                    writer.WriteElementString("Rigup", oSheets.Range("Q" + Trim(Str(PDRT_CUM + 6))).Value)

                                    writer.WriteEndElement() 'End Daily Metrics
                                    writer.WriteStartElement("Cumulative")
                                    writer.WriteElementString("Progress", oSheets.Range("S" + Trim(Str(PDRT_CUM + 2))).Value)
                                    writer.WriteElementString("SurfacePrep", oSheets.Range("S" + Trim(Str(PDRT_CUM + 3))).Value)
                                    writer.WriteElementString("IntermediateCoat", oSheets.Range("S" + Trim(Str(PDRT_CUM + 4))).Value)
                                    writer.WriteElementString("TopCoat", oSheets.Range("S" + Trim(Str(PDRT_CUM + 5))).Value)
                                    writer.WriteElementString("Rigup", oSheets.Range("S" + Trim(Str(PDRT_CUM + 6))).Value)

                                    writer.WriteEndElement() 'End Cumulative Metrics
                                    writer.WriteEndElement() 'End Metrics

                            End Select

                            writer.WriteEndElement() '"</DRT>" 


                            'writer.WriteElementString("", oSheets.Range("").Value)
                            'MH = oSheets.Cells.Find(What:="Man Hour Tracker”).Row
                            'PB = oSheets.Cells.Find(What:="Platform Based Subtotal:”).Row
                            'AB = oSheets.Cells.Find(What:="Autonomous Subtotal:”).Row
                            'LST = oSheets.Cells.Find(What:="LUMP SUM COST TRACKER”).Row
                            'DRT = oSheets.Cells.Find(What:="DAY RATE COST TRACKER”).Row
                            'DR1 = oSheets.Cells.Find(What:="Consulting Services CCT”).Row
                            'DR2 = oSheets.Cells.Find(What:="Project Management and Visual Inspection CCT”).Row
                            'DR3 = oSheets.Cells.Find(What:="Non Destructive Testing CCT”).Row
                            'DR4 = oSheets.Cells.Find(What:="Equipment Rental CCT”).Row
                            'DR5 = oSheets.Cells.Find(What:="Fuel, Water, Power, and Lubricant CCT”).Row
                            'DR6 = oSheets.Cells.Find(What:="Hydro-test Equipment CCT”).Row
                            'DR7 = oSheets.Cells.Find(What:="Offshore Electrical CCT”).Row
                            'DR8 = oSheets.Cells.Find(What:="Offshore Hookup CCT”).Row
                            'DR9 = oSheets.Cells.Find(What:="Offshore Instrumentation CCT”).Row
                            'DR10 = oSheets.Cells.Find(What:="Transportation - Marine CCT”).Row
                            'DR11 = oSheets.Cells.Find(What:="Platform Installation CCT”).Row
                            'DR12 = oSheets.Cells.Find(What:="Pipeline Installation CCT”).Row
                            'DR13 = oSheets.Cells.Find(What:="Pipeline Tie-in CCT”).Row
                            'DR14 = oSheets.Cells.Find(What:="Miscellaneous CCT”).Row
                            'TDC = oSheets.Cells.Find(What:="Total Daily / Cumulative”).Row
                            'HAZ = oSheets.Cells.Find(What:="Hazards (Y/N)”).Row
                            'RMDRT = oSheets.Cells.Find(What:="R&M DAY RATE COST TRACKER & POB”).Row
                            'PDRT = oSheets.Cells.Find(What:="PAINT DAY RATE COST TRACKER & POB”).Row
                            'PDRT_CUM = oSheets.Cells.Find(What:="PAINT Offshore / Cumulative”).Row
                            'Console.WriteLine(oSheets.Range("AE102").Value)
                            ' writer.WriteStartElement("<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>")

                        End If
                    '   writer.WriteEndElement()
                Next

                oBook.Close(False)
                writer.WriteEndElement()
                vFile2 = Replace(vFile, "&", "and")
                vFile2 = Replace(vFile2, "#", "")
                IO.File.Move(vFile, Replace(vFile2, tempPath + "DCR", "\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR\processed\" + strDate2))
                writer.WriteEndDocument()
                writer.Close()
            Else
                oBook.Close(False)
                vFile2 = Replace(vFile, "&", "and")
                vFile2 = Replace(vFile2, "#", "")
                Me.ListBox1.Items.Insert(0, "Missing AIMID: " + Trim(vFile))
                IO.File.Move(vFile, Replace(vFile2, tempPath + "DCR\", "\\arenaenergy0.sharepoint.com@SSL\DavWWWRoot\sites\AIM\DCR\Error\" + strDate2 + "_"))
            End If
            '    sqlCon.Open()
            '    sqlComm.ExecuteNonQuery()
            'End Using
        Next

        If GlobalVariables.Runtype = False Then
            Me.Close()
            Application.Exit()
        End If
    End Sub
    Private Sub End_Excel_App() '(datestart As Date, dateEnd As Date)
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            'If Process.StartTime >= datestart And Process.StartTime <= dateEnd Then
            Process.Kill()
            '   Exit For
            'End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Label1.Text = Now()
        Call scheduler()
    End Sub
    Function CheckDailyCost_Hours(pAIMID As Integer, pReportDate As Date, pDailyHours As Decimal, pDailyCost As Decimal)
        Dim connectionString As String = "Data source=HUSV-AOL-SQL16\AIM;Database=AIM; integrated security=True"
        Dim vSqlDailyCost, vSqlDailyHours As Decimal
        Dim cn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand
        cmd.Connection = cn
        cn.Open()
        Dim cmd2 As SqlCommand = New SqlCommand("spGetDCRTotals", cn)
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.Parameters.Add("@AIMID", SqlDbType.Int).Value = pAIMID
        cmd2.Parameters.Add("@ReportDate", SqlDbType.Date).Value = pReportDate
        cmd2.Parameters.Add("@vDailyCost", SqlDbType.Decimal).Direction = ParameterDirection.Output
        cmd2.Parameters.Add("@vDailyHoursTotal", SqlDbType.Decimal).Direction = ParameterDirection.Output
        cmd2.ExecuteNonQuery()
        cn.Close()
        If IsDBNull(cmd2.Parameters("@vDailyCost").Value) Or IsDBNull(cmd2.Parameters("@vDailyHoursTotal").Value) Or pReportDate = Today() Then
            Return False
        ElseIf Not IsDBNull(cmd2.Parameters("@vDailyCost").Value) And Not IsDBNull(cmd2.Parameters("@vDailyHoursTotal").Value) Then
            vSqlDailyCost = cmd2.Parameters("@vDailyCost").Value
            vSqlDailyHours = cmd2.Parameters("@vDailyHoursTotal").Value
            If Math.Round(vSqlDailyCost) = Math.Round(pDailyCost) And Math.Round(vSqlDailyHours) = Math.Round(pDailyHours) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = False Then
            GlobalVariables.Runtype = False

        Else
            GlobalVariables.Runtype = True
            ' Call scheduler()
            Debug.WriteLine(GlobalVariables.Runtype)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox1.Checked = False Then
            GlobalVariables.OnlyNew = False

        Else
            GlobalVariables.OnlyNew = True
            ' Call scheduler()
            Debug.WriteLine(GlobalVariables.Runtype)
        End If
    End Sub
End Class
Public Class GlobalVariables
    Public Shared Runtype As Boolean = False
    Public Shared OnlyNew As Boolean = True
End Class
