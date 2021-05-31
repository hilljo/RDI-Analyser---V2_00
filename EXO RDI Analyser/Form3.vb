
Imports System.Data.DataTable

'Bullion
'COM02:COM26 -> 2502
'COM04:COM26 -> 2504   
'COM04:COM23 -> 2632  
'COM06:COM26 -> 2503  
'COM08:COM26 -> 2501  
'COM08:COM23 -> 2634  


Public Class FormPABX

    Sub runEXOPSTN_GSM_Mismatch()
        Dim PABXExcel As Object = ""
        Dim PABXBook As Object = ""
        Dim PABXSheet As Object = ""
        Dim Dateinfile As String = ""
        Dim RowofPABXExt As Integer = 2

        Dim howmanybadmatches As Integer = 0
        Dim PSTNtoGSMSteppscount As Integer = 0
        Dim PSTNtoGSMDundeecount As Integer = 0
        Dim GSMtoPSTNSteppscount As Integer = 0
        Dim GSMtoPSTNDundeecount As Integer = 0

        Dim PABXExtCheck As String = ""
        Dim PABXDataCenter As String = ""
        Dim PABXTime As String = ""
        Dim DialingNumber As String = ""

        Dim PABXExt As String = ""

        Dim PABXrowcounter As Integer = 0
        Dim PABXlastusedrow As Integer = 0

        Dim firstchar As Char


        PABXExcel = CreateObject("Excel.Application")

        Dateinfile = DateTimePickerSplitterPABX.Value.Date.ToString("dd-MMM-yy")
        Dim networkpathlength As String = SelectedFolderPABX.Text.Length ' strip off \PABX
        Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, networkpathlength)

        ' Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, 20)
        Try
            PABXBook = PABXExcel.Workbooks.open(partialpath & "\Telemetry_Stats_" & Dateinfile & ".xls")
        Catch e As Exception
            MessageBox.Show("No File Found")
            Exit Sub
        End Try
        PABXSheet = PABXBook.Worksheets("crlx29")
        partialpath = SelectedFolderPABX.Text.Substring(0, 15)
        PABXlastusedrow = PABXSheet.UsedRange.Rows.Count - 1



        For PABXrowcounter = 2 To PABXlastusedrow ' 65536 ' PABXlastrow 'just limied for test
            ProgressBarreadPABXLogs.Maximum = PABXlastusedrow ' / 100
            ProgressBarreadPABXLogs.Value = PABXrowcounter

            ProgressLabel.Text = ProgressBarreadPABXLogs.Value & " Calls Analysed ...   " & String.Format("{0:n2}", PABXrowcounter / PABXlastusedrow) * 100 & " % Complete"
            Try

                PABXDataCenter = PABXSheet.cells(PABXrowcounter, 15).text

                If PABXDataCenter = "SW- Dundee Telemetry" Then
                    PABXExt = PABXSheet.cells(PABXrowcounter, 2).text '
                    If PABXExt = "2501" Or PABXExt = "2502" Or PABXExt = "2503" Or PABXExt = "2504" Or PABXExt = "2632" Or PABXExt = "2634" Then
                        'check gsms dialing pstn ports
                        DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text
                        If DialingNumber <> "" Then
                            PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
                            firstchar = DialingNumber(0)
                            If firstchar = "7" Then
                                howmanybadmatches = howmanybadmatches + 1
                                GSMtoPSTNDundeecount = GSMtoPSTNDundeecount + 1
                                PABXSheet.cells(howmanybadmatches + 1, 20).Value = "GSM dialed " & PABXDataCenter & " PSTN port " & PABXExt & " From " & DialingNumber & " at " & PABXTime
                                Listboxcli.Items.Add(DialingNumber) '= DialingNumber


                                '   MessageBox.Show("GSM on PSTN at row ", PABXrowcounter.ToString)
                            End If
                        End If
                    End If

                    'check PSTNs dialing GSM ports
                    If PABXExt = "2601" Or PABXExt = "2602" Or PABXExt = "2603" Or PABXExt = "2604" Then

                        DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text
                        If DialingNumber <> "" Then
                            PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
                            firstchar = DialingNumber(0)
                            If firstchar <> "7" Then
                                howmanybadmatches = howmanybadmatches + 1
                                PSTNtoGSMDundeecount = PSTNtoGSMDundeecount + 1
                                PABXSheet.cells(howmanybadmatches + 1, 20).Value = "PSTN dialed " & PABXDataCenter & " GSM port " & PABXExt & " From " & DialingNumber & " at " & PABXTime
                                Listboxcli.Items.Add(DialingNumber) '= DialingNumber

                            End If
                        End If
                    End If

                End If

                If PABXDataCenter = "SW- Bridge Telemetry" Then
                    PABXExt = PABXSheet.cells(PABXrowcounter, 2).text
                    If PABXExt = "2501" Or PABXExt = "2502" Or PABXExt = "2503" Or PABXExt = "2504" Then ' Or PABXExt = "2632" Or PABXExt = "2634" Then

                        DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text

                        If DialingNumber <> "" Then
                            PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
                            firstchar = DialingNumber(0)
                            If firstchar = "7" Then
                                howmanybadmatches = howmanybadmatches + 1
                                GSMtoPSTNSteppscount = GSMtoPSTNSteppscount + 1
                                PABXSheet.cells(howmanybadmatches + 1, 20).Value = "GSM dialed " & PABXDataCenter & " PSTN port " & PABXExt & " From " & DialingNumber & " at " & PABXTime
                                Listboxcli.Items.Add(DialingNumber) '= DialingNumber

                                '   MessageBox.Show("GSM on PSTN at row ", PABXrowcounter.ToString)

                            End If
                        End If
                    End If


                    'check PSTNs dialing GSM ports
                    If PABXExt = "2601" Or PABXExt = "2602" Then 'din't know what EXT 1modem work will allaocate to stepps Or PABXExt = "2603" Or PABXExt = "2604" Then

                        DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text
                        If DialingNumber <> "" Then
                            PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
                            firstchar = DialingNumber(0)
                            If firstchar <> "7" Then
                                PSTNtoGSMSteppscount = PSTNtoGSMSteppscount + 1
                                howmanybadmatches = howmanybadmatches + 1
                                Listboxcli.Items.Add(DialingNumber) '= DialingNumber

                                PABXSheet.cells(howmanybadmatches + 1, 20).Value = "PSTN dialed " & PABXDataCenter & " GSM port " & PABXExt & " From " & DialingNumber & " at " & PABXTime

                            End If
                        End If
                    End If

                End If

            Catch e As System.NullReferenceException
            Catch e As FormatException
            Catch e As IndexOutOfRangeException
            End Try

            OKcount.Text = ProgressBarreadPABXLogs.Value
            PSTNtoGSMStepps.Text = PSTNtoGSMSteppscount
            PSTNtoGSMDundee.Text = PSTNtoGSMDundeecount
            GSMtoPSTNStepps.Text = GSMtoPSTNSteppscount
            GSMtoPSTNDundee.Text = GSMtoPSTNDundeecount

            FoundCount.Text = PSTNtoGSMSteppscount + PSTNtoGSMDundeecount + GSMtoPSTNSteppscount + GSMtoPSTNDundeecount ' howmanybadmatches

        Next PABXrowcounter
        If howmanybadmatches = 0 Then
            PABXSheet.cells(howmanybadmatches + 1, 20).Value = "No Calls to wrong technology"

        End If
        ' save and exit
        PABXBook.Save()

        PABXSheet = Nothing
        PABXBook = Nothing
        PABXExcel.Quit()
        PABXExcel = Nothing
        GC.Collect()
    End Sub


    'Private Sub collolationINRDIandPABX()
    '    Dim PABXExcel As Object
    '    Dim RDIExcel As Object
    '    Dim PABXBook As Object
    '    Dim PABXSheet As Object
    '    Dim RDIBook As Object
    '    Dim RDISheet As Object

    '    Dim Dateinfile As String
    '    Dim RowofPABXExt As Integer = 2
    '    Dim howmanymatches As Integer = 0
    '    '      Dim DialingNumber As String
    '    '   Dim rowcounter As Integer
    '    Dim PABXExtCheck As String = ""
    '    Dim PABXTime As String
    '    Dim DialingNumber As String
    '    Dim PABXDuration As String
    '    Dim PABXExt As String = ""

    '    Dim RDITime As String = ""
    '    Dim RDICommsPort As String
    '    Dim PABXDataCenter As String = ""
    '    Dim rditimeoffsetmax As String = ""
    '    Dim rditimeoffsetmin As String = ""
    '    '  Dim rditimeoffset1 As String = ""
    '    '  Dim rditimeoffset0 As String = ""
    '    Dim mincallduration As Date = "00:00:30"
    '    '     Dim calldurationinTime As Date
    '    '  Dim PABXTime As String = DateTimePickerSplitter.Value.time.ToString("hh:mm:ss")

    '    RDIExcel = CreateObject("Excel.Application")
    '    PABXExcel = CreateObject("Excel.Application")

    '    Dateinfile = DateTimePickerSplitterPABX.Value.Date.ToString("dd-MMM-yyyy")
    '    Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, 20)

    '    '    PABXBook = PABXExcel.Workbooks.open(partialpath & "\Telemetry_Stats-" & Dateinfile & ".xls")
    '    PABXBook = PABXExcel.Workbooks.open(partialpath & "\Telemetry_Stats_02-Sep-15 test.xls")
    '    PABXSheet = PABXBook.Worksheets("crlx29")

    '    partialpath = Form1.SelectedFolder.Text.Substring(0, 15)
    '    ' RDIBook = RDIExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")

    '    RDIBook = RDIExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI 02-Sep-2015 - test.xls")

    '    RDISheet = RDIBook.Worksheets("EXO")


    '    For RDIrowcounter = 7 To 17 '3 To 10
    '        Try 'just incase it's shite
    '            RDITime = RDISheet.cells(RDIrowcounter, 19).text
    '            'the BTS logs only have to nearest minute so as we don't know how far the clocks are out we will have to try a min offset or either side 
    '            ' well perhaps the pabx is quite slow

    '            'dundee rdi 00:05:53
    '            'pabx 00:03:24 - 204 seconds
    '            '149 seconds 
    '            rditimeoffsetmax = Format(CDate(DateAdd("s", -169, RDITime)), "HH:mm:ss")
    '            rditimeoffsetmin = Format(CDate(DateAdd("s", -129, RDITime)), "HH:mm:ss")
    '            '  rditimeoffset1 = Format(CDate(DateAdd("s", -150, RDITime)), "HH:mm:ss")

    '            'assume correct time
    '            'rditimeoffset0 = Format(CDate(DateAdd("s", -149, RDITime)), "HH:mm:ss")

    '            RDICommsPort = RDISheet.cells(RDIrowcounter, 16).text




    '        Catch e As Exception
    '        End Try
    '        'TEST' limit loop for test
    '        For PABXrowcounter = 2 To 800 'just limied for test
    '            PABXExt = PABXSheet.cells(PABXrowcounter, 2).text
    '            PABXDataCenter = PABXSheet.cells(PABXrowcounter, 15).text
    '            If PABXDataCenter = "SW- Dundee Telemetry" And PABXExt = "2502" Then
    '                PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
    '                PABXDuration = PABXSheet.cells(PABXrowcounter, 5).text
    '                DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text




    '                '  Case Is > TimeValue(rditimeoffsetmin)
    '                If TimeValue(PABXTime) < TimeValue(rditimeoffsetmin) Then
    '                    If TimeValue(PABXTime) > TimeValue(rditimeoffsetmax) Then
    '                        If PABXDuration < mincallduration Then
    '                            howmanymatches = howmanymatches + 1 ' we may have >1 so need to put in seperate cols
    '                            RDISheet.cells(RDIrowcounter, 20 + howmanymatches).Value = "PABXTIME " & PABXTime & " RDITime " & RDITime & " Duration " & PABXDuration & " PABXExt " & PABXExt & " Dialingnumber " & DialingNumber

    '                            '       MessageBox.Show("PABXTIME " & PABXTime & " RDITime " & RDITime & " in range")
    '                            '     InRange = True
    '                        End If

    '                    End If
    '                End If




    '                'Select Case PABXTime
    '                '    Case (rditimeoffset0), (rditimeoffset1), (rditimeoffset2), (rditimeoffset3)

    '                '        '       MessageBox.Show("PABXTIME " & PABXTime & " RDITime " & RDITime & " PABXExt " & PABXExt & " Dialingnumber " & DialingNumber)
    '                '        '          populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 15, "NO CARRIER", CommPortDetails, devicephonenumber, "Dialed at ", timeofrequest)
    '                '        '    RDISheet = RDIExcel.Worksheets("EXO")

    '                '        '        calldurationinTime = "00:" & PABXDuration

    '                '        If PABXDuration < mincallduration Then
    '                '            howmanymatches = howmanymatches + 1 ' we may have >1 so need to put in seperate cols
    '                '            RDISheet.cells(RDIrowcounter, 20 + howmanymatches).Value = "PABXTIME " & PABXTime & " RDITime " & RDITime & " Duration " & PABXDuration & " PABXExt " & PABXExt & " Dialingnumber " & DialingNumber
    '                '        End If

    '                'End Select
    '            End If



    '        Next PABXrowcounter
    '        howmanymatches = 0
    '    Next RDIrowcounter
    '    ' save and exit
    '    RDIBook.Save()

    '    PABXSheet = Nothing
    '    PABXBook = Nothing
    '    RDISheet = Nothing
    '    RDIBook = Nothing
    '    RDIExcel.Quit()
    '    RDIExcel = Nothing
    '    PABXExcel.Quit()
    '    PABXExcel = Nothing
    '    GC.Collect()
    'End Sub



















    Private Sub openexcelsheet()
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object
        Dim Dateinfile As String
        'start new workbook
        On Error GoTo Errornofilefound

        oExcel = CreateObject("Excel.Application")

        Dateinfile = DateTimePickerSplitterPABX.Value.Date.ToString("dd-MMM-yy")
        Dim networkpathlength As String = SelectedFolderPABX.Text.Length ' strip off \PABX
        Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, networkpathlength)

        'Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, 20)
        '   oBook = oExcel.Workbooks.open(partialpath & "\Telemetry_Stats-02-Mar-2015.xls")
        System.Diagnostics.Process.Start(partialpath & "\Telemetry_Stats_" & Dateinfile & ".xls") ' 02-Mar-2015.xls")
        'add data
        ' save and exit
        'oBook.Save()
        'oBook.Application.DisplayAlerts = False

        'Dim Dateinfile As String = DateTimePickerSplitterDL.Value.Date.ToString("dd-MMM-yyyy")
        'oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        'oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        '    oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
        Exit Sub
Errornofilefound:
        MessageBox.Show("Sorry No file or Path Found." & partialpath & "\Telemetry_Stats-" & Dateinfile & ".xls")
        '      PleaseWait.Close()

    End Sub



    Private Function SelectFolderPABX() As String
        Dim PABXdefaultfolder = driveletter + "RDI Analyser\PABX"
        Dim FilesFolder As String

        FolderBrowserDialog1.ShowDialog()
        FilesFolder = FolderBrowserDialog1.SelectedPath
        If FilesFolder = "" Then
            FilesFolder = PABXdefaultfolder
        End If
        SelectedFolderPABX.Text = FilesFolder
        Return FilesFolder
    End Function

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ClearForm()
        SelectedFolderPABX.Text = driveletter + "RDI Analyser\PABX"

    End Sub
    Private Sub btnSelectFolderPABX_Click(sender As Object, e As EventArgs) Handles btnSelectFolderPABX.Click
        Dim Folder As String
        Folder = SelectFolderPABX()
    End Sub

    Private Sub FormPABX_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub GroupBox6_Enter(sender As Object, e As EventArgs)

    End Sub


    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub









    Sub OpenPABXLogButton_Click(sender As Object, e As EventArgs) Handles OpenPABXLogButton.Click
        Dim Dateinfile As String = DateTimePickerSplitterPABX.Value.Date.ToString("dd-MMM-yy")
        Dim networkpathlength As String = SelectedFolderPABX.Text.Length ' strip off \PABX
        Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, networkpathlength)

        '   Dim partialpath As String = SelectedFolderPABX.Text.Substring(0, 20)
        '   oBook = oExcel.Workbooks.open(partialpath & "\Telemetry_Stats-02-Mar-2015.xls")
        System.Diagnostics.Process.Start(partialpath & "\Telemetry_Stats_" & Dateinfile & ".xls") ' 02-Mar-2015.xls")

    End Sub

    Private Sub EXOPSTN_GSM_Mismatch_Click(sender As Object, e As EventArgs) Handles EXOPSTN_GSM_Mismatch.Click
        runEXOPSTN_GSM_Mismatch()
    End Sub



    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub OKcount_TextChanged(sender As Object, e As EventArgs) Handles OKcount.TextChanged

    End Sub

    Private Sub Howmany_Click(sender As Object, e As EventArgs) Handles Howmany.Click

    End Sub

    Private Sub FoundCount_TextChanged(sender As Object, e As EventArgs) Handles FoundCount.TextChanged

    End Sub

    Private Sub SelectedFolderPABX_TextChanged(sender As Object, e As EventArgs) Handles SelectedFolderPABX.TextChanged

    End Sub
End Class