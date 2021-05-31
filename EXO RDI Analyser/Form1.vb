Imports System.IO

Public Module DefineGlobals

    Public driveletter As String = "c:\" 'test build
    'Public driveletter As String = "d:\" 'thick client build
    Public populatexcelsheetdetailsRowCounter As Integer = 3 'where to put the detailed log descripts in the excel sheet
    Public populatexcelsheetdetailsModemprodsRowCounter As Integer = 2 'where to put the detailed modem prods data in the excel sheet

End Module

Public Class Form1
    ' Public Shared populatexcelsheetdetailsRowCounter As Integer = 3 'where to put the detailed log descripts in the excel sheet
    ' Public Shared populatexcelsheetdetailsModemprodsRowCounter As Integer = 2 'where to put the detailed modem prods data in the excel sheet

    Private Property DateoffilestoSplit As String


    Private Sub Split_logs(FiletosplitFullPath As String, WhatToLookFor As String)
        Dim Filetosplit As System.IO.FileInfo
        Dim filename As String
        Dim filepath As String

        Dim FileReader As StreamReader
        Dim strLine As String
        Dim FileWriter As StreamWriter
        Dim FileWriterM As StreamWriter

        Dim Filesplit As String
        Dim FilesplitWithM As String
        Dim ExcludeComPortX As String = ""
        Dim ExcludeComPortY As String = ""
        Dim ExcludeComPortZ As String = ""

        Filetosplit = My.Computer.FileSystem.GetFileInfo(FiletosplitFullPath)

        filename = Filetosplit.Name
        ' MessageBox.Show(filename)
        filepath = Filetosplit.DirectoryName
        'MessageBox.Show(filepath)

        'we now want to include the debug messages but exclude the messages for other com ports
        Select Case WhatToLookFor
            Case "COM23"
                ExcludeComPortX = "COM24"
                ExcludeComPortY = "COM25"
                ExcludeComPortZ = "COM26"

            Case "COM24"
                ExcludeComPortX = "COM23"
                ExcludeComPortY = "COM25"
                ExcludeComPortZ = "COM26"

            Case "COM25"
                ExcludeComPortX = "COM23"
                ExcludeComPortY = "COM24"
                ExcludeComPortZ = "COM26"

            Case "COM26"
                ExcludeComPortX = "COM23"
                ExcludeComPortY = "COM24"
                ExcludeComPortZ = "COM25"

        End Select



        '"d:\Users\jon\Documents\work\EXORDI\COM01\comxx.txt" 

        Filesplit = filepath & "\" & WhatToLookFor & "_" & filename 'make the file name and path based on the unsplit version
        FilesplitWithM = filepath & "\" & WhatToLookFor & "_M_" & filename 'make the file name and path based on the unsplit version

        'MessageBox.Show(Filesplit)

        FileReader = New StreamReader(FiletosplitFullPath) ' set up read
        FileWriter = New StreamWriter(Filesplit)            'and write
        FileWriterM = New StreamWriter(FilesplitWithM)            'and write

        strLine = FileReader.ReadLine               ' read the first line until no more data
        Do While Not strLine Is Nothing
            strLine = FileReader.ReadLine

            If CheckBoxMMessages.Checked Then
                Select Case True
                    Case (InStr(1, strLine, ExcludeComPortX) > 0)
                    Case (InStr(1, strLine, ExcludeComPortY) > 0)
                    Case (InStr(1, strLine, ExcludeComPortZ) > 0)

                    Case (InStr(1, strLine, WhatToLookFor) > 0)
                        FileWriter.WriteLine(strLine)
                        FileWriterM.WriteLine(strLine)
                    Case (InStr(1, strLine, "[M]") > 0)
                        FileWriterM.WriteLine(strLine)
                End Select

            Else
                If InStr(1, strLine, WhatToLookFor) Then
                    FileWriter.WriteLine(strLine)
                End If
            End If


            'string search also include all [M] debug messages from RDI            



        Loop

        FileReader.Close()
        FileWriter.Close()
        FileWriterM.Close()

    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'If My.Computer.Keyboard.AltKeyDown = True And My.Computer.Keyboard.CtrlKeyDown = True Then
        '    ExtraINloggingbutton.Enabled = True
        '    ExtraINloggingbutton.Show()
        'ElseIf My.Computer.Keyboard.AltKeyDown = False Or My.Computer.Keyboard.CtrlKeyDown = False Then
        '    ExtraINloggingbutton.Enabled = False
        '    ExtraINloggingbutton.Hide()
        'End If

        ClearForm()
        populatexcelsheetdetailsRowCounter = 3
        SelectedFolder.Text = driveletter + "RDI Analyser\EXORDI"
        makeblankexceltemplate()

        'set up pickers etc
        'DateoffilestoSplit = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        'DateTimePickertimestart.Format = DateTimePickerFormat.Time
        'DateTimePickertimestart.ShowUpDown = True
        'DateTimePickertimestart.CustomFormat = "DD-MMM-YYYY 00:00:00"
        'DateTimePickertimestart.Value = DateoffilestoSplit & " 00:00:00"

        'DateTimePickertimeend.Format = DateTimePickerFormat.Time
        'DateTimePickertimeend.ShowUpDown = True
        'DateTimePickertimeend.CustomFormat = "DD-MMM-YYYY 00:00:00"
        'DateTimePickertimeend.Value = DateoffilestoSplit & " 23:59:59"


    End Sub

    Private Sub FindLogFilesofDate(Folderpath As String, DateoffiletoSplit As String)
        '    Dim Folderpath As String
        Dim Filenametag As String
        Filenametag = DateoffiletoSplit & "-??????" & ".log"

        '  Folderpath = "d:\Users\jon\Documents\work\EXORDI\COM01"
        Dim fileNames = My.Computer.FileSystem.GetFiles(
           Folderpath, FileIO.SearchOption.SearchTopLevelOnly, Filenametag)

        For Each fileName As String In fileNames
            FilestosplitListbox.Items.Add(fileName)
        Next
    End Sub


    Private Sub SplitAllLogFiles()
        Dim DateoffilestoSplit As String
        Dim LogfilesFolder As String

        On Error GoTo Errornofilefound
        PleaseWait.Show()

        FilestosplitListbox.Items.Clear()
        DateoffilestoSplit = DateTimePickerSplitter.Value.Date.ToString("yyyyMMdd")
        LogfilesFolder = SelectedFolder.Text

        GetLogFilesofDate(LogfilesFolder & "\COM01", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM02", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM03", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM04", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM05", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM06", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM07", DateoffilestoSplit)

        GetLogFilesofDate(LogfilesFolder & "\COM08", DateoffilestoSplit)

        For Each filename As String In FilestosplitListbox.Items 'used comm ports
            Split_logs(filename, "COM23")
            Split_logs(filename, "COM24")
            Split_logs(filename, "COM25")
            Split_logs(filename, "COM26")
        Next
        PleaseWait.Close()

        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry No file or Path Found. Try browsing for the directory with select root")
        PleaseWait.Close()
    End Sub

    Private Sub GetLogFilesofDate(Folderpath As String, DateoffiletoSplit As String)

        Dim Filenametag As String
        Filenametag = DateoffiletoSplit & "-??????" & ".log"

        Dim fileNames = My.Computer.FileSystem.GetFiles(
           Folderpath, FileIO.SearchOption.SearchTopLevelOnly, Filenametag)


        For Each fileName As String In fileNames
            FilestosplitListbox.Items.Add(fileName)
        Next
        Exit Sub



    End Sub



    Private Sub GetLogFilesofDateCommport(Folderpath As String, DateoffiletoSplit As String, Commport As String, commstype As String, CommsDirection As String)
        '   FilestoAnalyseListbox.Items.Clear()


        '    Dim Folderpath As String
        Dim Filenametag As String
        Filenametag = Commport & "_" & DateoffiletoSplit & "*.log"
        Dim FilenametagM = Commport & "_M_" & DateoffiletoSplit & "*.log"

        '  Folderpath = d:\Users\jon\Documents\work\EXORDI\COM01"
        Dim fileNames = My.Computer.FileSystem.GetFiles(
           Folderpath, FileIO.SearchOption.SearchTopLevelOnly, Filenametag)
        Dim fileNamesM = My.Computer.FileSystem.GetFiles(
           Folderpath, FileIO.SearchOption.SearchTopLevelOnly, FilenametagM)

        For Each fileName As String In fileNames
            If commstype = "PSTN" And CommsDirection = "OUT" Then
                PSTNOUTFilestoAnalyseListbox.Items.Add(fileName)
            End If
            If commstype = "GSM" And CommsDirection = "OUT" Then
                GSMOUTFilestoAnalyseListbox.Items.Add(fileName)
            End If
            If commstype = "PSTN" And CommsDirection = "IN" Then
                PSTNINFilestoAnalyseListbox.Items.Add(fileName)
            End If
            If commstype = "GSM" And CommsDirection = "IN" Then
                GSMINFilestoAnalyseListbox.Items.Add(fileName)
            End If
            '          If InStr(1, FilenametagM, "_M_") Then
            '          MFilestoAnalyseListbox.Items.Add(fileName)
            '          End If
        Next


        For Each fileName As String In fileNamesM
            If InStr(1, FilenametagM, "_M_") Then
                MFilestoAnalyseListbox.Items.Add(fileName)
            End If
        Next

    End Sub



    Private Sub btnSelectFolder_Click(sender As Object, e As EventArgs) Handles btnSelectFolder.Click
        Dim Folder As String
        Folder = SelectFolder()
        '      MessageBox.Show(Folder)

    End Sub


    Private Function SelectFolder() As String
        Dim defaultfolder = driveletter + "RDI Analyser\EXORDI"
        Dim FilesFolder As String

        FolderBrowserDialog1.ShowDialog()
        FilesFolder = FolderBrowserDialog1.SelectedPath
        If FilesFolder = "" Then
            FilesFolder = defaultfolder
        End If
        SelectedFolder.Text = FilesFolder
        Return FilesFolder
    End Function

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ClearForm()

    End Sub


    Sub ClearForm()
        FilestosplitListbox.Items.Clear()
        PSTNOUTFilestoAnalyseListbox.Items.Clear()
        GSMOUTFilestoAnalyseListbox.Items.Clear()
        PSTNINFilestoAnalyseListbox.Items.Clear()
        GSMINFilestoAnalyseListbox.Items.Clear()
        FilestosplitListbox.Items.Clear()
        PSTNDialCount.Text = 0
        PSTNConnectCount.Text = 0
        PSTNINFailpercent.Text = 0
        PSTNNoReplyCount.Text = 0
        PSTNBusyCount.Text = 0
        PSTNNoCarrierCount.Text = 0

        GSMDialCount.Text = 0
        GSMConnectCount.Text = 0
        GSMNoReplyCount.Text = 0
        GSMBusyCount.Text = 0
        GSMINFailpercent.Text = 0
        GSMNoCarrierCount.Text = 0
        GSMINnoConnectCount.Text = 0
        GSMOUTGoodpercent.Text = 0
        GSMOUTFailedCount.Text = 0
        GSMOUTFailpercent.Text = 0
        GSMDisconnectCount.Text = 0
        PSTNRingnotreadyCount.Text = 0
        GSMRingnotreadyCount.Text = 0

        PSTNINRingCount.Text = 0
        PSTNINConnectpercent.Text = 0
        PSTNINnoCarrierpercent.Text = 0
        PSTNINwrongBaudpercent.Text = 0
        PSTNINConnectBAUDCount.Text = 0
        PSTNINfailedCount.Text = 0
        PSTNOUTGoodpercent.Text = 0
        PSTNOUTFailpercent.Text = 0
        PSTNDisconnectCount.Text = 0
        PSTNBlackListed.Text = 0

        PSTNINConnectedWrongBaudCount.Text = 0
        PSTNINAnsweringCount.Text = 0
        PSTNINNoCarrierCount.Text = 0
        PSTNINmultiRingCount.Text = 0
        PSTNINGoodTermalarmsCount.Text = 0
        PSTNINConnectBAUDCount.Text = 0
        PSTNINConnectBAUDCount.Text = 0
        PSTNINParityErrorCount.Text = 0

        GSMINGoodTermalarmsCount.Text = 0
        GSMINGoodTermtimesyncCount.Text = 0
        GSMINConnectBAUDCount.Text = 0
        GSMINParityErrorCount.Text = 0
        GSMOUTGoodTermalarmsCount.Text = 0
        GSMOUTGoodTermtimesyncCount.Text = 0
        GSMOUTBadLogsTermCount.Text = 0


        PSTNParityErrorCount.Text = 0
        PSTNOUTFailedCount.Text = 0
        PSTNOUTBadLogsTermCount.Text = 0
        PSTNOUTGoodTermalarmsCount.Text = 0
        PSTNOUTGoodTermtimesyncCount.Text = 0

        GSMParityErrorCount.Text = 0
        GSMINRingCount.Text = 0
        GSMINConnectpercent.Text = 0
        GSMINnoCarrierpercent.Text = 0
        GSMINwrongBaudpercent.Text = 0
        GSMINConnectBAUDCount.Text = 0
        GSMINConnectedWrongBaudCount.Text = 0
        GSMINAnsweringCount.Text = 0
        GSMINNoCarrierCount.Text = 0
        PSTNProdtimeResultCount.Text = 0

        'tickallforallsave(False)

    End Sub

    Private Sub CountOutPSTN()
        Dim FileReader As StreamReader
        Dim strLine As String
        Dim addressposition As Integer
        Dim exoaddress As String = "XXXX"
        Dim PLAH As String = "PLA"
        Dim ELAH As String = "ELA"
        Dim timeofconnect As String = "00:00:00"
        Dim timeofdisconnect As String = "00:00:00"
        Dim timeofnoCarrier As String = "00:00:00"
        Dim timeofpreviouscleardown As String = "00:00:00"
        Dim CommPortDetails As String = ""
        'Dim logedmessage As String = "nothing to report"
        Dim logedmessagetimesync As String = "nothing to report"
        Dim logedmessagelogs As String = "nothing to report"
        Dim logedmessagealarms As String = "nothing to report"

        Dim timeofcleardown As String = ""
        Dim numbertocheckfor As String

        Dim DialCount As Integer
        DialCount = 0
        PSTNDialCount.Text = DialCount

        Dim ConnectCount As Integer
        ConnectCount = 0
        PSTNConnectCount.Text = ConnectCount

        Dim GoodTermAlarmsCount As Integer
        GoodTermAlarmsCount = 0
        PSTNOUTGoodTermalarmsCount.Text = GoodTermAlarmsCount

        Dim GoodTermtimesyncCount As Integer
        GoodTermtimesyncCount = 0
        PSTNOUTGoodTermtimesyncCount.Text = GoodTermtimesyncCount


        Dim BadLogsCount As Integer
        BadLogsCount = 0
        PSTNOUTBadLogsTermCount.Text = BadLogsCount

        Dim NoReplyCount As Integer
        NoReplyCount = 0
        PSTNNoReplyCount.Text = NoReplyCount

        Dim PSTNBlackListedcount As Integer
        PSTNBlackListedcount = 0
        PSTNBlackListed.Text = PSTNBlackListedcount

        Dim PSTNDisconnectcounter As Integer
        PSTNDisconnectcounter = 0
        PSTNDisconnectCount.Text = PSTNDisconnectcounter

        Dim PSTNOUTFailedCounter As Integer
        PSTNOUTFailedCounter = 0
        PSTNOUTFailedCount.Text = PSTNOUTFailedCounter

        Dim BusyCount As Integer
        BusyCount = 0
        PSTNBusyCount.Text = BusyCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        PSTNNoCarrierCount.Text = NoCarrierCount

        Dim PSTNOUTParityErrorCounter As Integer
        PSTNOUTParityErrorCounter = 0
        PSTNParityErrorCount.Text = PSTNOUTParityErrorCounter


        Dim ProdtimeResultCount As Integer
        ProdtimeResultCount = 0
        PSTNProdtimeResultCount.Text = ProdtimeResultCount

        Dim ParityErrorthiscallCount As Integer
        ParityErrorthiscallCount = 0

        Dim devicephonenumber As String = "XXXX"
        Dim timeofdialrequest As String = "00:00:00"


        Dim FirstCharacter As Integer
        Dim ModemIsIdle = False
        Dim ModemIsConfiged = False
        Dim ModemIsDialing = False
        'Dim ModemIsAnswering = False
        Dim ModemIsConnectedCorrectBaud = False
        Dim ModemIsConnectedWRONGBaud = False
        Dim ModemIsNoCarrier = False
        Dim ModemIsCallActive = False
        Dim ModemIsGotAddress = False
        Dim ModemIsGoodtimesync = False
        Dim ModemIsCallTerminated = False
        Dim ModemIsBadLogs = False
        Dim ModemIsAlarmRecovery = False
        Dim ThisCallHasManyParityerrors = False
        exoaddress = "XXXX"


        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        ' Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        ' '''' leave to workbook open to speed things up

        For Each filename As String In PSTNOUTFilestoAnalyseListbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data

            Do While Not strLine Is Nothing
                'loop through messages and see where we are within the call
                ' read message to see where we are
                Try
                    Dim TestPosition As Integer = InStr(1, filename, "\EXORDI\", CompareMethod.Text) + 8 'add on 7 for the test string length too
                    CommPortDetails = Mid(filename, TestPosition, 11)

                    strLine = FileReader.ReadLine

                    Select Case True
                        Case strLine.Contains("+++")
                            ModemIsDialing = False
                            ModemIsCallTerminated = True
                            ModemIsConfiged = False
                            exoaddress = "XXXX"
                            ModemIsAlarmRecovery = False

                        Case strLine.Contains("AT&F")
                            ModemIsDialing = False
                            ModemIsCallTerminated = True
                            ModemIsConfiged = True

                        Case strLine.Contains("ATDT9")
                            'new call so reset the flags for sucesss
                            ModemIsGoodtimesync = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsDialing = True
                            ModemIsGotAddress = False
                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False
                            ModemIsCallTerminated = False
                            ModemIsBadLogs = False
                            ModemIsAlarmRecovery = False

                            DialCount = DialCount + 1
                            PSTNDialCount.Text = DialCount
                            timeofdialrequest = Microsoft.VisualBasic.Left(strLine, 12)
                            FirstCharacter = strLine.IndexOf("ATDT")
                            devicephonenumber = Microsoft.VisualBasic.Mid(strLine, FirstCharacter + 7, 11)

                            If CheckBoxoutphonenumber.Checked Then
                                numbertocheckfor = TextBoxphonenumber.Text

                                FirstCharacter = strLine.IndexOf(numbertocheckfor)
                                '  timeofdialrequest = Microsoft.VisualBasic.Left(strLine, 12)
                                '  devicephonenumber = Microsoft.VisualBasic.Mid(strLine, FirstCharacter + 7, 11)

                                If CheckBoxoutphonenumber.Checked Then
                                    numbertocheckfor = TextBoxphonenumber.Text

                                    populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 20, "Dial to Site of interest", CommPortDetails, numbertocheckfor, "Dialled at ", timeofdialrequest)
                                End If
                            End If

                        Case strLine.Contains("10F100")
                            ModemIsGotAddress = True
                            addressposition = InStr(1, strLine, "10F100")
                            'get PLA and ELS in Hex string values
                            PLAH = Mid(strLine, addressposition - 4, 2)
                            ELAH = Mid(strLine, addressposition - 2, 2)
                            'convert to decimal
                            Dim PLA = CInt("&H" & PLAH)
                            Dim ELA = CInt("&H" & ELAH)
                            exoaddress = PLA & ":" & ELA & " (Hex-" & PLAH & ELAH & ")"




                        Case strLine.Contains("no reply received")
                            NoReplyCount = NoReplyCount + 1
                            PSTNNoReplyCount.Text = NoReplyCount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Reply"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If

                        Case strLine.Contains("BUSY<CR><LF>")
                            BusyCount = BusyCount + 1
                            PSTNBusyCount.Text = BusyCount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Busy Device"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If

                        Case strLine.Contains("NO CARRIER")
                            NoCarrierCount = NoCarrierCount + 1
                            PSTNNoCarrierCount.Text = NoCarrierCount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Carrier"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If
                        Case strLine.Contains("BLACKLISTED")
                            PSTNBlackListedcount = PSTNBlackListedcount + 1
                            PSTNBlackListed.Text = PSTNBlackListedcount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Blacklisted"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If
                        Case strLine.Contains("Parity ")

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            End If

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount >= 3 Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                                ThisCallHasManyParityerrors = True
                                '
                            End If


                        Case strLine.Contains("CONNECT 2400<CR><LF>")
                            ConnectCount = ConnectCount + 1
                            PSTNConnectCount.Text = ConnectCount

                            ModemIsConnectedCorrectBaud = True

                            'new call so reset the flags for sucesss
                            ModemIsCallTerminated = False
                            ModemIsGoodtimesync = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsDialing = False
                            ModemIsGotAddress = False


                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False

                            timeofconnect = Microsoft.VisualBasic.Left(strLine, 12)

                            '0F100
                            'was "A401A401A401A401"
                            '102900                             'Alarm download complete for module
                        Case strLine.Contains("102900")

                            'alarm recovery 507220
                            ModemIsAlarmRecovery = True'

                        Case strLine.Contains("A4A43E") ' 
                            'log recovery failed
                            ModemIsBadLogs = True


                        Case strLine.Contains("0F1000B") 'set the year
                            If ModemIsGotAddress = True Then
                                'time sync messages
                                ModemIsGoodtimesync = True

                            End If




                        Case strLine.Contains("DISCONNECTING")
                            ModemIsCallTerminated = True
                            PSTNDisconnectcounter = PSTNDisconnectcounter + 1
                            PSTNDisconnectCount.Text = PSTNDisconnectcounter

                            timeofdisconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            'here check if the termination is from an existing good call
                            '
                            'also check if we got the logs or not



                            If ModemIsGoodtimesync = True Then 'And ModemIsBadLogs Then ' = False And ModemIsAlarmRecovery = True Then
                                logedmessagetimesync = "Call Termination (with time sync)"
                                GoodTermtimesyncCount = GoodTermtimesyncCount + 1
                                PSTNOUTGoodTermtimesyncCount.Text = GoodTermtimesyncCount
                            Else
                                logedmessagetimesync = "Call Termination (no time sync)"
                            End If

                            If ModemIsBadLogs = True Then
                                logedmessagelogs = "Bad Logs"
                                BadLogsCount = BadLogsCount + 1
                                PSTNOUTBadLogsTermCount.Text = BadLogsCount
                            Else
                                logedmessagelogs = "Good Logs"
                            End If

                            If ModemIsAlarmRecovery = True Then
                                logedmessagealarms = "Alarms Recovered"
                                GoodTermAlarmsCount = GoodTermAlarmsCount + 1
                                PSTNOUTGoodTermalarmsCount.Text = GoodTermAlarmsCount
                            Else
                                logedmessagealarms = "No Alarms Recovered"
                            End If





                            If ThisCallHasManyParityerrors = True Then
                                PSTNOUTParityErrorCounter = PSTNOUTParityErrorCounter + 1
                                PSTNParityErrorCount.Text = PSTNOUTParityErrorCounter
                            End If
                            If CheckBoxLogOut.Checked And timeofconnect <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated at " & timeofdisconnect & " From Dial Sent at " & timeofdialrequest
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                If ParityErrorthiscallCount > 0 Then
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors " & ParityErrorthiscallCount
                                Else
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Parity errors "
                                End If
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & exoaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = logedmessagetimesync
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 6).Value = logedmessagelogs
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 7).Value = logedmessagealarms

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                            End If



                    End Select


                Catch e As System.NullReferenceException
                Catch e As FormatException
                Catch e As OverflowException
                Catch e As System.InvalidCastException 'miss interprited string to dates
                    '  strLine = FileReader.ReadLine 'move on
                End Try



            Loop

            FileReader.Close()
        Next
        ' save and exit
        oBook.Save()
        oBook.Application.DisplayAlerts = False
        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()

        PSTNOUTFailedCounter = NoCarrierCount + NoReplyCount + PSTNOUTParityErrorCounter + PSTNBlackListedcount
        PSTNOUTFailedCount.Text = PSTNOUTFailedCounter
        '
        PSTNOUTGoodpercent.Text = String.Format("{0:n2}", PSTNDisconnectcounter / DialCount * 100)

        PSTNOUTFailpercent.Text = String.Format("{0:n2}", (DialCount - PSTNDisconnectcounter) / DialCount * 100)


    End Sub

    Private Sub CountOutGSM()
        Dim FileReader As StreamReader
        Dim strLine As String
        Dim addressposition As Integer
        Dim exoaddress As String = "XXXX"
        Dim PLAH As String = "PLA"
        Dim ELAH As String = "ELA"
        Dim timeofconnect As String = "00:00:00"
        Dim timeofdisconnect As String = "00:00:00"
        Dim timeofnoCarrier As String = "00:00:00"
        Dim timeofpreviouscleardown As String = "00:00:00"
        Dim CommPortDetails As String = ""
        'Dim logedmessage As String = "nothing to report"
        Dim logedmessagetimesync As String = "nothing to report"
        Dim logedmessagelogs As String = "nothing to report"
        Dim logedmessagealarms As String = "nothing to report"

        Dim timeofcleardown As String = ""
        Dim numbertocheckfor As String

        Dim DialCount As Integer
        DialCount = 0
        GSMDialCount.Text = DialCount

        Dim ConnectCount As Integer
        ConnectCount = 0
        GSMConnectCount.Text = ConnectCount

        Dim GoodTermAlarmsCount As Integer
        GoodTermAlarmsCount = 0
        GSMOUTGoodTermalarmsCount.Text = GoodTermAlarmsCount

        Dim GoodTermtimesyncCount As Integer
        GoodTermtimesyncCount = 0
        GSMOUTGoodTermtimesyncCount.Text = GoodTermtimesyncCount


        Dim BadLogsCount As Integer
        BadLogsCount = 0
        GSMOUTBadLogsTermCount.Text = BadLogsCount

        Dim NoReplyCount As Integer
        NoReplyCount = 0
        GSMNoReplyCount.Text = NoReplyCount

        Dim GSMBlackListedcount As Integer
        GSMBlackListedcount = 0
        GSMBlackListed.Text = GSMBlackListedcount

        Dim GSMDisconnectcounter As Integer
        GSMDisconnectcounter = 0
        GSMDisconnectCount.Text = GSMDisconnectcounter

        Dim GSMOUTFailedCounter As Integer
        GSMOUTFailedCounter = 0
        GSMOUTFailedCount.Text = GSMOUTFailedCounter

        Dim BusyCount As Integer
        BusyCount = 0
        GSMBusyCount.Text = BusyCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        GSMNoCarrierCount.Text = NoCarrierCount

        Dim GSMOUTParityErrorCounter As Integer
        GSMOUTParityErrorCounter = 0
        GSMParityErrorCount.Text = GSMOUTParityErrorCounter


        '   Dim ProdtimeResultCount As Integer
        '   ProdtimeResultCount = 0
        '   GSMProdtimeResultCount.Text = ProdtimeResultCount

        Dim ParityErrorthiscallCount As Integer
        ParityErrorthiscallCount = 0

        Dim devicephonenumber As String = "XXXX"
        Dim timeofdialrequest As String = "00:00:00"


        Dim FirstCharacter As Integer
        Dim ModemIsIdle = False
        Dim ModemIsConfiged = False
        Dim ModemIsDialing = False
        'Dim ModemIsAnswering = False
        Dim ModemIsConnectedCorrectBaud = False
        Dim ModemIsConnectedWRONGBaud = False
        Dim ModemIsNoCarrier = False
        Dim ModemIsCallActive = False
        Dim ModemIsGotAddress = False
        Dim ModemIsGoodtimesync = False
        Dim ModemIsCallTerminated = False
        Dim ModemIsBadLogs = False
        Dim ModemIsAlarmRecovery = False
        Dim ThisCallHasManyParityerrors = False
        exoaddress = "XXXX"


        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        ' Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        ' '''' leave to workbook open to speed things up

        For Each filename As String In GSMOUTFilestoAnalyseListbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data

            Do While Not strLine Is Nothing
                'loop through messages and see where we are within the call
                ' read message to see where we are
                Try
                    Dim TestPosition As Integer = InStr(1, filename, "\EXORDI\", CompareMethod.Text) + 8 'add on 7 for the test string length too
                    CommPortDetails = Mid(filename, TestPosition, 11)

                    strLine = FileReader.ReadLine

                    Select Case True
                        Case strLine.Contains("+++")
                            ModemIsDialing = False
                            ModemIsCallTerminated = True
                            ModemIsConfiged = False
                            exoaddress = "XXXX"
                            ModemIsAlarmRecovery = False

                        Case strLine.Contains("AT&F")
                            ModemIsDialing = False
                            ModemIsCallTerminated = True
                            ModemIsConfiged = True

                        Case strLine.Contains("ATDT9")
                            'new call so reset the flags for sucesss
                            ModemIsGoodtimesync = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsDialing = True
                            ModemIsGotAddress = False
                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False
                            ModemIsCallTerminated = False
                            ModemIsBadLogs = False
                            ModemIsAlarmRecovery = False

                            DialCount = DialCount + 1
                            GSMDialCount.Text = DialCount
                            timeofdialrequest = Microsoft.VisualBasic.Left(strLine, 12)
                            FirstCharacter = strLine.IndexOf("ATDT")
                            devicephonenumber = Microsoft.VisualBasic.Mid(strLine, FirstCharacter + 7, 11)

                            If CheckBoxoutphonenumber.Checked Then
                                numbertocheckfor = TextBoxphonenumber.Text

                                FirstCharacter = strLine.IndexOf(numbertocheckfor)
                                '  timeofdialrequest = Microsoft.VisualBasic.Left(strLine, 12)
                                '  devicephonenumber = Microsoft.VisualBasic.Mid(strLine, FirstCharacter + 7, 11)

                                If CheckBoxoutphonenumber.Checked Then
                                    numbertocheckfor = TextBoxphonenumber.Text

                                    populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 20, "Dial to Site of interest", CommPortDetails, numbertocheckfor, "Dialled at ", timeofdialrequest)
                                End If
                            End If

                        Case strLine.Contains("10F100")
                            ModemIsGotAddress = True
                            addressposition = InStr(1, strLine, "10F100")
                            'get PLA and ELS in Hex string values
                            PLAH = Mid(strLine, addressposition - 4, 2)
                            ELAH = Mid(strLine, addressposition - 2, 2)
                            'convert to decimal
                            Dim PLA = CInt("&H" & PLAH)
                            Dim ELA = CInt("&H" & ELAH)
                            exoaddress = PLA & ":" & ELA & " (Hex-" & PLAH & ELAH & ")"




                        Case strLine.Contains("no reply received")
                            NoReplyCount = NoReplyCount + 1
                            GSMNoReplyCount.Text = NoReplyCount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Reply"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If

                        Case strLine.Contains("BUSY<CR><LF>")
                            BusyCount = BusyCount + 1
                            GSMBusyCount.Text = BusyCount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Busy Device"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If

                        Case strLine.Contains("NO CARRIER<CR><LF>")
                            '   

                            'on GSM we get a NO Carrier after +++ disconnect so check if normal
                            If ModemIsCallTerminated = False Then
                                NoCarrierCount = NoCarrierCount + 1
                                GSMNoCarrierCount.Text = NoCarrierCount


                                If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then

                                    oSheet = oBook.Worksheets("EXO")
                                    Dim col As Integer = 20

                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Carrier"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                                End If



                            End If



                        Case strLine.Contains("BLACKLISTED")
                            GSMBlackListedcount = GSMBlackListedcount + 1
                            GSMBlackListed.Text = GSMBlackListedcount
                            If CheckBoxLogOut.Checked And timeofdialrequest <> "" Then
                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Blacklisted"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Dialed Number " & devicephonenumber
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Dial Sent at " & timeofdialrequest
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                            End If
                        Case strLine.Contains("Parity ")

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            End If

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount >= 3 Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                                ThisCallHasManyParityerrors = True

                            End If


                        Case strLine.Contains("CONNECT 9600<CR><LF>")
                            ConnectCount = ConnectCount + 1
                            GSMConnectCount.Text = ConnectCount

                            ModemIsConnectedCorrectBaud = True

                            'new call so reset the flags for sucesss
                            ModemIsCallTerminated = False
                            ModemIsGoodtimesync = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsDialing = False
                            ModemIsGotAddress = False


                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False

                            timeofconnect = Microsoft.VisualBasic.Left(strLine, 12)


                            'was "A401A401A401A401"
                            '102900                             'Alarm download complete for module
                        Case strLine.Contains("102900")


                            ModemIsAlarmRecovery = True'

                        Case strLine.Contains("A4A43E") ' 
                            'log recovery failed
                            ModemIsBadLogs = True
                        Case strLine.Contains("0F1000B") 'set the year
                            If ModemIsGotAddress = True Then
                                'time sync messages
                                ModemIsGoodtimesync = True

                            End If




                        Case strLine.Contains("DISCONNECTING")
                            ModemIsCallTerminated = True
                            GSMDisconnectcounter = GSMDisconnectcounter + 1
                            GSMDisconnectCount.Text = GSMDisconnectcounter

                            timeofdisconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            'here check if the termination is from an existing good call
                            '
                            'also check if we got the logs or not



                            If ModemIsGoodtimesync = True Then 'And ModemIsBadLogs Then ' = False And ModemIsAlarmRecovery = True Then
                                logedmessagetimesync = "Call Termination (with time sync)"
                                GoodTermtimesyncCount = GoodTermtimesyncCount + 1
                                GSMOUTGoodTermtimesyncCount.Text = GoodTermtimesyncCount
                            Else
                                logedmessagetimesync = "Call Termination (no time sync)"
                            End If

                            If ModemIsBadLogs = True Then
                                logedmessagelogs = "Bad Logs"
                                BadLogsCount = BadLogsCount + 1
                                GSMOUTBadLogsTermCount.Text = BadLogsCount
                            Else
                                logedmessagelogs = "Good Logs"
                            End If

                            If ModemIsAlarmRecovery = True Then
                                logedmessagealarms = "Alarms Recovered"
                                GoodTermAlarmsCount = GoodTermAlarmsCount + 1
                                GSMOUTGoodTermalarmsCount.Text = GoodTermAlarmsCount
                            Else
                                logedmessagealarms = "No Alarms Recovered"
                            End If





                            If ThisCallHasManyParityerrors = True Then
                                GSMOUTParityErrorCounter = GSMOUTParityErrorCounter + 1
                                GSMParityErrorCount.Text = GSMOUTParityErrorCounter
                            End If
                            If CheckBoxLogOut.Checked And timeofconnect <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated at " & timeofdisconnect & " From Dial Sent at " & timeofdialrequest
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                If ParityErrorthiscallCount > 0 Then
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors " & ParityErrorthiscallCount
                                Else
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Parity errors "
                                End If
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & exoaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = logedmessagetimesync
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 6).Value = logedmessagelogs
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 7).Value = logedmessagealarms

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                            End If



                    End Select


                Catch e As System.NullReferenceException
                Catch e As FormatException
                Catch e As OverflowException
                Catch e As System.InvalidCastException 'miss interprited string to dates
                    '  strLine = FileReader.ReadLine 'move on
                End Try



            Loop

            FileReader.Close()
        Next
        ' save and exit
        oBook.Save()
        oBook.Application.DisplayAlerts = False
        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()

        GSMOUTFailedCounter = NoCarrierCount + NoReplyCount + GSMOUTParityErrorCounter + GSMBlackListedcount
        GSMOUTFailedCount.Text = GSMOUTFailedCounter
        '
        GSMOUTGoodpercent.Text = String.Format("{0:n2}", GSMDisconnectcounter / DialCount * 100)

        GSMOUTFailpercent.Text = String.Format("{0:n2}", (DialCount - GSMDisconnectcounter) / DialCount * 100)


    End Sub




    Private Sub CountINGSM()

        Dim FileReader As StreamReader
        Dim strLine As String
        Dim addressposition As Integer
        Dim exoaddress As String = "XXXX"
        Dim PLAH As String = "PLA"
        Dim ELAH As String = "ELA"

        Dim RingCount As Integer
        RingCount = 0
        GSMINRingCount.Text = RingCount

        Dim ConnectBAUDCount As Integer
        ConnectBAUDCount = 0
        GSMINConnectBAUDCount.Text = ConnectBAUDCount

        Dim GSMINfailedCounter As Integer
        GSMINfailedCounter = 0
        GSMINnoConnectCount.Text = GSMINfailedCounter

        Dim GSMINParityErrorCounter As Integer
        GSMINParityErrorCounter = 0
        GSMINParityErrorCount.Text = GSMINParityErrorCounter


        Dim GSMINmultiRingCounter As Integer
        GSMINmultiRingCounter = 0
        GSMINmultiRingCount.Text = GSMINmultiRingCounter

        Dim GSMRingnotreadyCounter As Integer
        GSMRingnotreadyCounter = 0
        GSMRingnotreadyCount.Text = GSMRingnotreadyCounter

        Dim ConnectedWrongBaudCount As Integer
        ConnectedWrongBaudCount = 0
        GSMINConnectedWrongBaudCount.Text = ConnectedWrongBaudCount

        Dim AnsweringCount As Integer
        AnsweringCount = 0
        GSMINAnsweringCount.Text = AnsweringCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        GSMINNoCarrierCount.Text = NoCarrierCount

        Dim ParityErrorthiscallCount As Integer
        ParityErrorthiscallCount = 0

        Dim GoodTermCount As Integer
        GoodTermCount = 0
        GSMINGoodTermtimesyncCount.Text = GoodTermCount

        Dim GoodTermAlarmsCount As Integer
        GoodTermAlarmsCount = 0
        GSMINGoodTermalarmsCount.Text = GoodTermAlarmsCount

        Dim GoodTermtimesyncCount As Integer
        GoodTermtimesyncCount = 0
        GSMINGoodTermtimesyncCount.Text = GoodTermtimesyncCount


        Dim GSMINnoaddressCounter As Integer = 0

        Dim BadLogsCount As Integer
        BadLogsCount = 0
        '  PSTNINBadLogsTermCount.Text = BadLogsCount


        Dim ModemIsIdle As Boolean
        Dim ModemIsConfiged As Boolean
        Dim ModemIsRinging As Boolean
        Dim ModemIsAnswering As Boolean
        Dim ModemIsConnectedCorrectBaud As Boolean
        Dim ModemIsConnectedWRONGBaud As Boolean
        Dim ModemIsNoCarrier As Boolean
        Dim ModemIsCallActive As Boolean
        Dim ModemIsGotAddress As Boolean
        Dim ModemIsGoodCallEnd As Boolean
        Dim ModemIsCallTerminated As Boolean
        Dim ModemIsGoodtimesync = False
        Dim ModemIsBadLogs = False
        Dim ModemIsAlarmRecovery = False
        Dim ThisCallHasManyParityerrors As Boolean



        Dim CallTerminaledCount As Integer
        CallTerminaledCount = 0


        Dim timeofring As String = "00:00:00"
        Dim timeofconnect As String = "00:00:00"
        Dim timeofdisconnect As String = "00:00:00"
        Dim timeofnoCarrier As String = "00:00:00"
        Dim timeofpreviouscleardown As String = "00:00:00"
        Dim CommPortDetails As String = ""
        Dim logedmessage As String = "nothing to report"
        Dim timeofcleardown As String = ""
        Dim realtimeofcleardown As Date
        Dim realtimeoftimeofring As Date

        Dim FirstCharacter As Integer
        Dim logedmessagetimesync As String = "nothing to report"
        Dim logedmessagelogs As String = "nothing to report"
        Dim logedmessagealarms As String = "nothing to report"

        ''''
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        ''start new workbook
        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        ' Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        ' '''' leave to workbook open to speed things up

        For Each filename As String In GSMINFilestoAnalyseListbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data

            ModemIsIdle = False
            ModemIsConfiged = False
            ModemIsRinging = False
            ModemIsAnswering = False
            ModemIsConnectedCorrectBaud = False
            ModemIsConnectedWRONGBaud = False
            ModemIsNoCarrier = False
            ModemIsCallActive = False
            ModemIsGotAddress = False
            ModemIsGoodCallEnd = False
            ModemIsCallTerminated = False
            ThisCallHasManyParityerrors = False
            exoaddress = "XXXX"


            Do While Not strLine Is Nothing




                Try

                    Dim TestPosition As Integer = InStr(1, filename, "\EXORDI\", CompareMethod.Text) + 8 'add on 7 for the test string length too
                    CommPortDetails = Mid(filename, TestPosition, 11)





                    'loop through messages and see where we are within the call

callprogress:' read message to see where we are

                    strLine = FileReader.ReadLine


                    Select Case True
                        Case strLine.Contains("+++")
                            ModemIsCallTerminated = True
                            ModemIsConfiged = False
                            ModemIsRinging = False
                            exoaddress = "XXXX"
                            ModemIsAlarmRecovery = False
                        Case strLine.Contains("AUTO-ANSWER (state 5)")
                            ModemIsCallTerminated = False
                            ModemIsConfiged = True

                            AnsweringCount = AnsweringCount + 1
                            GSMINAnsweringCount.Text = AnsweringCount

                            If CheckBoxidletime.Checked Then
                                Dim busytime As System.TimeSpan
                                busytime = realtimeofcleardown - realtimeoftimeofring
                                Dim totalbusytime As System.TimeSpan
                                totalbusytime = totalbusytime + busytime
                                ''''

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")

                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Modem ready"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "Received at "
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofcleardown
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                col = 21

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "busyduration"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = timeofcleardown
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = busytime.ToString
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = totalbusytime.ToString

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                                '       populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 21, "busyduration", CommPortDetails, timeofring, timeofcleardown, busytime.ToString)
                            End If




                        Case strLine.Contains("NG<CR><LF>")
                            timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                            If ModemIsRinging = True Then
                                ' alredy ringing
                                'may ring >1 now so ok
                                ' timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                                GSMINmultiRingCounter = GSMINmultiRingCounter + 1
                                GSMINmultiRingCount.Text = GSMINmultiRingCounter
                            Else
                                'new ring
                                ModemIsRinging = True
                                ' timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                            End If


                            If ModemIsRinging = True And ModemIsConfiged = False Then
                                'modem not configured yet so may fail
                                GSMRingnotreadyCounter = GSMRingnotreadyCounter + 1
                                GSMRingnotreadyCount.Text = GSMRingnotreadyCounter
                                If CheckBoxLogIn.Checked And timeofring <> "" Then

                                    '   add Data
                                    oSheet = oBook.Worksheets("EXO")
                                    Dim col As Integer = 20

                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring without modem ready"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "Ring without modem ready"
                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                End If
                            End If

                        Case strLine.Contains("ANSWERING (state 9)")
                            ModemIsAnswering = True
                            RingCount = RingCount + 1
                            GSMINRingCount.Text = RingCount

                        Case strLine.Contains("NO CARRIER<CR><LF>")
                            ModemIsNoCarrier = True
                            ModemIsRinging = False
                            'on GSM we get a NO Carrier after +++ disconnect so check if normal
                            If ModemIsConfiged = True Then
                                If CheckBoxLogIn.Checked And timeofring <> "" Then
                                    FirstCharacter = strLine.IndexOf("NO CARRIER<CR><LF>")
                                    timeofnoCarrier = Microsoft.VisualBasic.Left(strLine, 12)


                                    ' populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 20, "No Carrier", CommPortDetails, "", "Received at ", timeofring)
                                    oSheet = oBook.Worksheets("EXO")
                                    Dim col As Integer = 20

                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Carrier"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "No Carrier"

                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                End If
                                NoCarrierCount = NoCarrierCount + 1
                                GSMINNoCarrierCount.Text = NoCarrierCount
                                'fail counter
                                GSMINfailedCounter = GSMINfailedCounter + 1
                                GSMINnoConnectCount.Text = GSMINfailedCounter
                                ' ThisCallIsConnected = False
                                '   GoTo foundresult
                            End If


                        Case strLine.Contains("Rx: CONNECT 1200")
                            ModemIsConnectedWRONGBaud = True
                            ModemIsConnectedCorrectBaud = False
                            ModemIsRinging = False
                            GSMINfailedCounter = GSMINfailedCounter + 1
                            ConnectedWrongBaudCount = ConnectedWrongBaudCount + 1
                            GSMINConnectedWrongBaudCount.Text = ConnectedWrongBaudCount
                            If CheckBoxLogIn.Checked And timeofring <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Wrong Baud"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If
                        Case strLine.Contains("Rx: CONNECT 2400")
                            ModemIsConnectedWRONGBaud = True
                            ModemIsConnectedCorrectBaud = False
                            ModemIsRinging = False
                            GSMINfailedCounter = GSMINfailedCounter + 1
                            ConnectedWrongBaudCount = ConnectedWrongBaudCount + 1
                            GSMINConnectedWrongBaudCount.Text = ConnectedWrongBaudCount
                            If CheckBoxLogIn.Checked And timeofring <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Wrong Baud"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If
                        Case strLine.Contains("Rx: CONNECT 9600<CR><LF>")
                            ModemIsConnectedCorrectBaud = True

                            'new call so reset the flags for sucesss
                            ModemIsGoodCallEnd = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsRinging = False
                            ModemIsGotAddress = False
                            ModemIsGoodCallEnd = False

                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False

                            timeofconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            ConnectBAUDCount = ConnectBAUDCount + 1
                            GSMINConnectBAUDCount.Text = ConnectBAUDCount


                        Case strLine.Contains("Parity")

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            End If

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount >= 3 Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                                ThisCallHasManyParityerrors = True

                            End If





                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            Else
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                                ThisCallHasManyParityerrors = True

                            End If

                        Case strLine.Contains("10F100")
                            ModemIsGotAddress = True
                            addressposition = InStr(1, strLine, "10F100")
                            'get PLA and ELS in Hex string values
                            PLAH = Mid(strLine, addressposition - 4, 2)
                            ELAH = Mid(strLine, addressposition - 2, 2)
                            'convert to decimal
                            Dim PLA = CInt("&H" & PLAH)
                            Dim ELA = CInt("&H" & ELAH)
                            exoaddress = PLA & ":" & ELA & " (Hex-" & PLAH & ELAH & ")"



                       'was "A401A401A401A401"
                            '102900                             'Alarm download complete for module
                        Case strLine.Contains("102900")
                            'alarm recovery 
                            ModemIsAlarmRecovery = True'

                        Case strLine.Contains("A4A43E") ' 
                            'log recovery failed
                            ModemIsBadLogs = True

                        Case strLine.Contains("0F100")
                            If ModemIsGotAddress = True Then
                                'time sync messages
                                ModemIsGoodtimesync = True
                                GoodTermCount = GoodTermCount + 1
                                PSTNINGoodTermtimesyncCount.Text = GoodTermCount
                            End If





                        Case strLine.Contains("DISCONNECTING (state 8)")


                            ModemIsCallTerminated = True

                            If ModemIsGotAddress = False Then
                                GSMINnoaddressCounter = GSMINnoaddressCounter + 1
                                GSMINfailedCounter = GSMINfailedCounter + 1
                            End If

                            timeofdisconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            If ThisCallHasManyParityerrors = True Then
                                GSMINParityErrorCounter = GSMINParityErrorCounter + 1
                            End If

                            If ModemIsGoodtimesync = True Then 'And ModemIsBadLogs Then ' = False And ModemIsAlarmRecovery = True Then
                                logedmessagetimesync = "Call Termination (with time sync)"
                                GoodTermtimesyncCount = GoodTermtimesyncCount + 1
                                GSMINGoodTermtimesyncCount.Text = GoodTermtimesyncCount
                            Else
                                logedmessagetimesync = "Call Termination (no time sync)"
                            End If


                            'no logs collected on incomming 
                            'If ModemIsBadLogs = True Then
                            '    logedmessagelogs = "Bad Logs"
                            '    BadLogsCount = BadLogsCount + 1
                            '    '          GSMINBadLogsTermCount.Text = BadLogsCount
                            'Else
                            '    logedmessagelogs = "Good Logs"
                            'End If

                            If ModemIsAlarmRecovery = True Then
                                logedmessagealarms = "Alarms Recovered"
                                GoodTermAlarmsCount = GoodTermAlarmsCount + 1
                                GSMINGoodTermalarmsCount.Text = GoodTermAlarmsCount
                            Else
                                logedmessagealarms = "No Alarms Recovered"
                            End If



                            If CheckBoxLogIn.Checked And timeofconnect <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated" & " from ring recieved at " & timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                If ParityErrorthiscallCount > 0 Then
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors " & ParityErrorthiscallCount
                                Else
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Parity errors "
                                End If
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & exoaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = logedmessagetimesync
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 6).Value = logedmessagelogs
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 7).Value = logedmessagealarms

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1



                            End If



                            'timeofdisconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            ''here check if the termination is from an existing or new call
                            ''new call 
                            'If ModemIsGoodCallEnd = True Then logedmessage = "Good Call Termination"
                            'If ModemIsGoodCallEnd = False Then logedmessage = "Bad Call Termination"
                            'If CheckBoxLogIn.Checked And timeofconnect <> "" Then

                            '    '   add Data
                            '    oSheet = oBook.Worksheets("EXO")
                            '    Dim col As Integer = 20

                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated" & " from ring recieved at " & timeofring

                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors " & ParityErrorthiscallCount
                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & exoaddress
                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                            '    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = logedmessage


                            '    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                            'End If


                    End Select



                Catch e As System.NullReferenceException
                Catch e As FormatException
                Catch e As OverflowException
                Catch e As System.InvalidCastException 'miss interprited string to dates
                    '  strLine = FileReader.ReadLine 'move on
                End Try



            Loop


            FileReader.Close()

        Next
        ''''
        ' save and exit
        oBook.Save()
        oBook.Application.DisplayAlerts = False
        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()


        ''''


        GSMINnoConnectCount.Text = GSMINfailedCounter
        GSMINConnectpercent.Text = String.Format("{0:n2}", (ConnectBAUDCount / RingCount) * 100)
        GSMINFailpercent.Text = String.Format("{0:n2}", (GSMINfailedCounter / RingCount) * 100)
        GSMINnoCarrierpercent.Text = String.Format("{0:n2}", (NoCarrierCount / RingCount) * 100)
        GSMINwrongBaudpercent.Text = String.Format("{0:n2}", (ConnectedWrongBaudCount / RingCount) * 100)
        GSMINParityErrorCount.Text = "" & GSMINParityErrorCounter
    End Sub
    Private Sub CountINPSTN()

        Dim FileReader As StreamReader
        Dim strLine As String
        Dim addressposition As Integer
        Dim exoaddress As String = "XXXX"
        Dim PLAH As String = "PLA"
        Dim ELAH As String = "ELA"

        Dim RingCount As Integer
        RingCount = 0
        PSTNINRingCount.Text = RingCount

        Dim ConnectBAUDCount As Integer
        ConnectBAUDCount = 0
        PSTNINConnectBAUDCount.Text = ConnectBAUDCount

        Dim PSTNINfailedCounter As Integer
        PSTNINfailedCounter = 0
        PSTNINfailedCount.Text = PSTNINfailedCounter

        Dim PSTNINParityErrorCounter As Integer
        PSTNINParityErrorCounter = 0
        PSTNINParityErrorCount.Text = PSTNINParityErrorCounter


        Dim PSTNINmultiRingCounter As Integer
        PSTNINmultiRingCounter = 0
        PSTNINmultiRingCount.Text = PSTNINmultiRingCounter

        Dim PSTNRingnotreadyCounter As Integer
        PSTNRingnotreadyCounter = 0
        PSTNRingnotreadyCount.Text = PSTNRingnotreadyCounter

        Dim ConnectedWrongBaudCount As Integer
        ConnectedWrongBaudCount = 0
        PSTNINConnectedWrongBaudCount.Text = ConnectedWrongBaudCount

        Dim AnsweringCount As Integer
        AnsweringCount = 0
        PSTNINAnsweringCount.Text = AnsweringCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        PSTNINNoCarrierCount.Text = NoCarrierCount

        Dim ParityErrorthiscallCount As Integer
        ParityErrorthiscallCount = 0

        Dim GoodTermCount As Integer
        GoodTermCount = 0
        PSTNINConnectBAUDCount.Text = GoodTermCount

        Dim GoodTermAlarmsCount As Integer
        GoodTermAlarmsCount = 0
        PSTNINGoodTermalarmsCount.Text = GoodTermAlarmsCount

        Dim GoodTermtimesyncCount As Integer
        GoodTermtimesyncCount = 0
        PSTNINGoodTermtimesyncCount.Text = GoodTermtimesyncCount

        Dim PSTNINnoaddressCounter As Integer = 0

        Dim BadLogsCount As Integer
        BadLogsCount = 0
        '  PSTNINBadLogsTermCount.Text = BadLogsCount

        Dim ModemIsIdle As Boolean
        Dim ModemIsConfiged As Boolean
        Dim ModemIsRinging As Boolean
        Dim ModemIsAnswering As Boolean
        Dim ModemIsConnectedCorrectBaud As Boolean
        Dim ModemIsConnectedWRONGBaud As Boolean
        Dim ModemIsNoCarrier As Boolean
        Dim ModemIsCallActive As Boolean
        Dim ModemIsGotAddress As Boolean
        Dim ModemIsGoodCallEnd As Boolean
        Dim ModemIsCallTerminated As Boolean
        Dim ModemIsGoodtimesync = False
        Dim ModemIsBadLogs = False
        Dim ModemIsAlarmRecovery = False
        Dim ThisCallHasManyParityerrors As Boolean



        Dim CallTerminaledCount As Integer
        CallTerminaledCount = 0


        Dim timeofring As String = "00:00:00"
        Dim timeofconnect As String = "00:00:00"
        Dim timeofdisconnect As String = "00:00:00"
        Dim timeofnoCarrier As String = "00:00:00"
        Dim timeofpreviouscleardown As String = "00:00:00"
        Dim CommPortDetails As String = ""
        Dim logedmessage As String = "nothing to report"
        Dim timeofcleardown As String = ""
        Dim realtimeofcleardown As Date
        Dim realtimeoftimeofring As Date

        Dim FirstCharacter As Integer
        Dim logedmessagetimesync As String = "nothing to report"
        Dim logedmessagelogs As String = "nothing to report"
        Dim logedmessagealarms As String = "nothing to report"

        ''''
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        ''start new workbook
        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        ' Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        ' '''' leave to workbook open to speed things up

        For Each filename As String In PSTNINFilestoAnalyseListbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data

            ModemIsIdle = False
            ModemIsConfiged = False
            ModemIsRinging = False
            ModemIsAnswering = False
            ModemIsConnectedCorrectBaud = False
            ModemIsConnectedWRONGBaud = False
            ModemIsNoCarrier = False
            ModemIsCallActive = False
            ModemIsGotAddress = False
            ModemIsGoodCallEnd = False
            ModemIsCallTerminated = False

            ThisCallHasManyParityerrors = False
            exoaddress = "XXXX"


            Do While Not strLine Is Nothing




                Try

                    Dim TestPosition As Integer = InStr(1, filename, "\EXORDI\", CompareMethod.Text) + 8 'add on 7 for the test string length too
                    CommPortDetails = Mid(filename, TestPosition, 11)





                    'loop through messages and see where we are within the call

callprogress:' read message to see where we are

                    strLine = FileReader.ReadLine


                    Select Case True
                        Case strLine.Contains("+++")
                            ModemIsCallTerminated = True
                            ModemIsConfiged = False
                            ModemIsRinging = False
                            exoaddress = "XXXX"
                            ModemIsAlarmRecovery = False

                        Case strLine.Contains("AUTO-ANSWER (state 5)")
                            ModemIsCallTerminated = False
                            ModemIsConfiged = True

                            AnsweringCount = AnsweringCount + 1
                            PSTNINAnsweringCount.Text = AnsweringCount

                            If CheckBoxidletime.Checked Then
                                Dim busytime As System.TimeSpan
                                busytime = realtimeofcleardown - realtimeoftimeofring
                                Dim totalbusytime As System.TimeSpan
                                totalbusytime = totalbusytime + busytime
                                ''''

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")

                                Dim col As Integer = 20
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Modem ready"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Received at " & timeofcleardown
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                col = 21

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "busyduration"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = timeofcleardown
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = busytime.ToString
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = totalbusytime.ToString

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                                '       populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 21, "busyduration", CommPortDetails, timeofring, timeofcleardown, busytime.ToString)
                            End If




                        Case strLine.Contains("NG<CR><LF>")
                            timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                            If ModemIsRinging = True Then
                                ' alredy ringing
                                'may ring >1 now so ok
                                ' timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                                PSTNINmultiRingCounter = PSTNINmultiRingCounter + 1
                                PSTNINmultiRingCount.Text = PSTNINmultiRingCounter
                            Else
                                'new ring
                                ModemIsRinging = True
                                ' timeofring = Microsoft.VisualBasic.Left(strLine, 12)
                            End If


                            If ModemIsRinging = True And ModemIsConfiged = False Then
                                'modem not configured yet so may fail
                                PSTNRingnotreadyCounter = PSTNRingnotreadyCounter + 1
                                PSTNRingnotreadyCount.Text = PSTNRingnotreadyCounter
                                If CheckBoxLogIn.Checked And timeofring <> "" Then

                                    '   add Data
                                    oSheet = oBook.Worksheets("EXO")
                                    Dim col As Integer = 20

                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring without modem ready"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "Ring without modem ready"
                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                End If
                            End If

                        Case strLine.Contains("ANSWERING (state 9)")
                            ModemIsAnswering = True
                            RingCount = RingCount + 1
                            PSTNINRingCount.Text = RingCount

                        Case strLine.Contains("NO CARRIER<CR><LF>")
                            ModemIsNoCarrier = True
                            ModemIsRinging = False
                            If CheckBoxLogIn.Checked And timeofring <> "" Then
                                FirstCharacter = strLine.IndexOf("NO CARRIER<CR><LF>")
                                timeofnoCarrier = Microsoft.VisualBasic.Left(strLine, 12)


                                ' populatexcelsheetdetails(populatexcelsheetdetailsRowCounter, 20, "No Carrier", CommPortDetails, "", "Received at ", timeofring)
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Carrier"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "No Carrier"

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If
                            NoCarrierCount = NoCarrierCount + 1
                            PSTNINNoCarrierCount.Text = NoCarrierCount
                            'fail counter
                            PSTNINfailedCounter = PSTNINfailedCounter + 1
                            PSTNINfailedCount.Text = PSTNINfailedCounter


                        Case strLine.Contains("Rx: CONNECT 1200<CR><LF>")
                            ModemIsConnectedWRONGBaud = True
                            ModemIsConnectedCorrectBaud = False
                            ModemIsRinging = False
                            PSTNINfailedCounter = PSTNINfailedCounter + 1
                            ConnectedWrongBaudCount = ConnectedWrongBaudCount + 1
                            PSTNINConnectedWrongBaudCount.Text = ConnectedWrongBaudCount
                            PSTNINfailedCount.Text = PSTNINfailedCounter

                            If CheckBoxLogIn.Checked And timeofring <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Wrong Baud"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = ""
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "From Ring Received at " & timeofring

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If

                        Case strLine.Contains("Rx: CONNECT 2400<CR><LF>")
                            ModemIsConnectedCorrectBaud = True

                            'new call so reset the flags for sucesss
                            ModemIsGoodCallEnd = False 'reset the check for a call as this is a new one
                            ModemIsConnectedWRONGBaud = False
                            ModemIsRinging = False
                            ModemIsGotAddress = False
                            ModemIsGoodCallEnd = False

                            ParityErrorthiscallCount = 0
                            ThisCallHasManyParityerrors = False

                            timeofconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            ConnectBAUDCount = ConnectBAUDCount + 1
                            PSTNINConnectBAUDCount.Text = ConnectBAUDCount


                        Case strLine.Contains("Parity")

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            End If

                            If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount >= 3 Then
                                ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                                ThisCallHasManyParityerrors = True

                            End If

                        Case strLine.Contains("10F100")
                            ModemIsGotAddress = True
                            addressposition = InStr(1, strLine, "10F100")
                            'get PLA and ELS in Hex string values
                            PLAH = Mid(strLine, addressposition - 4, 2)
                            ELAH = Mid(strLine, addressposition - 2, 2)
                            'convert to decimal
                            Dim PLA = CInt("&H" & PLAH)
                            Dim ELA = CInt("&H" & ELAH)
                            exoaddress = PLA & ":" & ELA & " (Hex-" & PLAH & ELAH & ")"

                            'was "A401A401A401A401"
                            '102900                             'Alarm download complete for module
                        Case strLine.Contains("102900")
                            'alarm recovery 
                            ModemIsAlarmRecovery = True'

                        Case strLine.Contains("A4A43E") ' 
                            'log recovery failed
                            ModemIsBadLogs = True

                        Case strLine.Contains("0F100")
                            If ModemIsGotAddress = True Then
                                'time sync messages
                                ModemIsGoodtimesync = True
                                GoodTermCount = GoodTermCount + 1
                                PSTNINGoodTermtimesyncCount.Text = GoodTermCount
                            End If


                        Case strLine.Contains("DISCONNECTING (state 8)")


                            ModemIsCallTerminated = True

                            If ModemIsGotAddress = False Then
                                PSTNINnoaddressCounter = PSTNINnoaddressCounter + 1
                                PSTNINfailedCounter = PSTNINfailedCounter + 1
                            End If

                            timeofdisconnect = Microsoft.VisualBasic.Left(strLine, 12)
                            If ThisCallHasManyParityerrors = True Then
                                PSTNINParityErrorCounter = PSTNINParityErrorCounter + 1
                            End If

                            If ModemIsGoodtimesync = True Then 'And ModemIsBadLogs Then ' = False And ModemIsAlarmRecovery = True Then
                                logedmessagetimesync = "Call Termination (with time sync)"
                                GoodTermtimesyncCount = GoodTermtimesyncCount + 1
                                PSTNINGoodTermtimesyncCount.Text = GoodTermtimesyncCount
                            Else
                                logedmessagetimesync = "Call Termination (no time sync)"
                            End If

                            'no Logs on Incomming
                            'If ModemIsBadLogs = True Then
                            '    logedmessagelogs = "Bad Logs"
                            '    BadLogsCount = BadLogsCount + 1
                            '    '      PSTNINBadLogsTermCount.Text = BadLogsCount
                            'Else
                            '    logedmessagelogs = "Good Logs"
                            'End If

                            If ModemIsAlarmRecovery = True Then
                                logedmessagealarms = "Alarms Recovered"
                                GoodTermAlarmsCount = GoodTermAlarmsCount + 1
                                PSTNINGoodTermalarmsCount.Text = GoodTermAlarmsCount
                            Else
                                logedmessagealarms = "No Alarms Recovered"
                            End If



                            If CheckBoxLogIn.Checked And timeofconnect <> "" Then

                                '   add Data
                                oSheet = oBook.Worksheets("EXO")
                                Dim col As Integer = 20

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated" & " from ring recieved at " & timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                If ParityErrorthiscallCount > 0 Then
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors " & ParityErrorthiscallCount
                                Else
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Parity errors "
                                End If
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & exoaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = logedmessagetimesync
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 6).Value = logedmessagelogs
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 7).Value = logedmessagealarms

                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                            End If




                    End Select



                Catch e As System.NullReferenceException
                Catch e As FormatException
                Catch e As OverflowException
                Catch e As System.InvalidCastException 'miss interprited string to dates
                    '  strLine = FileReader.ReadLine 'move on
                End Try



            Loop


            FileReader.Close()

        Next
        ''''
        ' save and exit
        oBook.Save()
        oBook.Application.DisplayAlerts = False
        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()


        ''''


        PSTNINfailedCount.Text = PSTNINfailedCounter
        PSTNINConnectpercent.Text = String.Format("{0:n2}", (ConnectBAUDCount / RingCount) * 100)
        PSTNINFailpercent.Text = String.Format("{0:n2}", (PSTNINfailedCounter / RingCount) * 100)
        PSTNINnoCarrierpercent.Text = String.Format("{0:n2}", (NoCarrierCount / RingCount) * 100)
        PSTNINwrongBaudpercent.Text = String.Format("{0:n2}", (ConnectedWrongBaudCount / RingCount) * 100)
        PSTNINParityErrorCount.Text = "" & PSTNINParityErrorCounter
    End Sub







    Private Sub GetLogfilestoProcess() '(Commstype As String)
        Dim DateoffilestoAnalyse As String
        Dim LogfilesFolder As String
        DateoffilestoAnalyse = DateTimePickerSplitter.Value.Date.ToString("yyyyMMdd")

        'start by flushing out old list and recheck all tick boxes
        PSTNOUTFilestoAnalyseListbox.Items.Clear()
        PSTNINFilestoAnalyseListbox.Items.Clear()
        GSMOUTFilestoAnalyseListbox.Items.Clear()
        GSMINFilestoAnalyseListbox.Items.Clear()

        'com1 PSTN

        If CheckBoxPSTNCom01IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom01OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If


        'com2
        If CheckBoxPSTNCom02IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom02OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If

        'com3

        If CheckBoxPSTNCom03IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom03IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom03OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If


        'com4

        If CheckBoxPSTNCom04IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom04IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
        End If

        If CheckBoxPSTNCom04OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If

        'com5

        If CheckBoxPSTNCom05IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom05OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If


        'com6
        If CheckBoxPSTNCom06IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom06OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If



        'com7

        If CheckBoxPSTNCom07IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If

        If CheckBoxPSTNCom07IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com07", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
        End If

        If CheckBoxPSTNCom07OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If


        'com8

        If CheckBoxPSTNCom08IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
        End If
        If CheckBoxPSTNCom08IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
        End If

        If CheckBoxPSTNCom08OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
        End If

        'GSM
        'com1

        If CheckBoxGSMCom01OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
        End If
        If CheckBoxGSMCom01IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM23", "GSM", "IN")
        End If


        'com2
        If CheckBoxGSMCom02IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM23", "GSM", "IN")
        End If
        If CheckBoxGSMCom02OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
        End If

        'com3

        If CheckBoxGSMCom03IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM25", "GSM", "IN")
        End If


        'com4

        If CheckBoxGSMCom04IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM25", "GSM", "IN")
        End If

        'com5

        If CheckBoxGSMCom05OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
        End If
        If CheckBoxGSMCom05IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM05", DateoffilestoAnalyse, "COM23", "GSM", "IN")
        End If

        'com6
        If CheckBoxGSMCom06IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM23", "GSM", "IN")
        End If
        If CheckBoxGSMCom06OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
        End If


        'com7

        If CheckBoxGSMCom07IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM25", "GSM", "IN")
        End If

        'com8

        If CheckBoxGSMCom08IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM25", "GSM", "IN")
        End If


    End Sub









    'Private Sub GetPSTNLogs_Click(sender As Object, e As EventArgs)
    '    GetLogfilestoProcess("PSTN")
    'End Sub

    'Private Sub GetGSMLogs_Click(sender As Object, e As EventArgs)
    '    GetLogfilestoProcess("GSM")
    'End Sub








    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        DLForm.Show()
    End Sub



    Private Sub splitalllogsofdate_Click(sender As Object, e As EventArgs) Handles splitalllogsofdate.Click
        SplitAllLogFiles()
    End Sub


    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles EnableDebug.Click
        GSMINFilestoAnalyseListbox.Visible = True
        PSTNINFilestoAnalyseListbox.Visible = True
        GSMOUTFilestoAnalyseListbox.Visible = True
        PSTNOUTFilestoAnalyseListbox.Visible = True
        FilestosplitListbox.Visible = True
        MFilestoAnalyseListbox.Visible = True
        CheckBoxMMessages.Visible = True

    End Sub

    Private Sub DisableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisableToolStripMenuItem.Click
        GSMINFilestoAnalyseListbox.Visible = False
        PSTNINFilestoAnalyseListbox.Visible = False
        GSMOUTFilestoAnalyseListbox.Visible = False
        PSTNOUTFilestoAnalyseListbox.Visible = False
        FilestosplitListbox.Visible = False

    End Sub


    Private Sub CalculateTotals_Click(sender As Object, e As EventArgs) Handles CalculateTotals.Click
        GetLogfilestoProcess()
        CountOutPSTN()
        CountINPSTN()
        CountOutGSM()
        CountINGSM()
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        helpabout.Show()

    End Sub

    Private Sub FilestosplitListbox_DoubleClick(sender As Object, e As EventArgs) Handles FilestosplitListbox.DoubleClick
        System.Diagnostics.Process.Start(FilestosplitListbox.SelectedItem.ToString())

    End Sub
    Private Sub GSMINFilestoAnalyseListbox_DoubleClick(sender As Object, e As EventArgs) Handles GSMINFilestoAnalyseListbox.DoubleClick
        System.Diagnostics.Process.Start(GSMINFilestoAnalyseListbox.SelectedItem.ToString())

    End Sub
    Private Sub GSMOUTilestoAnalyseListbox_DoubleClick(sender As Object, e As EventArgs) Handles GSMOUTFilestoAnalyseListbox.DoubleClick
        System.Diagnostics.Process.Start(GSMOUTFilestoAnalyseListbox.SelectedItem.ToString())

    End Sub
    Private Sub PSTNINFilestoAnalyseListbox_DoubleClick(sender As Object, e As EventArgs) Handles PSTNINFilestoAnalyseListbox.DoubleClick
        System.Diagnostics.Process.Start(PSTNINFilestoAnalyseListbox.SelectedItem.ToString())

    End Sub
    Private Sub PSTNOUTFilestoAnalyseListbox_DoubleClick(sender As Object, e As EventArgs) Handles PSTNOUTFilestoAnalyseListbox.DoubleClick
        System.Diagnostics.Process.Start(PSTNOUTFilestoAnalyseListbox.SelectedItem.ToString())

    End Sub
    Private Sub MFilestoAnalyseListbox_DoubleClick(sender As Object, e As EventArgs) Handles MFilestoAnalyseListbox.DoubleClick
        System.Diagnostics.Process.Start(MFilestoAnalyseListbox.SelectedItem.ToString())

    End Sub

    Private Sub tickallforallsave(onoroff As Boolean)
        'CheckBoxPSTNCom01IN.Checked = True
        CheckBoxPSTNCom01IN.Checked = onoroff
        CheckBoxPSTNCom01OUT.Checked = onoroff
        CheckBoxPSTNCom02IN.Checked = onoroff
        CheckBoxPSTNCom02OUT.Checked = onoroff
        CheckBoxPSTNCom03IN.Checked = onoroff
        CheckBoxPSTNCom03OUT.Checked = onoroff
        CheckBoxPSTNCom03IN2.Checked = onoroff
        CheckBoxPSTNCom04IN.Checked = onoroff
        CheckBoxPSTNCom04OUT.Checked = onoroff
        CheckBoxPSTNCom04IN2.Checked = onoroff
        CheckBoxPSTNCom05IN.Checked = onoroff
        CheckBoxPSTNCom05OUT.Checked = onoroff
        CheckBoxPSTNCom06IN.Checked = onoroff
        CheckBoxPSTNCom06OUT.Checked = onoroff
        CheckBoxPSTNCom07IN.Checked = onoroff
        CheckBoxPSTNCom07OUT.Checked = onoroff
        CheckBoxPSTNCom07IN2.Checked = onoroff
        CheckBoxPSTNCom08IN.Checked = onoroff
        CheckBoxPSTNCom08OUT.Checked = onoroff
        CheckBoxPSTNCom08IN2.Checked = onoroff

        CheckBoxGSMCom01IN.Checked = onoroff
        CheckBoxGSMCom01OUT.Checked = onoroff
        CheckBoxGSMCom02IN.Checked = onoroff
        CheckBoxGSMCom02OUT.Checked = onoroff
        CheckBoxGSMCom03IN.Checked = onoroff
        ' CheckBoxGSMCom03OUT.Checked = onoroff
        CheckBoxGSMCom04IN.Checked = onoroff
        'CheckBoxGSMCom04OUT.Checked = onoroff
        CheckBoxGSMCom05IN.Checked = onoroff
        CheckBoxGSMCom05OUT.Checked = onoroff
        CheckBoxGSMCom06IN.Checked = onoroff
        CheckBoxGSMCom06OUT.Checked = onoroff
        CheckBoxGSMCom07IN.Checked = onoroff
        '  CheckBoxGSMCom07OUT.Checked = onoroff
        CheckBoxGSMCom08IN.Checked = onoroff
        '  CheckBoxGSMCom08OUT.Checked = onoroff

    End Sub


    Private Sub RunAnalysisforallsave()
        Dim DateoffilestoAnalyse As String
        Dim LogfilesFolder As String
        DateoffilestoAnalyse = DateTimePickerSplitter.Value.Date.ToString("yyyyMMdd")
        On Error GoTo Errornofilefound
        PleaseWait.Show()


        'start by flushing out old list and recheck aall tick boxes
        PSTNOUTFilestoAnalyseListbox.Items.Clear()
        PSTNINFilestoAnalyseListbox.Items.Clear()

        tickallforallsave(True)
        RunAnalysisforeachsave()
        PleaseWait.Close()

        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry ??? file open??")
        PleaseWait.Close()


    End Sub


    Private Sub RunAnalysisforeachsave()
        Dim DateoffilestoAnalyse As String
        Dim LogfilesFolder As String
        DateoffilestoAnalyse = DateTimePickerSplitter.Value.Date.ToString("yyyyMMdd")
        On Error GoTo Errornofilefound
        PleaseWait.Show()

        'start by flushing out old list and recheck aall tick boxes
        ' PSTNOUTFilestoAnalyseListbox.Items.Clear()
        ' PSTNINFilestoAnalyseListbox.Items.Clear()

        'com1

        If CheckBoxPSTNCom01IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(3, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)
            ClearForm()
        End If


        If CheckBoxPSTNCom01OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(15, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If





        'com2
        If CheckBoxPSTNCom02IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()

            populatexcelsheet(4, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)
            ClearForm()


        End If

        If CheckBoxPSTNCom02OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(16, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()

        End If
        'com3

        If CheckBoxPSTNCom03IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()

            populatexcelsheet(6, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        If CheckBoxPSTNCom03IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(5, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If


        If CheckBoxPSTNCom03OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(17, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()

        End If
        'com4

        If CheckBoxPSTNCom04IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(8, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        If CheckBoxPSTNCom04IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(7, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)
            ClearForm()
        End If

        If CheckBoxPSTNCom04OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(18, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If
        'com5

        If CheckBoxPSTNCom05IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(9, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If


        If CheckBoxPSTNCom05OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(19, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If

        'com6
        If CheckBoxPSTNCom06IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(10, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If


        If CheckBoxPSTNCom06OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(20, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If


        'com7

        If CheckBoxPSTNCom07IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(12, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        If CheckBoxPSTNCom07IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com07", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(11, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        If CheckBoxPSTNCom07OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(21, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If

        'com8

        If CheckBoxPSTNCom08IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM26", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(14, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If


        '
        If CheckBoxPSTNCom08IN2.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM23", "PSTN", "IN")
            CountINPSTN()
            populatexcelsheet(13, 5, PSTNINRingCount.Text,
              PSTNINConnectBAUDCount.Text,
              PSTNINfailedCount.Text,
              PSTNINConnectpercent.Text,
              PSTNINFailpercent.Text,
              PSTNINConnectedWrongBaudCount.Text,
              PSTNINNoCarrierCount.Text,
              PSTNINAnsweringCount.Text,
              PSTNINnoCarrierpercent.Text,
              PSTNINParityErrorCount.Text, PSTNRingnotreadyCount.Text, PSTNINGoodTermalarmsCount.Text, PSTNINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        If CheckBoxPSTNCom08OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM24", "PSTN", "OUT")
            CountOutPSTN()
            populatexcelsheet(22, 5, PSTNDialCount.Text,
    PSTNDisconnectCount.Text,
    PSTNOUTFailedCount.Text,
    PSTNOUTGoodpercent.Text,
    PSTNOUTFailpercent.Text,
    PSTNNoReplyCount.Text,
    PSTNNoCarrierCount.Text,
    PSTNDisconnectCount.Text,
    PSTNBusyCount.Text,
    PSTNParityErrorCount.Text,
    PSTNBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If

        '**********GSM************

        'com1
        If CheckBoxGSMCom01IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM23", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(23, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        If CheckBoxGSMCom01OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM01", DateoffilestoAnalyse, "COM25", "GSM", "OUT")

            CountOutGSM()
            populatexcelsheet(33, 5, GSMDialCount.Text,
    GSMDisconnectCount.Text,
    GSMOUTFailedCount.Text,
    GSMOUTGoodpercent.Text,
    GSMOUTFailpercent.Text,
    GSMNoReplyCount.Text,
    GSMNoCarrierCount.Text,
    GSMDisconnectCount.Text,
    GSMBusyCount.Text,
    GSMParityErrorCount.Text,
    GSMBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()

        End If
        'com2
        If CheckBoxGSMCom02IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM23", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(24, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        If CheckBoxGSMCom02OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM02", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
            CountOutGSM()
            populatexcelsheet(34, 5, GSMDialCount.Text,
    GSMDisconnectCount.Text,
    GSMOUTFailedCount.Text,
    GSMOUTGoodpercent.Text,
    GSMOUTFailpercent.Text,
    GSMNoReplyCount.Text,
    GSMNoCarrierCount.Text,
    GSMDisconnectCount.Text,
    GSMBusyCount.Text,
    GSMParityErrorCount.Text,
    GSMBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If
        'com3

        If CheckBoxGSMCom03IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com03", DateoffilestoAnalyse, "COM25", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(25, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If

        'com4

        If CheckBoxGSMCom04IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com04", DateoffilestoAnalyse, "COM25", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(26, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        'com5
        If CheckBoxGSMCom05IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM05", DateoffilestoAnalyse, "COM23", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(27, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        If CheckBoxGSMCom05OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com05", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
            CountOutGSM()
            populatexcelsheet(35, 5, GSMDialCount.Text,
    GSMDisconnectCount.Text,
    GSMOUTFailedCount.Text,
    GSMOUTGoodpercent.Text,
    GSMOUTFailpercent.Text,
    GSMNoReplyCount.Text,
    GSMNoCarrierCount.Text,
    GSMDisconnectCount.Text,
    GSMBusyCount.Text,
    GSMParityErrorCount.Text,
    GSMBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()

        End If

        'com6
        If CheckBoxGSMCom06IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM23", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(28, 5, GSMINRingCount.Text,
                  GSMINConnectBAUDCount.Text,
                  GSMINnoConnectCount.Text,
                  GSMINConnectpercent.Text,
                  GSMINFailpercent.Text,
                  GSMINConnectedWrongBaudCount.Text,
                  GSMINNoCarrierCount.Text,
                  GSMINAnsweringCount.Text,
                  GSMINnoCarrierpercent.Text,
                  GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        If CheckBoxGSMCom06OUT.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\Com06", DateoffilestoAnalyse, "COM25", "GSM", "OUT")
            CountOutGSM()
            populatexcelsheet(36, 5, GSMDialCount.Text,
    GSMDisconnectCount.Text,
    GSMOUTFailedCount.Text,
    GSMOUTGoodpercent.Text,
    GSMOUTFailpercent.Text,
    GSMNoReplyCount.Text,
    GSMNoCarrierCount.Text,
    GSMDisconnectCount.Text,
    GSMBusyCount.Text,
    GSMParityErrorCount.Text,
    GSMBlackListed.Text, PSTNOUTGoodTermalarmsCount.Text, PSTNOUTGoodTermtimesyncCount.Text)
            ClearForm()
        End If



        'com7

        If CheckBoxGSMCom07IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM07", DateoffilestoAnalyse, "COM25", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(29, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If
        'com8

        If CheckBoxGSMCom08IN.Checked Then
            LogfilesFolder = SelectedFolder.Text
            GetLogFilesofDateCommport(LogfilesFolder & "\COM08", DateoffilestoAnalyse, "COM25", "GSM", "IN")
            CountINGSM()
            populatexcelsheet(30, 5, GSMINRingCount.Text,
              GSMINConnectBAUDCount.Text,
              GSMINnoConnectCount.Text,
              GSMINConnectpercent.Text,
              GSMINFailpercent.Text,
              GSMINConnectedWrongBaudCount.Text,
              GSMINNoCarrierCount.Text,
              GSMINAnsweringCount.Text,
              GSMINnoCarrierpercent.Text,
              GSMINParityErrorCount.Text, GSMRingnotreadyCount.Text, GSMINGoodTermalarmsCount.Text, GSMINGoodTermtimesyncCount.Text)

            ClearForm()
        End If




        ClearForm()

        PleaseWait.Close()

        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry ??? file open??")
        PleaseWait.Close()

    End Sub





    Private Sub populatexcelsheet(Row As Integer, Col As Integer, Data As String, Data2 As String, Data3 As String, Data4 As String, Data5 As String, Data6 As String, Data7 As String, Data8 As String, Data9 As String, Data10 As String, Data11 As String, Data12 As String, Data13 As String)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        'Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        'datestamp
        oSheet = oBook.Worksheets("Summary")
        oSheet.cells(3, 1).Value = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")

        'add data
        oSheet = oBook.Worksheets("EXO")
        oSheet.cells(Row, Col).Value = Data
        oSheet.cells(Row, Col + 1).Value = Data2
        oSheet.cells(Row, Col + 2).Value = Data3
        oSheet.cells(Row, Col + 3).Value = Data4
        oSheet.cells(Row, Col + 4).Value = Data5
        oSheet.cells(Row, Col + 5).Value = Data6
        oSheet.cells(Row, Col + 6).Value = Data7
        oSheet.cells(Row, Col + 7).Value = Data8
        oSheet.cells(Row, Col + 8).Value = Data9
        oSheet.cells(Row, Col + 9).Value = Data10
        oSheet.cells(Row, Col + 10).Value = Data11
        oSheet.cells(Row, Col + 11).Value = Data12
        oSheet.cells(Row, Col + 12).Value = Data13
        ' save and exit
        oBook.Save()



        oBook.Application.DisplayAlerts = False
        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
    End Sub
    Private Sub populatexcelsheetdetails(ByRef populatexcelsheetdetailsRowCounter As Integer, Col As Integer, Data As String, Data2 As String, Data3 As String, Data4 As String, Data5 As String)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        ' Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        'datestamp the results
        oSheet = oBook.Worksheets("Summary")
        oSheet.cells(3, 1).Value = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")

        '   add Data
        oSheet = oBook.Worksheets("EXO")
        oSheet.cells(populatexcelsheetdetailsRowCounter, Col).Value = Data
        oSheet.cells(populatexcelsheetdetailsRowCounter, Col + 1).Value = Data2
        oSheet.cells(populatexcelsheetdetailsRowCounter, Col + 2).Value = Data3
        oSheet.cells(populatexcelsheetdetailsRowCounter, Col + 3).Value = Data4
        oSheet.cells(populatexcelsheetdetailsRowCounter, Col + 4).Value = Data5

        populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1



        ' save and exit
        oBook.Save()

        oBook.Application.DisplayAlerts = False


        Dim Dateinfile As String = DateTimePickerSplitter.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
    End Sub

    Private Sub makeblankexceltemplate()
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")

        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI - Blank.xls")


        ' save and exit
        oBook.Application.DisplayAlerts = False
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI.xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
    End Sub

    Private Sub save_Click(sender As Object, e As EventArgs) Handles save.Click
        SplitAllLogFiles()
        RunAnalysisforeachsave()
    End Sub

    Private Sub OpenResultsFile_Click(sender As Object, e As EventArgs) Handles OpenResultsFile.Click
        Dim networkpathlength As String = SelectedFolder.Text.Length - 7 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolder.Text.Substring(0, networkpathlength)
        'Dim partialpath As String = SelectedFolder.Text.Substring(0, 15)

        System.Diagnostics.Process.Start(partialpath & "\OE Comms Channel Report_RDI.xls")
    End Sub




    'Private Sub DateTimePickertimestart_ValueChanged(sender As Object, e As EventArgs)
    '    Dim StartHour As String
    '    Dim StartMinute As String
    '    StartHour = DateTimePickertimestart.Value.Hour
    '    StartMinute = DateTimePickertimestart.Value.Minute

    'End Sub



    Private Sub PABXToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PABXToolStripMenuItem.Click
        FormPABX.Show()
    End Sub



    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem1.Click

    End Sub

    Private Sub TransparentToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub TransparentToolStrip_ItemClicked_1(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ' SaveTickedStuff()
    End Sub

    Private Sub AlldoandSave_Click(sender As Object, e As EventArgs) Handles Alldoandsave.Click
        SplitAllLogFiles()
        RunAnalysisforallsave()
    End Sub




    Private Sub ToolStripMenuItem3_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        CheckBoxLogIn.Visible = True
        CheckBoxidletime.Visible = True
        CheckboxRing.Visible = True
        LabelIN.Visible = True
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        CheckBoxLogIn.Visible = False
        CheckBoxidletime.Visible = False
        CheckboxRing.Visible = False
        LabelIN.Visible = False
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        CheckBoxLogOut.Visible = True
        ' CheckBoxLogGoods.Visible = True
        CheckBoxoutphonenumber.Visible = True
        TextBoxphonenumber.Visible = True
        LabelOUT.Visible = True
        PSTNProdtimeResultCount.Visible = True
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        CheckBoxLogOut.Visible = False
        '  CheckBoxLogGoods.Visible = False
        CheckBoxoutphonenumber.Visible = False
        TextBoxphonenumber.Visible = False
        LabelOUT.Visible = False
        PSTNProdtimeResultCount.Visible = False
    End Sub


    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles GSMDialCount.TextChanged

    End Sub

    Private Sub GroupBox7_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label45_Click(sender As Object, e As EventArgs) Handles Label45.Click

    End Sub

    Private Sub PSTNINFilestoAnalyseListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PSTNINFilestoAnalyseListbox.SelectedIndexChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub CheckboxRing_CheckedChanged(sender As Object, e As EventArgs) Handles CheckboxRing.CheckedChanged
        If CheckboxRing.Checked Then
            PSTNINmultiRingCount.Visible = True
            GSMINmultiRingCount.Visible = True
            PSTNmultiringlabel.Visible = True
            GSMmultiringlabel.Visible = True
        Else
            PSTNINmultiRingCount.Visible = False
            GSMINmultiRingCount.Visible = False
            PSTNmultiringlabel.Visible = False
            GSMmultiringlabel.Visible = False
        End If

    End Sub

    Private Sub ExtraINlogging_Click(sender As Object, e As EventArgs)
        CheckBoxLogIn.Visible = True
        CheckBoxidletime.Visible = True
        CheckboxRing.Visible = True
        LabelIN.Visible = True
    End Sub

    Private Sub ToolStripMenuPABXLogs_Click(sender As Object, e As EventArgs) Handles ToolStripMenuPABXLogs.Click
        Form4PABX.Show()
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemLD.Click
        Form4.Show()
    End Sub

    Private Sub MFilestoAnalyseListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MFilestoAnalyseListbox.SelectedIndexChanged

    End Sub

    Private Sub CheckBoxLogTimeouts_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLogOut.CheckedChanged

    End Sub

    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Label58.Click

    End Sub

    Private Sub PSTNOUTFailpercent_TextChanged(sender As Object, e As EventArgs) Handles PSTNOUTFailpercent.TextChanged

    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click

    End Sub

    Private Sub GSMOUTGoodpercent_TextChanged(sender As Object, e As EventArgs) Handles GSMOUTGoodpercent.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles GSMOUTGoodTermalarmsCount.TextChanged

    End Sub
End Class
