Imports System.IO
Imports System.Data.Odbc

Public Class DLForm

    Private Sub btnSelectFolderDL_Click(sender As Object, e As EventArgs) Handles btnSelectFolderDL.Click
        Dim Folder As String
        Folder = SelectFolderdl()
    End Sub


    Private Function SelectFolderdl() As String
        Dim DLdefaultfolder = "c:\RDI Analyser\DLRDI"
        Dim FilesFolder As String

        FolderBrowserDialog1.ShowDialog()
        FilesFolder = FolderBrowserDialog1.SelectedPath
        If FilesFolder = "" Then
            FilesFolder = DLdefaultfolder
        End If
        SelectedFolderDL.Text = FilesFolder
        Return FilesFolder
    End Function

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ClearForm()
        SelectedFolderDL.Text = "c:\RDI Analyser\DLRDI"

    End Sub




    Private Sub GetDLLogFilesDLofDate(commscontroler As String, Port As String, listboxtosendto As ListBox)

        Dim Dateinfile As String
        Dim LogfilesFolder As String
        Dim FileReader As StreamReader
        Dim strLine As String
        Dim Filenametag As String
        ' Dim Port As String


        On Error GoTo Errornofilefound
        '      PleaseWait.Show()
        LogfilesFolder = SelectedFolderDL.Text

        Filenametag = "*.log"
        Dateinfile = DateTimePickerSplitterDL.Value.Date.ToString("ddd MMM dd")


        LogfilesFolder = SelectedFolderDL.Text & "\" & commscontroler
        Filenametag = "Port" & Port & "-" & "*.log"

        Dim fileNames = My.Computer.FileSystem.GetFiles(
          LogfilesFolder, FileIO.SearchOption.SearchTopLevelOnly, Filenametag)

        For Each fileName As String In fileNames

            FileReader = New StreamReader(fileName) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine
                If InStr(1, strLine, Dateinfile) Then    'string search 
                    ' Dlfilesofdatelistbox.Items.Add(fileName)
                    listboxtosendto.Items.Add(fileName)

                    GoTo foundafile

                End If

            Loop
foundafile:
            FileReader.Close()
        Next
        '      PleaseWait.Close()
        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry No file or Path Found." & LogfilesFolder)
        '      PleaseWait.Close()

    End Sub




    Private Sub CountINPSTN()
        Dim FileReader As StreamReader
        Dim strLine As String
        Dim Dateinfile As String
        '  Dim ModemConfigString As String
        Dim nextrdimessage As Integer


        Dim ModemConfigString2400 As String = "AT&FV1N0S0=1S7=60&W&W1"
        Dim ModemConfigString1200 As String = "ATV1E0&A0"
        Dim ModemConfigString300 As String = "AT&A2V1E0&W0"

        Dim RingCount As Integer
        RingCount = 0
        PSTNINRingCount.Text = RingCount

        Dim ConnectBAUDCount As Integer
        ConnectBAUDCount = 0
        PSTNINConnectBAUDCount.Text = ConnectBAUDCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        PSTNINNoCarrierCount.Text = NoCarrierCount

        Dim AlarmRequestCount As Integer
        AlarmRequestCount = 0
        PSTNINAlarmRequestCount.Text = AlarmRequestCount

        Dim ResetAfterConnectCount As Integer
        ResetAfterConnectCount = 0
        PSTNINResetAfterConnectCount.Text = ResetAfterConnectCount

        Dim NoSpeakorResetCount As Integer
        NoSpeakorResetCount = 0
        PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount

        For Each filename As String In pstninlistbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine

                Dateinfile = DateTimePickerSplitterDL.Value.Date.ToString("ddd MMM dd")
                If InStr(1, strLine, Dateinfile) Then ' only do this for the messages with the correct date

                    'ok in here we must have the correct date BUT not all messaged are time stamped so start tighr loop before we go and checj again

                    For DateStampScopex = 1 To 20 ' loop round after the incomming ring looking for what happens next
  
                        strLine = FileReader.ReadLine

     
                        Dim Searchfor As String = "RING"
                        If InStr(1, strLine, Searchfor) Then
                            RingCount = RingCount + 1
                            PSTNINRingCount.Text = RingCount

                            For x = 1 To 20 ' loop round after the incomming ring looking for what happens next
                                'should be ring
                                'connect
                                'T<A>
                                'replyb from rtu with string
                                '*workPackageComplete
                                '* Work Request DLxxxx:DL
ProcessIncomming:
                                strLine = FileReader.ReadLine

                                If InStr(1, strLine, "NO CARRIER") Then    'string search 
                                    NoCarrierCount = NoCarrierCount + 1
                                    PSTNINNoCarrierCount.Text = NoCarrierCount
                                    'gracefull exit nothing to see here
                                    GoTo foundresult
                                End If

                                If InStr(1, strLine, "CONNECT") Then    'string search 
                                    ConnectBAUDCount = ConnectBAUDCount + 1
                                    PSTNINConnectBAUDCount.Text = ConnectBAUDCount

                                    For nextrdimessage = 1 To 10 'loop round after the connect to chech for what happens

                                        ''just a check if we loop too long looking for someting that never happens
                                        strLine = FileReader.ReadLine
                                        If InStr(1, strLine, "RING") Then 'holy smoke batman getting a new call get out
                                            '           MessageBox.Show(" Crashed into next ring  previous call nothing after connect in 10 message loops")
                                            GoTo ProcessIncomming
                                        End If

                                        If InStr(1, strLine, "  R A<") Then 'responce from RTU 
                                            AlarmRequestCount = AlarmRequestCount + 1
                                            PSTNINAlarmRequestCount.Text = AlarmRequestCount
                                            GoTo foundresult
                                        End If

                                        If InStr(1, strLine, ModemConfigString300) Or InStr(1, strLine, ModemConfigString1200) Or InStr(1, strLine, ModemConfigString2400) Then 'responce from RTU 
                                            ResetAfterConnectCount = ResetAfterConnectCount + 1
                                            PSTNINResetAfterConnectCount.Text = ResetAfterConnectCount
                                            GoTo foundresult
                                        End If

                                        If nextrdimessage > 9 Then
                                            NoSpeakorResetCount = NoSpeakorResetCount + 1
                                            PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount

                                            'MessageBox.Show(" no alarm reply or reset found")
                                        End If
                                    Next nextrdimessage
                                End If
                            Next x
                        End If

                    Next DateStampScopex



                End If


foundresult:
            Loop

            FileReader.Close()

        Next

        Failure3PSTN.Text = (ConnectBAUDCount / RingCount) * 100
        Failure4PSTN.Text = (NoCarrierCount / RingCount) * 100
        Failure5PSTN.Text = (AlarmRequestCount / ConnectBAUDCount) * 100
    End Sub

    Private Sub CountOutPSTN()
        Dim FileReader As StreamReader
        Dim strLine As String
        Dim Dateinfile As String


        Dim DialCount As Integer
        DialCount = 0

        Dim DialCount300 As Integer
        DialCount300 = 0
        PSTNDialCount300.Text = DialCount300

        Dim WorkPackageNoDialResultCount As Integer
        WorkPackageNoDialResultCount = 0
        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount


        Dim DialCount1200 As Integer
        DialCount1200 = 0
        PSTNDialCount1200.Text = DialCount1200

        Dim DialCount2400 As Integer
        DialCount2400 = 0
        PSTNDialCount2400.Text = DialCount2400

        Dim ConnectCountAll As Integer
        ConnectCountAll = 0

        Dim ConnectCount300 As Integer
        ConnectCount300 = 0
        PSTNConnectCount300.Text = ConnectCount300

        Dim ConnectCount1200 As Integer
        ConnectCount1200 = 0
        PSTNConnectCount1200.Text = ConnectCount1200

        Dim ConnectCount2400 As Integer
        ConnectCount2400 = 0
        PSTNConnectCount2400.Text = ConnectCount2400

        Dim BusyCount As Integer
        BusyCount = 0
        PSTNBusyCount.Text = BusyCount

        Dim NoCarrierCount As Integer
        NoCarrierCount = 0
        PSTNNoCarrierCount.Text = NoCarrierCount

        Dim deviceaddress As String
        Dim timeofrequest As String
        Dim deviceDialString As String = ""
        Dim CommPortDetails As String

        For Each filename As String In pstnoutlistbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine

                Dateinfile = DateTimePickerSplitterDL.Value.Date.ToString("ddd MMM dd")
                If InStr(1, strLine, Dateinfile) Then ' only do this for the messages with the correct date


                    Dim Searchfor As String = "Work Request"
                    If InStr(1, strLine, Searchfor) Then

                        deviceaddress = Microsoft.VisualBasic.Right(strLine, 9)
                        deviceaddress = Microsoft.VisualBasic.Left(deviceaddress, 6)
                        timeofrequest = Microsoft.VisualBasic.Mid(strLine, 18, 12)

                        For x = 1 To 40 ' loop round after the work request looking for a dial

                            '       MessageBox.Show("strLine = '" & strLine & " dial device ='" & deviceaddress & "time of request " & timeofrequest)

                            strLine = FileReader.ReadLine


                            If InStr(1, strLine, "ATV1&A2Q0M0DT9") Then    'string search for dial out at 300baud
                                deviceDialString = Microsoft.VisualBasic.Right(strLine, 11)
                                DialCount300 = DialCount300 + 1
                                PSTNDialCount300.Text = DialCount300

                                For nextrdimessage = 1 To 10
                                    strLine = FileReader.ReadLine
                                    If InStr(1, strLine, "CONNECT") Then
                                        ConnectCount300 = ConnectCount300 + 1
                                        PSTNConnectCount300.Text = ConnectCount300
                                        For nextrdimessageafterconnect = 1 To 15
                                            strLine = FileReader.ReadLine
                                            If InStr(1, strLine, "RIER") Then ' Or nextrdimessageafterconnect = 10 Then ' check here die to 300 Baud tries from Stepps fail carrier afer connect

                                                'If InStr(1, strLine, "W +++") Then
                                                WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                                WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount

                                                If CheckBoxLogTimeouts.Checked Then
                                                    Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                                    CommPortDetails = Mid(filename, TestPosition, 12)
                                                    populatexcelsheet(WorkPackageNoDialResultCount, 15, "ABORTED 300 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "")
                                                    GoTo foundresult
                                                End If
                                            End If
                                        Next nextrdimessageafterconnect

                                        GoTo foundresult
                                    End If

                                    If InStr(1, strLine, "BUSY") Then
                                        BusyCount = BusyCount + 1
                                        PSTNBusyCount.Text = BusyCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "NO CARRIER") Then
                                        NoCarrierCount = NoCarrierCount + 1
                                        PSTNNoCarrierCount.Text = NoCarrierCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "WorkPackageComplete") Or nextrdimessage = 10 Then
                                        strLine = FileReader.ReadLine
                                        'If InStr(1, strLine, "W +++") Then
                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount

                                        If CheckBoxLogTimeouts.Checked Then
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 15, "ABORTED 300 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "")
                                        End If

                                        GoTo foundresult
                                        'End If
                                    End If


                                    'If nextrdimessage = 10 Then
                                    '    MessageBox.Show("no connect busy or no carrier result in 10 messages -300--- " & filename)
                                    'End If





                                Next nextrdimessage

                            End If

                            If InStr(1, strLine, "ATV1Q0M0DT9") Then    'string search for dial out at 1200baud
                                deviceDialString = Microsoft.VisualBasic.Right(strLine, 11)
                                DialCount1200 = DialCount1200 + 1
                                PSTNDialCount1200.Text = DialCount1200


                                For nextrdimessage = 1 To 10

                                    ' need to trap a dial that goes nowhere 'WorkPackageComplete' then W +++
                                    'WorkPackageComplete
                                    ' W +++


                                    strLine = FileReader.ReadLine
                                    If InStr(1, strLine, "CONNECT 1200") Then    'string search and write back if match
                                        ConnectCount1200 = ConnectCount1200 + 1
                                        PSTNConnectCount1200.Text = ConnectCount1200
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "BUSY") Then
                                        BusyCount = BusyCount + 1
                                        PSTNBusyCount.Text = BusyCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "NO CARRIER") Then
                                        NoCarrierCount = NoCarrierCount + 1
                                        PSTNNoCarrierCount.Text = NoCarrierCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "WorkPackageComplete") Or nextrdimessage = 10 Then
                                        strLine = FileReader.ReadLine
                                        'If InStr(1, strLine, "W +++") Then
                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount


                                        If CheckBoxLogTimeouts.Checked Then
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 25, "ABORTED 1200 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "")
                                        End If

                                        GoTo foundresult
                                        'End If
                                    End If
                                    'If nextrdimessage = 10 Then
                                    '    MessageBox.Show("no connect busy or no carrier result in 10 messages -1200--- " & filename)
                                    'End If
                                Next nextrdimessage

                            End If

                            If InStr(1, strLine, "ATDT9") Then    'string search for dial out at 1200baud
                                deviceDialString = Microsoft.VisualBasic.Right(strLine, 11)
                                DialCount2400 = DialCount2400 + 1
                                PSTNDialCount2400.Text = DialCount2400


                                For nextrdimessage = 1 To 10
                                    strLine = FileReader.ReadLine
                                    If InStr(1, strLine, "CONNECT 2400") Then    'string search and write back if match
                                        ConnectCount2400 = ConnectCount2400 + 1
                                        PSTNConnectCount2400.Text = ConnectCount2400
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "BUSY") Then
                                        BusyCount = BusyCount + 1
                                        PSTNBusyCount.Text = BusyCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "NO CARRIER") Then
                                        NoCarrierCount = NoCarrierCount + 1
                                        PSTNNoCarrierCount.Text = NoCarrierCount
                                        GoTo foundresult
                                    End If
                                    If InStr(1, strLine, "WorkPackageComplete") Or nextrdimessage = 10 Then
                                        strLine = FileReader.ReadLine
                                        'If InStr(1, strLine, "W +++") Then
                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount



                                        If CheckBoxLogTimeouts.Checked Then

                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 35, "ABORTED 2400 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "")
                                        End If

                                        GoTo foundresult
                                        'End If
                                    End If
                                    'If nextrdimessage = 10 Then
                                    '    MessageBox.Show("no connect busy or no carrier result in 10 messages -2400--- " & filename)
                                    'End If
                                Next nextrdimessage

                            End If

                        Next x


                        'findresult:

                        'newdial:
foundresult:
                    End If

                End If


            Loop

            FileReader.Close()


        Next


        DialCount = DialCount300 + DialCount1200 + DialCount2400
        ConnectCountAll = ConnectCount300 + ConnectCount1200 + ConnectCount2400
        PSTNDialCountAll.Text = DialCount
        Failure1.Text = (DialCount - (ConnectCountAll + BusyCount)) / DialCount * 100
        Failure2.Text = (NoCarrierCount / DialCount) * 100
        '   MessageBox.Show("Dialcount = " & DialCount & " Connectcout all = " & ConnectCountAll & " busycount" & BusyCount)

    End Sub




    Private Sub ProcessPSTN_Click(sender As Object, e As EventArgs) Handles ProcessPSTN.Click

        Dlfilesofdatelistbox.Items.Clear()
        pstninlistbox.Items.Clear()
        pstnoutlistbox.Items.Clear()

        'com01
        If CheckBoxPSTNCom01IN.Checked Then
            GetDLLogFilesDLofDate("COM01", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom01OUT.Checked Then
            GetDLLogFilesDLofDate("COM01", "11", pstnoutlistbox)
        End If
        'gsm 1200
        If CheckBoxGSMCom011200.Checked Then
            GetDLLogFilesDLofDate("COM01", "36", pstninlistbox)
        End If

        'com02
        If CheckBoxPSTNCom02IN.Checked Then
            GetDLLogFilesDLofDate("COM02", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom02OUT.Checked Then
            GetDLLogFilesDLofDate("COM02", "11", pstnoutlistbox)
        End If

        If CheckBoxPSTNCom02IN2.Checked Then
            GetDLLogFilesDLofDate("COM02", "38", pstninlistbox)
        End If

        'gsm
        If CheckBoxGSMCom022400.Checked Then
            GetDLLogFilesDLofDate("COM02", "36", pstninlistbox)
        End If
        If CheckBoxGSMCom02_2_2400.Checked Then
            GetDLLogFilesDLofDate("COM02", "37", pstninlistbox)
        End If

        'com03
        If CheckBoxPSTNCom03IN.Checked Then
            GetDLLogFilesDLofDate("COM03", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom03OUT.Checked Then
            GetDLLogFilesDLofDate("COM03", "11", pstnoutlistbox)
        End If

        'gsm in 2400
        If CheckBoxGSMCom032400.Checked Then
            GetDLLogFilesDLofDate("COM03", "19", pstninlistbox)
        End If

        'com04
        If CheckBoxPSTNCom04IN.Checked Then
            GetDLLogFilesDLofDate("COM04", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom04OUT.Checked Then
            GetDLLogFilesDLofDate("COM04", "11", pstnoutlistbox)
            CountOutPSTN()
        End If
        'gsm in 2400
        If CheckBoxGSMCom042400.Checked Then
            GetDLLogFilesDLofDate("COM04", "19", pstninlistbox)
        End If



        'com05
        If CheckBoxPSTNCom05IN.Checked Then
            GetDLLogFilesDLofDate("COM05", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom05OUT.Checked Then
            GetDLLogFilesDLofDate("COM05", "11", pstnoutlistbox)
        End If
        'gsm in 2400
        If CheckBoxGSMCom052400.Checked Then
            GetDLLogFilesDLofDate("COM05", "39", pstninlistbox)
        End If


        'com06
        If CheckBoxPSTNCom06IN.Checked Then
            GetDLLogFilesDLofDate("COM06", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom06OUT.Checked Then
            GetDLLogFilesDLofDate("COM06", "11", pstnoutlistbox)
        End If

        If CheckBoxPSTNCom06IN2.Checked Then
            GetDLLogFilesDLofDate("COM06", "38", pstninlistbox)
        End If
        'gsm
        If CheckBoxGSMCom062400.Checked Then
            GetDLLogFilesDLofDate("COM06", "36", pstninlistbox)
        End If
        If CheckBoxGSMCom06_2_2400.Checked Then
            GetDLLogFilesDLofDate("COM06", "37", pstninlistbox)
        End If


        'com07
        If CheckBoxPSTNCom07IN.Checked Then
            GetDLLogFilesDLofDate("COM07", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom07OUT.Checked Then
            GetDLLogFilesDLofDate("COM07", "11", pstnoutlistbox)
        End If
        'gsm 1200
        If CheckBoxGSMCom071200.Checked Then
            GetDLLogFilesDLofDate("COM07", "19", pstninlistbox)
        End If

        'com08
        If CheckBoxPSTNCom08IN.Checked Then
            GetDLLogFilesDLofDate("COM08", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom08OUT.Checked Then
            GetDLLogFilesDLofDate("COM08", "11", pstnoutlistbox)
        End If
        'gsm 1200
        If CheckBoxGSMCom081200.Checked Then
            GetDLLogFilesDLofDate("COM08", "19", pstninlistbox)
        End If

        'com09
        If CheckBoxPSTNCom09IN.Checked Then
            GetDLLogFilesDLofDate("COM09", "29", pstninlistbox)
        End If
        If CheckBoxPSTNCom09IN2.Checked Then
            GetDLLogFilesDLofDate("COM09", "12", pstninlistbox)
        End If
        If CheckBoxPSTNCom09OUT.Checked Then
            GetDLLogFilesDLofDate("COM09", "11", pstnoutlistbox)
        End If
        If CheckBoxPSTNCom09OUT2.Checked Then
            GetDLLogFilesDLofDate("COM09", "28", pstnoutlistbox)
        End If
        'run the calulations ' if nothing checked then it zeros the totals
        CountINPSTN()
        CountOutPSTN()
    End Sub


    Private Sub Dlfilesofdatelistbox_DoubleClick(sender As Object, e As EventArgs) Handles Dlfilesofdatelistbox.DoubleClick
        System.Diagnostics.Process.Start(Dlfilesofdatelistbox.SelectedItem.ToString())
    End Sub

    Private Sub pstninlistbox_DoubleClick(sender As Object, e As EventArgs) Handles pstninlistbox.DoubleClick
        System.Diagnostics.Process.Start(pstninlistbox.SelectedItem.ToString())

    End Sub
    Private Sub pstnoutlistbox_DoubleClick(sender As Object, e As EventArgs) Handles pstnoutlistbox.DoubleClick
        System.Diagnostics.Process.Start(pstnoutlistbox.SelectedItem.ToString())

    End Sub
    'Private Sub gsminoutlistbox_DoubleClick(sender As Object, e As EventArgs)
    '    System.Diagnostics.Process.Start(gsminoutlistbox.SelectedItem.ToString())

    'End Sub


    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dlfilesofdatelistbox.Visible = True
        pstnoutlistbox.Visible = True
        pstninlistbox.Visible = True
        '     gsminoutlistbox.Visible = True
    End Sub

    Private Sub DisableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisableToolStripMenuItem.Click
        Dlfilesofdatelistbox.Visible = False
        pstnoutlistbox.Visible = False
        pstninlistbox.Visible = False
        '    gsminoutlistbox.Visible = False
    End Sub

    Private Sub clearforms()
        Dlfilesofdatelistbox.Items.Clear()
        pstninlistbox.Items.Clear()
        pstnoutlistbox.Items.Clear()
    End Sub




    Private Sub RunAnalysisforallsave()
        Dlfilesofdatelistbox.Items.Clear()
        pstninlistbox.Items.Clear()
        pstnoutlistbox.Items.Clear()
        On Error GoTo Errornofilefound
        PleaseWait.Show()
        'com01

        GetDLLogFilesDLofDate("COM01", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(4, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)
        clearforms()

        '    If CheckBoxPSTNCom01OUT.Checked Then
        GetDLLogFilesDLofDate("COM01", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(15, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()


        'gsm 1200
        '   If CheckBoxGSMCom011200.Checked Then
        '     GetDLLogFilesDLofDate("COM01", "36", pstninlistbox)
        '     CountINPSTN()
        '     populatexcelsheet(4, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, "", "", PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, "")
        '     clearforms()

        'com02

        GetDLLogFilesDLofDate("COM02", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(5, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        'If CheckBoxPSTNCom02OUT.Checked Then
        GetDLLogFilesDLofDate("COM02", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(16, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()

        GetDLLogFilesDLofDate("COM02", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(6, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)
        clearforms()
        'gsm
        ' If CheckBoxGSMCom022400.Checked Then
        GetDLLogFilesDLofDate("COM02", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(24, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        '   If CheckBoxGSMCom02_2_2400.Checked Then
        GetDLLogFilesDLofDate("COM02", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(25, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)
        clearforms()


        'com03

        GetDLLogFilesDLofDate("COM03", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(7, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()
        '   If CheckBoxPSTNCom03OUT.Checked Then
        GetDLLogFilesDLofDate("COM03", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(17, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()



        'gsm in 2400
        '  If CheckBoxGSMCom032400.Checked Then
        GetDLLogFilesDLofDate("COM03", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(26, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()


        'com04
        '      CheckBoxPSTNCom04IN.Checked Then
        GetDLLogFilesDLofDate("COM04", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(8, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()
        '      
        '   If CheckBoxPSTNCom04OUT.Checked Then
        GetDLLogFilesDLofDate("COM04", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(18, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()




        'gsm in 2400
        '  If CheckBoxGSMCom042400.Checked Then
        GetDLLogFilesDLofDate("COM04", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(27, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()




        'com05
        '   CheckBoxPSTNCom05IN.Checked Then
        GetDLLogFilesDLofDate("COM05", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(9, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        '   If CheckBoxPSTNCom05OUT.Checked Then
        GetDLLogFilesDLofDate("COM05", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(19, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()



        'gsm in 2400
        '    If CheckBoxGSMCom052400.Checked Then
        '      GetDLLogFilesDLofDate("COM05", "39", pstninlistbox)
        '      CountINPSTN()
        '      populatexcelsheet(26, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        '      clearforms()





        'com06
        ' CheckBoxPSTNCom06IN.Checked Then
        GetDLLogFilesDLofDate("COM06", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(10, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        '   If CheckBoxPSTNCom06OUT.Checked Then
        GetDLLogFilesDLofDate("COM06", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(20, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()


        '  CheckBoxPSTNCom06IN2.Checked Then
        GetDLLogFilesDLofDate("COM06", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(11, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        'gsm
        '    If CheckBoxGSMCom062400.Checked Then
        GetDLLogFilesDLofDate("COM06", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(28, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()



        'If CheckBoxGSMCom06_2_2400.Checked Then
        GetDLLogFilesDLofDate("COM06", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(29, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()





        'com07
        ' CheckBoxPSTNCom07IN.Checked Then
        GetDLLogFilesDLofDate("COM07", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(12, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        ' If CheckBoxPSTNCom07OUT.Checked Then
        GetDLLogFilesDLofDate("COM07", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(21, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()

        'gsm 1200
        '   If CheckBoxGSMCom071200.Checked Then
        GetDLLogFilesDLofDate("COM07", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(30, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()


        'com08
        ' CheckBoxPSTNCom08IN.Checked Then
        GetDLLogFilesDLofDate("COM08", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(13, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        '    If CheckBoxPSTNCom08OUT.Checked Then
        GetDLLogFilesDLofDate("COM08", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(22, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()

        'gsm 1200
        '  If CheckBoxGSMCom081200.Checked Then
        GetDLLogFilesDLofDate("COM08", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(31, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()


        'com09
        ' CheckBoxPSTNCom09IN.Checked Then
        GetDLLogFilesDLofDate("COM09", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(33, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)

        clearforms()

        ' CheckBoxPSTNCom09IN2.Checked Then
        GetDLLogFilesDLofDate("COM09", "29", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(34, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, Failure3PSTN.Text, Failure4PSTN.Text, Failure5PSTN.Text)
        clearforms()

        '    If CheckBoxPSTNCom09OUT.Checked Then
        GetDLLogFilesDLofDate("COM09", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(36, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()
        '    If CheckBoxPSTNCom09OUT2.Checked Then
        GetDLLogFilesDLofDate("COM09", "28", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(37, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, Failure1.Text, Failure2.Text)
        clearforms()

        PleaseWait.Close()

        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry ??? file open??")
        PleaseWait.Close()
    End Sub


    Private Sub populatexcelsheet(Row As Integer, Col As Integer, Data As String, Data2 As String, Data3 As String, Data4 As String, Data5 As String, Data6 As String, Data7 As String, Data8 As String)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")


        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")

        'add data
        oSheet = oBook.Worksheets("DL")
        oSheet.cells(Row, Col).Value = Data
        oSheet.cells(Row, Col + 1).Value = Data2
        oSheet.cells(Row, Col + 2).Value = Data3
        oSheet.cells(Row, Col + 3).Value = Data4
        oSheet.cells(Row, Col + 4).Value = Data5
        oSheet.cells(Row, Col + 5).Value = Data6
        oSheet.cells(Row, Col + 6).Value = Data7
        oSheet.cells(Row, Col + 7).Value = Data8

        ' save and exit
        oBook.Save()
        oBook.Application.DisplayAlerts = False

        Dim Dateinfile As String = DateTimePickerSplitterDL.Value.Date.ToString("dd-MMM-yyyy")
        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        oBook.Application.DisplayAlerts = True

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
    End Sub

    Private Sub clearlogdataexcelsheet()
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")


        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")

        'add data
        oSheet = oBook.Worksheets("DL")
        oSheet.Columns("O:AC").ClearContents()

        ' save and exit
        oBook.Save()

        'Dim Dateinfile As String = DateTimePickerSplitterDL.Value.Date.ToString("dd-MMM-yyyy")
        'oBook.Application.DisplayAlerts = False
        'oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI " & Dateinfile & ".xls")
        'oBook.Application.DisplayAlerts = True

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


        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI - Blank.xls")


        ' save and exit

        oBook.SaveAs(partialpath & "\OE Comms Channel Report_RDI.xls")
  

        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
    End Sub




    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        helpabout.Show()
    End Sub

    Private Sub EnableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnableToolStripMenuItem.Click

        Dlfilesofdatelistbox.Visible = True
        pstnoutlistbox.Visible = True
        pstninlistbox.Visible = True
    End Sub

    Private Sub DisableToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DisableToolStripMenuItem1.Click
        Dlfilesofdatelistbox.Visible = False
        pstnoutlistbox.Visible = False
        pstninlistbox.Visible = False
    End Sub

    Private Sub save_Click(sender As Object, e As EventArgs) Handles save.Click
        RunAnalysisforallsave()
    End Sub

    Private Sub OpenResultsFile_Click(sender As Object, e As EventArgs) Handles OpenResultsFile.Click
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
        System.Diagnostics.Process.Start(partialpath & "\OE Comms Channel Report_RDI.xls")
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles PSTNConnectCount1200.TextChanged

    End Sub


    Private Sub CheckBoxLogTimeouts_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLogTimeouts.CheckedChanged
        clearlogdataexcelsheet()

    End Sub




    Sub runsql()
        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As OdbcDataReader
        Dim connectionString As String
        Dim sql As String
        connectionString = "DSN=SWDL;SERVICE=DYNAMICLOGICA:RTRDB1,DYNAMICLOGICB:RTRDB1;UID=SYSTEM;PWD=SYSTTS09"
        sql = "SELECT * from dynamiclogicdevice where devicename like 'DL123%';"
        conn = New OdbcConnection(connectionString)
        conn.Open()
        comm = New OdbcCommand(sql, conn)
        dr = comm.ExecuteReader()

        While (dr.Read())

            '      DataGridViewDLDevice.DataSource = dr


            ListBoxsqlresults.Items.Add(dr.GetString(0))
            ListBoxsqlresults.Items.Add(dr.GetString(1))
            ListBoxsqlresults.Items.Add(dr.GetString(2))

            'Console.WriteLine(dr.GetValue(0).ToString())
            'Console.WriteLine(dr.GetValue(1).ToString())
            'Console.WriteLine(dr.GetValue(2).ToString())

        End While
        conn.Close()
        dr.Close()


        comm.Dispose()
        conn.Dispose()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        runsql()

    End Sub

    Private Sub OPTDebugToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OPTDebugToolStripMenuItem.Click
        DL_OPT.Show()
    End Sub




End Class