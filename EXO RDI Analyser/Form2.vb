Imports System.IO
Imports System.Data.Odbc

'dlpstnindone
Public Class DLForm

    Private Sub btnSelectFolderDL_Click(sender As Object, e As EventArgs) Handles btnSelectFolderDL.Click
        Dim Folder As String
        Folder = SelectFolderdl()
    End Sub


    Private Function SelectFolderdl() As String
        Dim DLdefaultfolder = driveletter + "RDI Analyser\DLRDI"
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
        SelectedFolderDL.Text = driveletter + "RDI Analyser\DLRDI"

    End Sub

    Private Sub stripLinefeedCRstuffOut(filepathandname As String, messagestring As String)
        'call for each message to strip off the CRCRLF crup so we can get the ring message on a timestamp so we know when it happens using read line by line objects
        Const ForReading = 1
        Const ForWriting = 2

        Dim objFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFile = objFSO.OpenTextFile(filepathandname, ForReading)
        Dim strLine
        Dim strNewContents = ""

        'open whole thing and strip out CR CR LF RING so the Ring is on the end of the timestamped message
        'call each time new message
        Do Until objFile.AtEndOfStream
            strLine = objFile.Readall
            strNewContents = Replace(strLine, Chr(013) & Chr(013) & Chr(010) & messagestring, messagestring) ' chr(010) = line feed chr(013) = carriage return 
            '  strNewContents = Replace(strLine, Chr(013) & Chr(013) & Chr(010) & "NO CARRIER", "mynewNO CARRIER") ' chr(010) = line feed chr(013) = carriage return 
        Loop
        'write it all back to the file with out the crud
        objFile.Close
        objFile = objFSO.OpenTextFile(filepathandname, ForWriting)
        objFile.Write(strNewContents)
        objFile.Close


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
        'Dim Dateinfileplus1day As String = DateAdd(DateInterval.Day, 1, DateTimePickerSplitterDL.Value.Date).ToString("ddd MMM dd")

        LogfilesFolder = SelectedFolderDL.Text & "\" & commscontroler
        Filenametag = "Port" & Port & "-" & "*.log"

        Dim fileNames = My.Computer.FileSystem.GetFiles(
          LogfilesFolder, FileIO.SearchOption.SearchTopLevelOnly, Filenametag)

        For Each fileName As String In fileNames
            'remote the crud from the files so we know when the interesting things happen with a time stamp


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
        '  Dim strLineAalarmString As String = ""

        Dim Dateinfile As String
        '  Dim ModemConfigString As String
        ' Dim nextrdimessage As Integer
        ' Dim Searchfor As String

        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'AT&FV1N0S0=2S7=60

        Dim ModemConfigString2400 As String = "AT&FV1N0S0=2S7=60"
        Dim ModemConfigString1200 As String = "ATV1E0&A0"
        Dim ModemConfigString300 As String = "AT&A2V1E0&W0"

        Dim ModemAutoAnswer2400 As String = "AT"
        Dim ModemAutoAnswer1200 As String = "ATV1S0=1Q0M0&A0"
        Dim ModemAutoAnswer300 As String = "ATV1&A2S0=1Q0M0"
        Dim addressposition As Integer
        Dim dldeviceaddress As String = "XXXX"

        Dim logedmessagealarms As String = ""

        Dim ModemIsautoanswer As Boolean = False
        Dim ModemIsConfiged As Boolean = False
        Dim ModemIsRinging As Boolean = False
        Dim ModemIsConnectedCorrectBaud As Boolean = False
        Dim ModemIsGoodalarms = False
        Dim ModemIsGoodAckedAlarms As Boolean = False
        Dim ParityError As Boolean = False
        Dim ModemIsSwamped As Boolean = False
        Dim OEIsWorkPackageComplete As Boolean = False

        ' Dim ModemIsConnectedWRONGBaud As Boolean = False
        ' Dim ModemIsNoCarrier As Boolean = False
        ' Dim ModemIsCallActive As Boolean = False
        ' Dim ModemIsGotAddress As Boolean = False

        Dim ModemIsCallTerminated As Boolean = False

        'Dim ModemIsBadLogs = False
        ' Dim ModemIsAlarmRecovery = False
        Dim ThisCallHasManyParityerrors As Boolean = False

        Dim ErrorInRDIcount As Integer
        ErrorInRDIcount = 0
        RDIErrorIncount.Text = ErrorInRDIcount

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

        Dim AlarmAckCount As Integer
        AlarmAckCount = 0
        PSTNINAlarmAckCount.Text = AlarmAckCount

        Dim NoSpeakorResetCount As Integer
        NoSpeakorResetCount = 0
        PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount

        Dim SwampedCount As Integer
        SwampedCount = 0
        PSTNINSwampedCount.Text = SwampedCount

        Dim FailuretospeaklogdetailsCount As Integer = 2

        'Dim Connectmessageloop As Integer = 15

        Dim timeofring As String = "00:00:00"
        Dim timeofconnect As String = "00:00:00"
        Dim timeoferror As String = "00:00:00"
        Dim timeofmodemconfig As String = "00:00:00"
        Dim timeofmodemautoanswer As String = "00:00:00"
        Dim readyIdelTimeString As String = "00:00:00"
        Dim realtimeofmodemautoanswer As DateTime
        Dim realtimeofmodemconfig As DateTime
        Dim realtimeofring As DateTime
        Dim realReadyIdelTime As TimeSpan



        '      Dim realtimeoftimeofring As Date

        Dim CommPortDetails As String
        Dim networkpathlength As String = SelectedFolderDL.Text.Length - 6 ' strip off \DLRDI
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, networkpathlength)
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.open(partialpath & "\OE Comms Channel Report_RDI.xls")
        ' '''' leave to workbook open to speed things up
        For Each filename As String In pstninlistbox.Items()
            stripLinefeedCRstuffOut(filename, "RING")
            stripLinefeedCRstuffOut(filename, "NO CARRIER")
            stripLinefeedCRstuffOut(filename, "CONNECT")
            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the next line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine

                Dateinfile = DateTimePickerSplitterDL.Value.Date.ToString("ddd MMM dd")
                If InStr(1, strLine, Dateinfile) Then ' only do this for the messages with the correct date


                    ' *-*-*-*-*-*-* As we now put the interesting things on timestamped strings in file is OK just to let this go
                    'ok in here we must have the correct date BUT not all messaged are time stamped so start tighter loop before we go and check again
                    '         For DateStampScopex = 1 To 20 ' loop round after the incomming ring looking for what happens next
                    '               strLine = FileReader.ReadLine

                    Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                    CommPortDetails = Mid(filename, TestPosition, 12)
                    oSheet = oBook.Worksheets("DL")
                    Dim col As Integer = 20





                    Select Case True

                        Case strLine.Contains("+++")
                            ModemIsCallTerminated = True
                            ModemIsConfiged = False
                            ModemIsRinging = False

                            ModemIsGoodalarms = False
                            ModemIsConnectedCorrectBaud = False

                            If OEIsWorkPackageComplete Then
                                If CheckBoxLogInCallDetails.Checked Then
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Call Terminated" & " from ring recieved at " & timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                If ParityError = True Then
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Parity errors "
                                Else
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Parity errors "
                                End If
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From address " & dldeviceaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Connect Received at " & timeofconnect
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "" 'logedmessagetimesync
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 6).Value = "" 'logedmessagelogs
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 7).Value = logedmessagealarms
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1


                            End If

                            End If


                        Case strLine.Contains(ModemConfigString300) Or strLine.Contains(ModemConfigString1200) Or strLine.Contains(ModemConfigString2400)
                            ModemIsautoanswer = False
                            ModemIsConfiged = True
                            ModemIsRinging = False
                            ModemIsConnectedCorrectBaud = False
                            ModemIsSwamped = False
                            OEIsWorkPackageComplete = False
                            timeofmodemconfig = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                            realtimeofmodemconfig = Convert.ToDateTime(timeofmodemconfig)
                        Case strLine.Contains(ModemAutoAnswer300) Or strLine.Contains(ModemAutoAnswer1200) Or strLine.Contains(ModemAutoAnswer2400)
                            ModemIsautoanswer = True
                            ModemIsConfiged = True
                            ModemIsRinging = False
                            ModemIsConnectedCorrectBaud = False
                            ModemIsSwamped = False
                            logedmessagealarms = ""
                            OEIsWorkPackageComplete = False
                            timeofmodemautoanswer = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                            realtimeofmodemautoanswer = Convert.ToDateTime(timeofmodemautoanswer)
                        Case strLine.Contains("RING")
                            dldeviceaddress = "XXXX"
                            If ModemIsRinging = True Then
                                ' alredy ringing
                                'may ring >1 now so ok


                            Else
                                'new ring
                                ModemIsRinging = True
                                If strLine.Length > 14 Then
                                    timeofring = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                                    realtimeofring = Convert.ToDateTime(timeofring)
                                End If
                                RingCount = RingCount + 1
                                PSTNINRingCount.Text = RingCount
                            End If

                            If CheckBoxLogPerformaceInDetails.Checked Then
                                realReadyIdelTime = realtimeofring - realtimeofmodemconfig
                                readyIdelTimeString = realReadyIdelTime.ToString()

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ready Status Modem is Autoanswer "
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Modem Ring at " & timeofring
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "Last Modem Config at " & timeofmodemconfig
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Last Modem AutoAnswer at " & timeofmodemautoanswer
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 5).Value = "Ready Idel time " & readyIdelTimeString
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If

                            If ModemIsConnectedCorrectBaud = True Then
                                ModemIsSwamped = True
                                ' NoSpeakorResetCount = NoSpeakorResetCount + 1
                                SwampedCount = SwampedCount + 1
                                PSTNINSwampedCount.Text = SwampedCount

                                If CheckBoxLogPerformaceInDetails.Checked Then

                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "***Swamped*** Ring Then Connected - Ring again"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Swamped"
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                                End If
                            End If




                        Case strLine.Contains("CONNECT")
                            ModemIsautoanswer = True
                            ModemIsConfiged = True
                            ModemIsRinging = False
                            ModemIsConnectedCorrectBaud = True
                            ConnectBAUDCount = ConnectBAUDCount + 1
                            PSTNINConnectBAUDCount.Text = ConnectBAUDCount
                            If strLine.Length > 14 Then
                                timeofconnect = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                            End If



                        Case strLine.Contains("Error")
                            ErrorInRDIcount = ErrorInRDIcount + 1
                            RDIErrorIncount.Text = ErrorInRDIcount
                            If strLine.Length > 14 Then
                                timeoferror = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                            End If
                            If CheckBoxLogPerformaceInDetails.Checked Then
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Error in RDI"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = strLine & "at " & timeoferror
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at " & timeofring & " From " & dldeviceaddress

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Duration " & timeofring & " to " & timeoferror
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If



                        Case strLine.Contains("NO CARRIER")
                            NoCarrierCount = NoCarrierCount + 1
                            PSTNINNoCarrierCount.Text = NoCarrierCount
                            'gracefull exit nothing to see here
                            If CheckBoxLogPerformaceInDetails.Checked Then
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "No Carrier"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Carrier"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If






                        Case strLine.Contains(" * " & ChrW(17) & " * " & ChrW(17) & " * ")


                            ModemIsSwamped = True
                            ' NoSpeakorResetCount = NoSpeakorResetCount + 1
                            SwampedCount = SwampedCount + 1
                            PSTNINSwampedCount.Text = SwampedCount

                            If CheckBoxLogPerformaceInDetails.Checked Then

                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "***Swamped*** - Ring Then Corrupt Device Control 1"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "CORRUPT"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If




                        Case strLine.Contains("  R A<")
                            addressposition = InStr(1, strLine, "  R A<")
                            ModemIsGoodalarms = True
                            ' strLineAalarmString = strLine
                            ' get rtu address
                            dldeviceaddress = Microsoft.VisualBasic.Mid(strLine, addressposition + 6, 4)
                            logedmessagealarms = logedmessagealarms & "Alarms Recovered"
                            AlarmRequestCount = AlarmRequestCount + 1
                            PSTNINAlarmRequestCount.Text = AlarmRequestCount

                            If CheckBoxLogPerformaceInDetails.Checked Then
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring Then RTU responce"
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "R A responce" & " From " & dldeviceaddress
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                                oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                                populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                            End If

                        Case strLine.Contains("  R T<E>")
                            ModemIsGoodAckedAlarms = True

                            AlarmAckCount = AlarmAckCount + 1
                            PSTNINAlarmAckCount.Text = AlarmAckCount
                            logedmessagealarms = logedmessagealarms & " Alarms Acknowleged"

                        Case strLine.Contains("WorkPackageComplete")
                            OEIsWorkPackageComplete = True

                            If ModemIsGoodalarms = False Then
                                logedmessagealarms = logedmessagealarms & "NO Alarms Recovered"
                                NoSpeakorResetCount = NoSpeakorResetCount + 1
                                PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount
                            End If

                            If ModemIsGoodAckedAlarms = False Then
                                logedmessagealarms = logedmessagealarms & "NO Alarms Acknowleged"
                                NoSpeakorResetCount = NoSpeakorResetCount + 1
                                PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount
                            End If

                        Case strLine.Contains("Parity")
                            ParityError = True
                            'If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount < 3 And ThisCallHasManyParityerrors = False Then
                            '    ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            'End If

                            'If ModemIsConnectedCorrectBaud = True And ParityErrorthiscallCount >= 3 Then
                            '    ParityErrorthiscallCount = ParityErrorthiscallCount + 1
                            '    ThisCallHasManyParityerrors = True

                            'End If


                    End Select




                    'Searchfor = "Error"
                    'If InStr(1, strLine, Searchfor) Then
                    '    ErrorInRDIcount = ErrorInRDIcount + 1
                    '    RDIErrorIncount.Text = ErrorInRDIcount
                    '    If strLine.Length > 14 Then
                    '        timeoferror = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                    '    End If
                    '    If CheckBoxLogFailInDetails.Checked Then
                    '        oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Error in RDI"
                    '        oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '        oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = strLine & "at " & timeoferror
                    '        oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at " & timeofring & " From " & strLineAalarmString

                    '        oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = "Duration " & timeofring & " to " & timeoferror
                    '        populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '    End If


                    'End If

                    '                    Searchfor = "RING"
                    '                    If InStr(1, strLine, Searchfor) Then
                    '                        If strLine.Length > 14 Then
                    '                            timeofring = Microsoft.VisualBasic.Mid(strLine, 17, 9)
                    '                        End If
                    '                        RingCount = RingCount + 1
                    '                        PSTNINRingCount.Text = RingCount

                    '                        For x = 1 To 20 ' loop round after the incomming ring looking for what happens next
                    '                            'should be ring
                    '                            'connect
                    '                            'T<A>
                    '                            'reply from rtu with string
                    '                            '*workPackageComplete
                    '                            '* Work Request DLxxxx:DL
                    'ProcessIncomming:
                    '                            strLine = FileReader.ReadLine


                    '                            '************** Start of error checkeing for call Afer RING ****************************

                    '                            If InStr(1, strLine, "NO CARRIER") Then    'string search 
                    '                                NoCarrierCount = NoCarrierCount + 1
                    '                                PSTNINNoCarrierCount.Text = NoCarrierCount
                    '                                'gracefull exit nothing to see here
                    '                                If CheckBoxLogFailInDetails.Checked Then
                    '                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring Then No Carrier"
                    '                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "No Carrier"
                    '                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                    oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                    populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                End If


                    '                                GoTo foundresult
                    '                            End If

                    '                            If InStr(1, strLine, "CONNECT") Then    'string search 
                    '                                ConnectBAUDCount = ConnectBAUDCount + 1
                    '                                PSTNINConnectBAUDCount.Text = ConnectBAUDCount

                    '                                For nextrdimessage = 1 To Connectmessageloop 'loop round after the connect to chech for what happens

                    '                                    ''just a check if we loop too long looking for someting that never happens
                    '                                    strLine = FileReader.ReadLine
                    '                                    If InStr(1, strLine, "RING") Then 'holy smoke batman getting a new call get in
                    '                                        '           MessageBox.Show(" Crashed into next ring  previous call nothing after connect in 10 message loops")

                    '                                        NoSpeakorResetCount = NoSpeakorResetCount + 1
                    '                                        SwampedCount = SwampedCount + 1
                    '                                        PSTNINSwampedCount.Text = SwampedCount

                    '                                        If CheckBoxLogFailInDetails.Checked Then

                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "***Swamped*** Ring Then Connected - Ring again"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Swamped"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                        End If


                    '                                        GoTo ProcessIncomming
                    '                                    End If


                    '                                    'decice control 1 char from modem
                    '                                    '   If InStr(1, strLine, ChrW(17)) Then 'holy smoke batman getting sh1t in
                    '                                    'sometime we get a single char from the modem so check for 2 
                    '                                    If strLine Like "*" & ChrW(17) & "*" & ChrW(17) & "*" Then 'holy smoke batman getting sh1t in

                    '                                        NoSpeakorResetCount = NoSpeakorResetCount + 1
                    '                                        SwampedCount = SwampedCount + 1
                    '                                        PSTNINSwampedCount.Text = SwampedCount

                    '                                        If CheckBoxLogFailInDetails.Checked Then

                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "***Swamped*** - Ring Then Corrupt Device Control 1"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Swamped"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                        End If


                    '                                        GoTo ProcessIncomming
                    '                                    End If




                    '                                    ' if call rings then connects but no traffic it a no speak, could be incorrect device type??? ******************************************

                    '                                    If InStr(1, strLine, "WorkPackageComplete") Then 'getting RDI not happy before we have valid responce
                    '                                        NoSpeakorResetCount = NoSpeakorResetCount + 1
                    '                                        PSTNINNOSpeakorResetCount.Text = NoSpeakorResetCount
                    '                                        If CheckBoxLogFailInDetails.Checked Then

                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring And connected but failed RDI said WorkPackageComplete"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Connected but did Not speak In " & nextrdimessage & " messages after connect"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1
                    '                                        End If

                    '                                        GoTo foundresult
                    '                                    End If

                    '                                    ' if call rings then connects but no traffic it a no speak, could be incorrect device type??? ******************************************

                    '                                    If InStr(1, strLine, ModemConfigString300) Or InStr(1, strLine, ModemConfigString1200) Or InStr(1, strLine, ModemConfigString2400) Then '

                    '                                        NoSpeakorResetCount = NoSpeakorResetCount + 1
                    '                                        SwampedCount = SwampedCount + 1
                    '                                        PSTNINSwampedCount.Text = SwampedCount

                    '                                        If CheckBoxLogFailInDetails.Checked Then

                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "***Swamped*** - Ring Then Reset"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Swamped"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                        End If
                    '                                        GoTo foundresult
                    '                                    End If




                    '                                    If InStr(1, strLine, "  R A<") Then 'responce from RTU 
                    '                                        strLineAalarmString = strLine
                    '                                        ' get rtu address

                    '                                        AlarmRequestCount = AlarmRequestCount + 1
                    '                                        PSTNINAlarmRequestCount.Text = AlarmRequestCount

                    '                                        If CheckBoxLogGoodInDetails.Checked Then
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring Then RTU responce"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "R A responce" & " From " & strLineAalarmString
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                        End If



                    '                                        GoTo foundresult
                    '                                    End If


                    '                                    If nextrdimessage > Connectmessageloop Then


                    '                                        NoSpeakorResetCount = NoSpeakorResetCount + 1
                    '                                        SwampedCount = SwampedCount + 1
                    '                                        PSTNINSwampedCount.Text = SwampedCount

                    '                                        If CheckBoxLogFailInDetails.Checked Then

                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col).Value = "Ring but Swamped - no valid reponce Or reset Or RDI Job done message In " & nextrdimessage & "messages"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 1).Value = CommPortDetails
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 2).Value = "Swamped"
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 3).Value = "From Ring Received at "
                    '                                            oSheet.cells(populatexcelsheetdetailsRowCounter, col + 4).Value = timeofring
                    '                                            populatexcelsheetdetailsRowCounter = populatexcelsheetdetailsRowCounter + 1

                    '                                        End If
                    '                                        'MessageBox.Show(" no alarm reply Or reset found")
                    '                                    End If
                    '                                Next nextrdimessage
                    '                            End If
                    '                        Next x
                    '                    End If





                End If


foundresult:
            Loop

            FileReader.Close()

        Next
        oBook.Save()
        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing

        RingthenConnectpct.Text = String.Format("{0:n2}", (ConnectBAUDCount / RingCount) * 100)
        RingthenNoCarrierpct.Text = String.Format("{0:n2}", (NoCarrierCount / RingCount) * 100)
        RingthenSpokepct.Text = String.Format("{0:n2}", (AlarmRequestCount / RingCount) * 100)
        FailuretospeakPSTN.Text = String.Format("{0:n2}", ((RingCount - AlarmRequestCount) / RingCount) * 100)

        'identified failure totals calcilation
        'FailuretospeakPSTN.Text = String.Format("{0:n2}", ((NoCarrierCount + NoSpeakorResetCount + ResetAfterConnectCount + ErrorInRDIcount) / RingCount) * 100)
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

        Dim IdleTimeLogCount As Integer
        IdleTimeLogCount = 2

        Dim ErrorOutRDIcount As Integer
        ErrorOutRDIcount = 0
        RDIErrorOutcount.Text = ErrorOutRDIcount

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

        Dim deviceaddress As String = ""
        Dim timeofrequest As String = "00:00:00"
        Dim timeofcomplete As String = "00:00:00"
        Dim deviceDialString As String = ""
        Dim CommPortDetails As String




        For Each filename As String In pstnoutlistbox.Items()

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine

                Dateinfile = DateTimePickerSplitterDL.Value.Date.ToString("ddd MMM dd")
                If InStr(1, strLine, Dateinfile) Then ' only do this for the messages with the correct date

                    Dim Searchfor As String = "Error"
                    If InStr(1, strLine, Searchfor) Then
                        ErrorOutRDIcount = ErrorOutRDIcount + 1
                        RDIErrorOutcount.Text = ErrorOutRDIcount
                    End If


                    Searchfor = "Work Request"
                    If InStr(1, strLine, Searchfor) Then

                        deviceaddress = Microsoft.VisualBasic.Right(strLine, 9)
                        deviceaddress = Microsoft.VisualBasic.Left(deviceaddress, 6)
                        timeofrequest = Microsoft.VisualBasic.Mid(strLine, 18, 12)

                        For x = 1 To 40 ' 40 ' loop round after the work request looking for a dial

                            '       MessageBox.Show("strLine = '" & strLine & " dial device ='" & deviceaddress & "time of request " & timeofrequest)

                            strLine = FileReader.ReadLine





                            ' check to see if we have crashed into new call
                            Searchfor = "Work Request"
                            If InStr(1, strLine, Searchfor) Then
                                WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                CommPortDetails = Mid(filename, TestPosition, 12)
                                populatexcelsheet(WorkPackageNoDialResultCount, 25, "Call unresolved work request and no outcome", deviceaddress, timeofrequest, CommPortDetails, "", "", "", "", "", "", "")
                            End If



                            If InStr(1, strLine, "ATV1&A2Q0M0DT9") Then    'string search for dial out at 300baud
                                deviceDialString = Microsoft.VisualBasic.Right(strLine, 11)
                                DialCount300 = DialCount300 + 1
                                PSTNDialCount300.Text = DialCount300

                                For nextrdimessage = 1 To 20
                                    strLine = FileReader.ReadLine
                                    If InStr(1, strLine, "CONNECT") Then
                                        ConnectCount300 = ConnectCount300 + 1
                                        PSTNConnectCount300.Text = ConnectCount300
                                        For nextrdimessageafterconnect = 1 To 20
                                            strLine = FileReader.ReadLine
                                            If InStr(1, strLine, "RIER") Then ' Or nextrdimessageafterconnect = 10 Then ' check here die to 300 Baud tries from Stepps fail carrier afer connect

                                                'If InStr(1, strLine, "W +++") Then
                                                WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                                WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount

                                                If CheckBoxLogTimeouts.Checked Then
                                                    Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                                    CommPortDetails = Mid(filename, TestPosition, 12)
                                                    populatexcelsheet(WorkPackageNoDialResultCount, 15, "ABORTED 300 Dial CORRUPTED NO CARRIER", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "", "", "", "")
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
                                    If InStr(1, strLine, "WorkPackageComplete") Or nextrdimessage = 20 Then

                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount
                                        timeofcomplete = Microsoft.VisualBasic.Mid(strLine, 18, 12)

                                        If CheckBoxLogTimeouts.Checked Then
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 25, "ABORTED 300 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "", "", "", "")
                                        End If

                                        If CheckBoxLogDurations.Checked Then
                                            IdleTimeLogCount = IdleTimeLogCount + 1
                                            '          WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(IdleTimeLogCount, 35, "Call Complete -ABORTED 300 Dial", deviceaddress, timeofrequest, timeofcomplete, CommPortDetails, "", "", "", "", "", "")
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

                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount
                                        timeofcomplete = Microsoft.VisualBasic.Mid(strLine, 18, 12)

                                        If CheckBoxLogTimeouts.Checked Then
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 25, "ABORTED 1200 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "", "", "", "")
                                        End If

                                        If CheckBoxLogDurations.Checked Then
                                            IdleTimeLogCount = IdleTimeLogCount + 1
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(IdleTimeLogCount, 35, "Call Complete ABORTED 1200 Dial", deviceaddress, timeofrequest, timeofcomplete, CommPortDetails, "", "", "", "", "", "")
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
                                        WorkPackageNoDialResultCount = WorkPackageNoDialResultCount + 1
                                        WorkPackageNoDialResult.Text = WorkPackageNoDialResultCount
                                        timeofcomplete = Microsoft.VisualBasic.Mid(strLine, 18, 12)


                                        If CheckBoxLogTimeouts.Checked Then

                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(WorkPackageNoDialResultCount, 25, "ABORTED 2400 Dial", deviceaddress, timeofrequest, deviceDialString, CommPortDetails, "", "", "", "", "", "")
                                        End If



                                        If CheckBoxLogDurations.Checked Then
                                            IdleTimeLogCount = IdleTimeLogCount + 1
                                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                                            CommPortDetails = Mid(filename, TestPosition, 12)
                                            populatexcelsheet(IdleTimeLogCount, 35, "Call Complete ABORTED 2400 Dial", deviceaddress, timeofrequest, timeofcomplete, CommPortDetails, "", "", "", "", "", "")
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
                        'after checking for what happend after dial for if good Or bad check for the end of the call for time tracking

                        'Searchfor = "WorkPackageComplete"
                        'If InStr(1, strLine, Searchfor) Then

                        '    timeofcomplete = Microsoft.VisualBasic.Mid(strLine, 18, 12)
                        '    If CheckBoxLogDurations.Checked Then
                        '        IdleTimeLogCount = IdleTimeLogCount + 1
                        '        Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                        '        CommPortDetails = Mid(filename, TestPosition, 12)
                        '        populatexcelsheet(IdleTimeLogCount, 35, "Call Complete after Dial checks- Good", deviceaddress, timeofrequest, timeofcomplete, CommPortDetails, strLine, "", "", "", "", "")
                        '    End If

                        'End If


foundresult:
                    End If 'end of 'work request Check'

                    ' check here for end of call to log 
                    Searchfor = "WorkPackageComplete"
                    If InStr(1, strLine, Searchfor) Then

                        timeofcomplete = Microsoft.VisualBasic.Mid(strLine, 18, 12)
                        If CheckBoxLogDurations.Checked Then
                            IdleTimeLogCount = IdleTimeLogCount + 1
                            Dim TestPosition As Integer = InStr(1, filename, "\DLRDI\", CompareMethod.Text) + 7 'add on 7 for the test string length too
                            CommPortDetails = Mid(filename, TestPosition, 12)
                            Dim idletime As TimeSpan = DateTime.Parse(timeofcomplete).Subtract(DateTime.Parse(timeofrequest))
                            Dim idletimetext As String = idletime.ToString
                            populatexcelsheet(IdleTimeLogCount, 35, "Call Complete after Dial checks- Good", deviceaddress, timeofrequest, timeofcomplete, idletimetext, CommPortDetails, strLine, "", "", "", "")
                        End If

                    End If
                End If 'end of date check 


            Loop

            FileReader.Close()


        Next


        DialCount = DialCount300 + DialCount1200 + DialCount2400
        ConnectCountAll = ConnectCount300 + ConnectCount1200 + ConnectCount2400
        PSTNDialCountAll.Text = DialCount
        'Failure1.Text = (DialCount - (ConnectCountAll + BusyCount)) / DialCount * 100
        PSTNGood.Text = String.Format("{0:n2}", (ConnectCount300 + ConnectCount1200 + ConnectCount2400 + BusyCount) / DialCount * 100)

        PSTNFails.Text = String.Format("{0:n2}", (NoCarrierCount + WorkPackageNoDialResultCount + ErrorOutRDIcount) / DialCount * 100)
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
        If CheckBoxPSTNCom01IN2.Checked Then
            GetDLLogFilesDLofDate("COM01", "38", pstninlistbox)
        End If
        If CheckBoxPSTNCom01OUT.Checked Then
            GetDLLogFilesDLofDate("COM01", "11", pstnoutlistbox)
        End If
        If CheckBoxPSTNCom01OUT18.Checked Then
            GetDLLogFilesDLofDate("COM01", "18", pstnoutlistbox)
        End If
        'gsm 2400
        If CheckBoxGSMCom012400.Checked Then
            GetDLLogFilesDLofDate("COM01", "36", pstninlistbox)
        End If
        'gsm 2400
        If CheckBoxGSMCom01_2_2400.Checked Then
            GetDLLogFilesDLofDate("COM01", "37", pstninlistbox)
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
        If CheckBoxPSTNCom05OUT17.Checked Then
            GetDLLogFilesDLofDate("COM05", "17", pstnoutlistbox)
        End If
        If CheckBoxPSTNCom05OUT35.Checked Then
            GetDLLogFilesDLofDate("COM05", "35", pstnoutlistbox)
        End If
        If CheckBoxPSTNCom05OUT18.Checked Then
            GetDLLogFilesDLofDate("COM05", "18", pstnoutlistbox)
        End If
        If CheckBoxPSTNCom05IN2.Checked Then
            GetDLLogFilesDLofDate("COM05", "38", pstninlistbox)
        End If
        'gsm 2400
        If CheckBoxGSMCom052400.Checked Then
            GetDLLogFilesDLofDate("COM05", "36", pstninlistbox)
        End If

        If CheckBoxGSMCom05_2_2400.Checked Then
            GetDLLogFilesDLofDate("COM05", "37", pstninlistbox)
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
        populatexcelsheet(4, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()

        GetDLLogFilesDLofDate("COM01", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(5, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()

        GetDLLogFilesDLofDate("COM01", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(17, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()
        GetDLLogFilesDLofDate("COM01", "18", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(48, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        'gsm 2400
        GetDLLogFilesDLofDate("COM01", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(26, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        GetDLLogFilesDLofDate("COM01", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(27, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()

        'com02

        GetDLLogFilesDLofDate("COM02", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(6, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        'If CheckBoxPSTNCom02OUT.Checked Then
        GetDLLogFilesDLofDate("COM02", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(18, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        GetDLLogFilesDLofDate("COM02", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(7, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()
        'gsm

        GetDLLogFilesDLofDate("COM02", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(28, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        GetDLLogFilesDLofDate("COM02", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(29, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()


        'com03

        GetDLLogFilesDLofDate("COM03", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(8, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()
        '   If CheckBoxPSTNCom03OUT.Checked Then
        GetDLLogFilesDLofDate("COM03", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(19, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()



        'gsm in 2400
        '  If CheckBoxGSMCom032400.Checked Then
        GetDLLogFilesDLofDate("COM03", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(30, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()


        'com04
        '      CheckBoxPSTNCom04IN.Checked Then
        GetDLLogFilesDLofDate("COM04", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(9, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()
        '      
        '   If CheckBoxPSTNCom04OUT.Checked Then
        GetDLLogFilesDLofDate("COM04", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(20, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()




        'gsm in 2400
        '  If CheckBoxGSMCom042400.Checked Then
        GetDLLogFilesDLofDate("COM04", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(31, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()




        'com05
        '   CheckBoxPSTNCom05IN.Checked Then
        GetDLLogFilesDLofDate("COM05", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(10, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        '   If CheckBoxPSTNCom05OUT.Checked Then
        GetDLLogFilesDLofDate("COM05", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(21, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        GetDLLogFilesDLofDate("COM05", "17", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(46, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        GetDLLogFilesDLofDate("COM05", "35", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(47, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()
        GetDLLogFilesDLofDate("COM05", "18", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(49, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()
        GetDLLogFilesDLofDate("COM05", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(32, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        GetDLLogFilesDLofDate("COM05", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(33, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()

        GetDLLogFilesDLofDate("COM05", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(11, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()


        'com06
        ' CheckBoxPSTNCom06IN.Checked Then
        GetDLLogFilesDLofDate("COM06", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(12, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        '   If CheckBoxPSTNCom06OUT.Checked Then
        GetDLLogFilesDLofDate("COM06", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(22, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()


        '  CheckBoxPSTNCom06IN2.Checked Then
        GetDLLogFilesDLofDate("COM06", "38", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(13, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        'gsm
        '    If CheckBoxGSMCom062400.Checked Then
        GetDLLogFilesDLofDate("COM06", "36", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(34, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()



        'If CheckBoxGSMCom06_2_2400.Checked Then
        GetDLLogFilesDLofDate("COM06", "37", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(35, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()


        'com07
        ' CheckBoxPSTNCom07IN.Checked Then
        GetDLLogFilesDLofDate("COM07", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(14, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        ' If CheckBoxPSTNCom07OUT.Checked Then
        GetDLLogFilesDLofDate("COM07", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(23, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        'gsm 1200
        '   If CheckBoxGSMCom071200.Checked Then
        GetDLLogFilesDLofDate("COM07", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(36, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()


        'com08
        ' CheckBoxPSTNCom08IN.Checked Then
        GetDLLogFilesDLofDate("COM08", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(15, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        '    If CheckBoxPSTNCom08OUT.Checked Then
        GetDLLogFilesDLofDate("COM08", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(24, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        'gsm 1200
        '  If CheckBoxGSMCom081200.Checked Then
        GetDLLogFilesDLofDate("COM08", "19", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(37, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()


        'com09
        ' CheckBoxPSTNCom09IN.Checked Then
        GetDLLogFilesDLofDate("COM09", "12", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(39, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)

        clearforms()

        ' CheckBoxPSTNCom09IN2.Checked Then
        GetDLLogFilesDLofDate("COM09", "29", pstninlistbox)
        CountINPSTN()
        populatexcelsheet(40, 5, PSTNINRingCount.Text, PSTNINConnectBAUDCount.Text, PSTNINAlarmRequestCount.Text, PSTNINNOSpeakorResetCount.Text, PSTNINNoCarrierCount.Text, RingthenConnectpct.Text, RingthenNoCarrierpct.Text, RingthenSpokepct.Text, FailuretospeakPSTN.Text, RDIErrorIncount.Text, PSTNINSwampedCount.Text)
        clearforms()

        '    If CheckBoxPSTNCom09OUT.Checked Then
        GetDLLogFilesDLofDate("COM09", "11", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(42, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()
        '    If CheckBoxPSTNCom09OUT2.Checked Then
        GetDLLogFilesDLofDate("COM09", "28", pstnoutlistbox)
        CountOutPSTN()
        populatexcelsheet(43, 5, PSTNDialCountAll.Text, PSTNConnectCount300.Text, PSTNConnectCount1200.Text, PSTNConnectCount2400.Text, PSTNBusyCount.Text, PSTNNoCarrierCount.Text, WorkPackageNoDialResult.Text, PSTNGood.Text, PSTNFails.Text, RDIErrorOutcount.Text, "")
        clearforms()

        PleaseWait.Close()

        Exit Sub

Errornofilefound:
        MessageBox.Show("Sorry ??? file open??")
        PleaseWait.Close()
    End Sub


    Private Sub populatexcelsheet(Row As Integer, Col As Integer, Data As String, Data2 As String, Data3 As String, Data4 As String, Data5 As String, Data6 As String, Data7 As String, Data8 As String, Data9 As String, Data10 As String, Data11 As String)
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'start new workbook
        oExcel = CreateObject("Excel.Application")

        Dim networkpathlength As String = SelectedFolderDL.Text.Length - 6 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, networkpathlength)
        '  Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
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
        oSheet.cells(Row, Col + 8).Value = Data9
        oSheet.cells(Row, Col + 9).Value = Data10
        oSheet.cells(Row, Col + 10).Value = Data11
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


        '  Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
        Dim networkpathlength As String = SelectedFolderDL.Text.Length - 6 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, networkpathlength)


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

        Dim networkpathlength As String = SelectedFolderDL.Text.Length - 6 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, networkpathlength)

        '  Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
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
        Dim networkpathlength As String = SelectedFolderDL.Text.Length - 6 ' strip off \EXORDI
        Dim partialpath As String = SelectedFolderDL.Text.Substring(0, networkpathlength)

        'Dim partialpath As String = SelectedFolderDL.Text.Substring(0, 15)
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




    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLogDurations.CheckedChanged

    End Sub

    Private Sub CheckBoxPSTNCom05OUT17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxPSTNCom05OUT17.CheckedChanged

    End Sub

    Private Sub PSTNINNOSpeakorResetCount_TextChanged(sender As Object, e As EventArgs) Handles PSTNINNOSpeakorResetCount.TextChanged

    End Sub

    Private Sub swamped_TextChanged(sender As Object, e As EventArgs) Handles PSTNINSwampedCount.TextChanged

    End Sub

    Private Sub PSTNINResetAfterConnectCount_TextChanged(sender As Object, e As EventArgs)

    End Sub
End Class