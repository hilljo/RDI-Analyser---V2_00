Imports System.IO
Public Class DL_OPT

    Private Sub OPTOpenFileDialog_Click()
        Dim FileReader As StreamReader
        ' Create an instance of the open file dialog box.
        Dim OpenFileDialogOPT As OpenFileDialog = New OpenFileDialog
        Dim strLine As String
        Dim Searchfor As String
        Dim GlasgowDestination As String = "90141"
        Dim DundeeDestination As String = "901382"
        Dim DialGlasgow As Integer = 0
        Dim OPTConnectGlasgow As Integer = 0
        Dim OPTGotResetGlasgow As Integer = 0
        Dim AlarmRequestGlasgow As Integer = 0
        Dim AlarmAckGlasgow As Integer = 0
        Dim NoCarrierGlasgow As Integer = 0
        Dim DialDundee As Integer = 0
        Dim OPTConnectDundee As Integer = 0
        Dim OPTGotResetDundee As Integer = 0
        Dim AlarmRequestDundee As Integer = 0
        Dim AlarmAckDundee As Integer = 0
        Dim NoCarrierDundee As Integer = 0

        ' Set filter options and filter index.
        OpenFileDialogOPT.InitialDirectory = "h:\RDI Analyser\OPTDebug"
        OpenFileDialogOPT.Filter = "Text Files (*.dbg)|*.dbg|All Files (*.*)|*.*"
        OpenFileDialogOPT.FilterIndex = 1

        OpenFileDialogOPT.Multiselect = False

        ' Call the ShowDialog method to show the dialogbox.
        Dim UserClickedOK As Boolean = OpenFileDialogOPT.ShowDialog

        ' Process input if the user clicked OK.
        If (UserClickedOK = True) Then
            'Open the selected file to read.
            Dim filename As String = OpenFileDialogOPT.FileName

            'System.Diagnostics.Process.Start("wordpad.exe", OpenFileDialogOPT.ToString())

            FileReader = New StreamReader(filename) ' set up read

            strLine = FileReader.ReadLine               ' read the first line until no more data
            Do While Not strLine Is Nothing
                strLine = FileReader.ReadLine

                Searchfor = "TX: ATDT" & GlasgowDestination
                If InStr(1, strLine, Searchfor) Then

                    DialGlasgow = DialGlasgow + 1
                    DialGlasgowCount.Text = DialGlasgow

                    For x = 1 To 10 ' loop round after the dial looking for what happens next
                        'should be connect
                        'T<A>
                        'TX of A< message string
                        'T<K>
                        'T<E>
                        '
                        strLine = FileReader.ReadLine

                        If InStr(1, strLine, "CONNECT") Then    'string search 
                            OPTConnectGlasgow = OPTConnectGlasgow + 1
                            OPTConnectGlasgowcount.Text = OPTConnectGlasgow

                            If InStr(1, strLine, "W&W1") Then    'but if ALSO contains reset string then bad
                                OPTGotResetGlasgow = OPTGotResetGlasgow + 1
                                OPTGotResetGlasgowcount.Text = OPTGotResetGlasgow
                                GoTo FoundresultGlasgow
                            End If
                        Else
                            '    ok no reset in here so keep going

                            If InStr(1, strLine, "TX: ATDT") Then 'holy smoke batman getting a new call get out
                                '  MessageBox.Show(" Crashed into next ring  previous call nothing after connect in 10 message loops")
                                GoTo ExitnextLoopGlasgow
                            End If

                            If InStr(1, strLine, "04T<A>") Then 'responce from top end 
                                AlarmRequestGlasgow = AlarmRequestGlasgow + 1
                                OPTalarmRequestGlasgowCount.Text = AlarmRequestGlasgow
                            End If

                            If InStr(1, strLine, "NO CARRIER") Then 'responce from modem 
                                NoCarrierGlasgow = NoCarrierGlasgow + 1
                                OPTNoCarrierGlasgowCount.Text = NoCarrierGlasgow
                                GoTo FoundresultGlasgow
                            End If

                            If InStr(1, strLine, "RX: 04T<K>") Then 'responce from top end 
                                AlarmAckGlasgow = AlarmAckGlasgow + 1
                                OPTAlarmAckGlasgowCount.Text = AlarmAckGlasgow
                                GoTo FoundresultGlasgow
                            End If
                        End If

                    Next x
                End If


ExitnextLoopGlasgow:
FoundresultGlasgow:

                '  Loop


                Searchfor = "TX: ATDT" & DundeeDestination
                If InStr(1, strLine, Searchfor) Then

                    DialDundee = DialDundee + 1
                    DialDundeeCount.Text = DialDundee

                    For x = 1 To 10 ' loop round after the dial looking for what happens next
                        'should be connect
                        'T<A>
                        'TX of A< message string
                        'T<K>
                        'T<E>
                        '
                        strLine = FileReader.ReadLine

                        If InStr(1, strLine, "CONNECT") Then    'string search 
                            OPTConnectDundee = OPTConnectDundee + 1
                            OPTConnectDundeecount.Text = OPTConnectDundee

                            If InStr(1, strLine, "W&W1") Then    'but if ALSO contains reset string then bad
                                OPTGotResetDundee = OPTGotResetDundee + 1
                                OPTGotResetDundeecount.Text = OPTGotResetDundee
                                GoTo FoundresultDundee
                            End If
                        Else
                            '    ok no reset in here so keep going

                            If InStr(1, strLine, "TX: ATDT") Then 'holy smoke batman getting a new call get out
                                '  MessageBox.Show(" Crashed into next ring  previous call nothing after connect in 10 message loops")
                                GoTo ExitnextLoopDundee
                            End If

                            If InStr(1, strLine, "04T<A>") Then 'responce from top end 
                                AlarmRequestDundee = AlarmRequestDundee + 1
                                OPTalarmRequestDundeeCount.Text = AlarmRequestDundee
                            End If

                            If InStr(1, strLine, "NO CARRIER") Then 'responce from modem 
                                NoCarrierDundee = NoCarrierDundee + 1
                                OPTNoCarrierDundeeCount.Text = NoCarrierDundee
                                GoTo FoundresultDundee
                            End If

                            If InStr(1, strLine, "RX: 04T<K>") Then 'responce from top end 
                                AlarmAckDundee = AlarmAckDundee + 1
                                OPTAlarmAckDundeeCount.Text = AlarmAckDundee
                                GoTo FoundresultDundee
                            End If
                        End If

                    Next x
                End If


ExitnextLoopDundee:
FoundresultDundee:

            Loop

            FileReader.Close()

        End If
    End Sub

    Private Sub OPTDebugOpen_Click(sender As Object, e As EventArgs) Handles OPTDebugOpen.Click
        OPTOpenFileDialog_Click()
    End Sub
End Class