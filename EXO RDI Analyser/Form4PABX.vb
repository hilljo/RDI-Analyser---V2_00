Imports System.IO
Public Class Form4PABX
    Dim RDIrowcounterMax As Integer
    Dim PABXrowcounterMax As Integer
    Dim DialingNumber As String
    Dim PABXDuration As String
    Dim PABXTime As String
    Dim howmanymatches As Integer = 0
    Dim PABXBook As Object
    Dim PABXSheet As Object
    Dim RDIBook As Object
    Dim RDISheet As Object
    Dim PABXrowcounter As Integer
    Dim RDIrowcounter As Integer
    Dim rditimeoffsetmax As String = ""
    Dim rditimeoffsetmin As String = ""
    Dim RDITime As String = ""
    Dim PABXExt As String = ""
    Dim PABXExcel As Object
    Dim RDIExcel As Object


    Dim RowofPABXExt As Integer = 2


    Dim PABXExtCheck As String = ""



    Dim lastPABXRow As Long

    Dim RDICommsPort As String = ""
    Dim PABXDataCenter As String = ""

    Dim rditimeMAXforPABXsearch As String = ""
    Dim rditimeMinforPABXsearch As String = ""


    Dim mincallduration As Date = "00:00:30"
    Dim keepLoopAlive As Boolean

    'Private Function SelectFolderPABX() As String
    '    Dim PABXdefaultfolder = driveletter + "RDI Analyser\PABX"
    '    Dim FilesFolder As String

    '    FolderBrowserDialog1.ShowDialog()
    '    FilesFolder = FolderBrowserDialog1.SelectedPath
    '    If FilesFolder = "" Then
    '        FilesFolder = PABXdefaultfolder
    '    End If
    '    SelectedFolderPABX.Text = FilesFolder
    '    Return FilesFolder
    'End Function

    Private Sub Form4PABX_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ClearForm()
        '  SelectedFolderPABX.Text = driveletter + "RDI Analyser\PABX"
        maxrdiNumericUpDown1.Value = 10
    End Sub
    'Private Sub btnSelectFolderPABX_Click(sender As Object, e As EventArgs)
    '    Dim Folder As String
    '    Folder = SelectFolderPABX()
    'End Sub
    Sub OpenPABXLogButton_Click(sender As Object, e As EventArgs) Handles OpenPABXLogButton.Click

        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = driveletter '"h:\"
        openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            SelectedpathPABXlog.Text = openFileDialog1.FileName
        End If

    End Sub





    Sub logmachedcall()
        PABXDuration = PABXSheet.cells(PABXrowcounter, 5).text
        DialingNumber = PABXSheet.cells(PABXrowcounter, 11).text
        PABXTime = PABXSheet.cells(PABXrowcounter, 4).text

        If TimeValue(PABXTime) > TimeValue(rditimeoffsetmin) Then
            If TimeValue(PABXTime) < TimeValue(rditimeoffsetmax) Then
                howmanymatches = howmanymatches + 1 ' we may have >1 so need to put in seperate cols
                RDISheet.cells(RDIrowcounter, 25 + howmanymatches).Value = "PABXTIME " & PABXTime & " RDITime " & RDITime & " Duration " & PABXDuration & " PABXExt " & PABXExt & " Dialingnumber " & DialingNumber

            End If
        End If
    End Sub
    Private Sub collolationINRDIandPABX()
        '    keepLoopAlive = True



        '      Dim maxrdiNumericUpDown1 As New NumericUpDown()

        '     MessageBox.Show(RDIrowcounterMax)
        RDIExcel = CreateObject("Excel.Application")
        PABXExcel = CreateObject("Excel.Application")

        PABXBook = PABXExcel.Workbooks.open(SelectedpathPABXlog.Text)
        PABXSheet = PABXBook.Worksheets("crlx29")

        lastPABXRow = PABXSheet.UsedRange.Rows.Count

        RDIBook = RDIExcel.Workbooks.open(SelectedpathRDIlog.Text)
        RDISheet = RDIBook.Worksheets("DL")

        For RDIrowcounter = 3 To RDIrowcounterMax

            ProgressBarRDILog.Maximum = RDIrowcounterMax - 3
            ProgressBarRDILog.Value = RDIrowcounter - 3

            ProgressRDILabel.Text = ProgressBarRDILog.Value & " Failed Calls Analysed ... " & String.Format("{0:n2}", (RDIrowcounter - 3) / (RDIrowcounterMax - 3)) * 100 & " % Complete"

            Try 'just incase it's shite
                RDITime = RDISheet.cells(RDIrowcounter, 24).text
                'stepps oct 2015 currenty RDI is 24-34Sec faster than PABX
                ' stepps mar 2016 rdi 21-22 sec faster than pabx
                rditimeoffsetmax = Format(CDate(DateAdd("s", +0, RDITime)), "HH: mm:ss")
                rditimeoffsetmin = Format(CDate(DateAdd("s", -50, RDITime)), "HH:mm:ss")
                rditimeMAXforPABXsearch = Format(CDate(DateAdd("n", +10, RDITime)), "HH:mm:ss")
                '     rditimeMinforPABXsearch = Format(CDate(DateAdd("n", -10, RDITime)), "HH:mm:ss")

                RDICommsPort = RDISheet.cells(RDIrowcounter, 21).text
            Catch e As Exception
            End Try



            'TEST' limit loop for test
            PABXrowcounterMax = 1000

            '  PABXrowcounterMax = lastPABXRow
            For PABXrowcounter = 2 To PABXrowcounterMax ' 

                ProgressBarPABXLog.Maximum = PABXrowcounterMax ' / 100
                ProgressBarPABXLog.Value = PABXrowcounter

                ProgressPABXLabel.Text = ProgressBarPABXLog.Value & " PABX Calls Analysed ... " & String.Format("{0:n2}", PABXrowcounter / PABXrowcounterMax) * 100 & " % Complete"

                Try 'just incase it's shite
                    PABXTime = PABXSheet.cells(PABXrowcounter, 4).text
                    If TimeValue(PABXTime) > TimeValue(rditimeMAXforPABXsearch) Then
                        GoTo nextfailure
                    End If
                    '            If TimeValue(PABXTime) < TimeValue(rditimeMinforPABXsearch) Then
                    '            GoTo startagain
                    '            End If
                    PABXExt = PABXSheet.cells(PABXrowcounter, 2).text
                    PABXDataCenter = PABXSheet.cells(PABXrowcounter, 15).text
                Catch e As Exception
                End Try

                If RDICommsPort = "COM01\Port38" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2235" Then
                    Call logmachedcall()
                End If
                If RDICommsPort = "COM01\Port12" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2231" Then
                    Call logmachedcall()
                End If
                If RDICommsPort = "COM03\Port12" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2232" Then
                    Call logmachedcall()
                End If
                If RDICommsPort = "COM05\Port38" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2236" Then
                    Call logmachedcall()
                End If
                If RDICommsPort = "COM05\Port12" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2233" Then
                    Call logmachedcall()
                End If
                If RDICommsPort = "COM07\Port12" And PABXDataCenter = "SW- Bridge Telemetry" And PABXExt = "2234" Then
                    Call logmachedcall()
                End If



            Next PABXrowcounter
            howmanymatches = 0

nextfailure:
        Next RDIrowcounter
        ' save and exit


        RDIBook.Save()

        PABXSheet = Nothing
        PABXBook = Nothing
        RDISheet = Nothing
        RDIBook = Nothing
        RDIExcel.Quit()
        RDIExcel = Nothing
        PABXExcel.Quit()
        PABXExcel = Nothing
        GC.Collect()
    End Sub


    Private Sub EXOPSTN_GSM_Mismatch_Click(sender As Object, e As EventArgs) Handles EXOPSTN_GSM_Mismatch.Click
        collolationINRDIandPABX()
    End Sub

    Private Sub OpenResultsFile_Click(sender As Object, e As EventArgs) Handles OpenResultsFile.Click
        System.Diagnostics.Process.Start(SelectedpathRDIlog.Text)

    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub OpenRDILogButton_Click(sender As Object, e As EventArgs) Handles OpenRDILogButton.Click
        Dim openFileDialog2 As New OpenFileDialog()

        openFileDialog2.InitialDirectory = driveletter '"h: \"
        openFileDialog2.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        openFileDialog2.FilterIndex = 2
        openFileDialog2.RestoreDirectory = True

        If openFileDialog2.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            SelectedpathRDIlog.Text = openFileDialog2.FileName
        End If
    End Sub

    Private Sub maxrdiNumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles maxrdiNumericUpDown1.ValueChanged
        RDIrowcounterMax = Integer.Parse(maxrdiNumericUpDown1.Value) + 3
        ' as calls start on line 3 need to add 3 for progress
        'MessageBox.Show(RDIrowcounterMax)
    End Sub
End Class