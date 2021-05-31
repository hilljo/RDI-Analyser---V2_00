<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPABX
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormPABX))
        Me.EXOPSTN_GSM_Mismatch = New System.Windows.Forms.Button()
        Me.btnSelectFolderPABX = New System.Windows.Forms.Button()
        Me.DateTimePickerSplitterPABX = New System.Windows.Forms.DateTimePicker()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SelectedFolderPABX = New System.Windows.Forms.TextBox()
        Me.OpenPABXLogButton = New System.Windows.Forms.Button()
        Me.ProgressBarreadPABXLogs = New System.Windows.Forms.ProgressBar()
        Me.ProgressLabel = New System.Windows.Forms.Label()
        Me.Howmany = New System.Windows.Forms.Label()
        Me.FoundCount = New System.Windows.Forms.TextBox()
        Me.OKcount = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PSTNtoGSMStepps = New System.Windows.Forms.TextBox()
        Me.GSMtoPSTNStepps = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GSMtoPSTNDundee = New System.Windows.Forms.TextBox()
        Me.PSTNtoGSMDundee = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Listboxcli = New System.Windows.Forms.ListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'EXOPSTN_GSM_Mismatch
        '
        Me.EXOPSTN_GSM_Mismatch.Location = New System.Drawing.Point(329, 75)
        Me.EXOPSTN_GSM_Mismatch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.EXOPSTN_GSM_Mismatch.Name = "EXOPSTN_GSM_Mismatch"
        Me.EXOPSTN_GSM_Mismatch.Size = New System.Drawing.Size(140, 46)
        Me.EXOPSTN_GSM_Mismatch.TabIndex = 206
        Me.EXOPSTN_GSM_Mismatch.Text = "EXO Mismatch"
        Me.EXOPSTN_GSM_Mismatch.UseVisualStyleBackColor = True
        '
        'btnSelectFolderPABX
        '
        Me.btnSelectFolderPABX.Location = New System.Drawing.Point(458, 11)
        Me.btnSelectFolderPABX.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSelectFolderPABX.Name = "btnSelectFolderPABX"
        Me.btnSelectFolderPABX.Size = New System.Drawing.Size(136, 26)
        Me.btnSelectFolderPABX.TabIndex = 230
        Me.btnSelectFolderPABX.Text = "Select Root Folder"
        Me.btnSelectFolderPABX.UseVisualStyleBackColor = True
        '
        'DateTimePickerSplitterPABX
        '
        Me.DateTimePickerSplitterPABX.Location = New System.Drawing.Point(229, 15)
        Me.DateTimePickerSplitterPABX.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DateTimePickerSplitterPABX.Name = "DateTimePickerSplitterPABX"
        Me.DateTimePickerSplitterPABX.Size = New System.Drawing.Size(178, 22)
        Me.DateTimePickerSplitterPABX.TabIndex = 229
        '
        'FolderBrowserDialog1
        '
        '
        'SelectedFolderPABX
        '
        Me.SelectedFolderPABX.Location = New System.Drawing.Point(229, 41)
        Me.SelectedFolderPABX.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SelectedFolderPABX.Name = "SelectedFolderPABX"
        Me.SelectedFolderPABX.Size = New System.Drawing.Size(365, 22)
        Me.SelectedFolderPABX.TabIndex = 233
        '
        'OpenPABXLogButton
        '
        Me.OpenPABXLogButton.Location = New System.Drawing.Point(329, 184)
        Me.OpenPABXLogButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.OpenPABXLogButton.Name = "OpenPABXLogButton"
        Me.OpenPABXLogButton.Size = New System.Drawing.Size(140, 46)
        Me.OpenPABXLogButton.TabIndex = 244
        Me.OpenPABXLogButton.Text = "Open"
        Me.OpenPABXLogButton.UseVisualStyleBackColor = True
        '
        'ProgressBarreadPABXLogs
        '
        Me.ProgressBarreadPABXLogs.Location = New System.Drawing.Point(204, 150)
        Me.ProgressBarreadPABXLogs.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBarreadPABXLogs.Name = "ProgressBarreadPABXLogs"
        Me.ProgressBarreadPABXLogs.Size = New System.Drawing.Size(390, 27)
        Me.ProgressBarreadPABXLogs.TabIndex = 245
        '
        'ProgressLabel
        '
        Me.ProgressLabel.AutoSize = True
        Me.ProgressLabel.Location = New System.Drawing.Point(367, 129)
        Me.ProgressLabel.Name = "ProgressLabel"
        Me.ProgressLabel.Size = New System.Drawing.Size(65, 17)
        Me.ProgressLabel.TabIndex = 246
        Me.ProgressLabel.Text = "Progress"
        '
        'Howmany
        '
        Me.Howmany.AutoSize = True
        Me.Howmany.BackColor = System.Drawing.Color.Red
        Me.Howmany.Location = New System.Drawing.Point(70, 122)
        Me.Howmany.Name = "Howmany"
        Me.Howmany.Size = New System.Drawing.Size(84, 17)
        Me.Howmany.TabIndex = 247
        Me.Howmany.Text = "Total Found"
        '
        'FoundCount
        '
        Me.FoundCount.Location = New System.Drawing.Point(68, 141)
        Me.FoundCount.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.FoundCount.Name = "FoundCount"
        Me.FoundCount.Size = New System.Drawing.Size(89, 22)
        Me.FoundCount.TabIndex = 248
        '
        'OKcount
        '
        Me.OKcount.Location = New System.Drawing.Point(111, 99)
        Me.OKcount.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.OKcount.Name = "OKcount"
        Me.OKcount.Size = New System.Drawing.Size(89, 22)
        Me.OKcount.TabIndex = 249
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(69, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(28, 17)
        Me.Label1.TabIndex = 250
        Me.Label1.Text = "OK"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(395, 274)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(127, 17)
        Me.Label2.TabIndex = 251
        Me.Label2.Text = "PSTN Dialing GSM"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(528, 274)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 17)
        Me.Label3.TabIndex = 252
        Me.Label3.Text = "GSM Dialing PSTN"
        '
        'PSTNtoGSMStepps
        '
        Me.PSTNtoGSMStepps.Location = New System.Drawing.Point(410, 303)
        Me.PSTNtoGSMStepps.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.PSTNtoGSMStepps.Name = "PSTNtoGSMStepps"
        Me.PSTNtoGSMStepps.Size = New System.Drawing.Size(89, 22)
        Me.PSTNtoGSMStepps.TabIndex = 253
        '
        'GSMtoPSTNStepps
        '
        Me.GSMtoPSTNStepps.Location = New System.Drawing.Point(549, 303)
        Me.GSMtoPSTNStepps.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GSMtoPSTNStepps.Name = "GSMtoPSTNStepps"
        Me.GSMtoPSTNStepps.Size = New System.Drawing.Size(89, 22)
        Me.GSMtoPSTNStepps.TabIndex = 254
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Red
        Me.Label4.Location = New System.Drawing.Point(352, 306)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 17)
        Me.Label4.TabIndex = 255
        Me.Label4.Text = "Stepps"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(352, 334)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 17)
        Me.Label5.TabIndex = 258
        Me.Label5.Text = "Dundee"
        '
        'GSMtoPSTNDundee
        '
        Me.GSMtoPSTNDundee.Location = New System.Drawing.Point(549, 331)
        Me.GSMtoPSTNDundee.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GSMtoPSTNDundee.Name = "GSMtoPSTNDundee"
        Me.GSMtoPSTNDundee.Size = New System.Drawing.Size(89, 22)
        Me.GSMtoPSTNDundee.TabIndex = 257
        '
        'PSTNtoGSMDundee
        '
        Me.PSTNtoGSMDundee.Location = New System.Drawing.Point(410, 331)
        Me.PSTNtoGSMDundee.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.PSTNtoGSMDundee.Name = "PSTNtoGSMDundee"
        Me.PSTNtoGSMDundee.Size = New System.Drawing.Size(89, 22)
        Me.PSTNtoGSMDundee.TabIndex = 256
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Red
        Me.GroupBox1.Controls.Add(Me.Listboxcli)
        Me.GroupBox1.Controls.Add(Me.FoundCount)
        Me.GroupBox1.Controls.Add(Me.Howmany)
        Me.GroupBox1.Location = New System.Drawing.Point(342, 258)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(326, 201)
        Me.GroupBox1.TabIndex = 259
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Errors"
        '
        'Listboxcli
        '
        Me.Listboxcli.FormattingEnabled = True
        Me.Listboxcli.ItemHeight = 16
        Me.Listboxcli.Location = New System.Drawing.Point(200, 111)
        Me.Listboxcli.Name = "Listboxcli"
        Me.Listboxcli.Size = New System.Drawing.Size(120, 84)
        Me.Listboxcli.TabIndex = 263
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox4.Controls.Add(Me.OKcount)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 258)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(326, 201)
        Me.GroupBox4.TabIndex = 261
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Good"
        '
        'FormPABX
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(680, 500)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.GSMtoPSTNDundee)
        Me.Controls.Add(Me.PSTNtoGSMDundee)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GSMtoPSTNStepps)
        Me.Controls.Add(Me.PSTNtoGSMStepps)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ProgressLabel)
        Me.Controls.Add(Me.ProgressBarreadPABXLogs)
        Me.Controls.Add(Me.OpenPABXLogButton)
        Me.Controls.Add(Me.SelectedFolderPABX)
        Me.Controls.Add(Me.btnSelectFolderPABX)
        Me.Controls.Add(Me.DateTimePickerSplitterPABX)
        Me.Controls.Add(Me.EXOPSTN_GSM_Mismatch)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Name = "FormPABX"
        Me.Text = "Mismatch"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents EXOPSTN_GSM_Mismatch As System.Windows.Forms.Button
    Friend WithEvents btnSelectFolderPABX As System.Windows.Forms.Button
    Friend WithEvents DateTimePickerSplitterPABX As System.Windows.Forms.DateTimePicker
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents SelectedFolderPABX As System.Windows.Forms.TextBox
    Friend WithEvents OpenPABXLogButton As System.Windows.Forms.Button
    Friend WithEvents ProgressBarreadPABXLogs As System.Windows.Forms.ProgressBar
    Friend WithEvents ProgressLabel As System.Windows.Forms.Label
    Friend WithEvents Howmany As System.Windows.Forms.Label
    Friend WithEvents FoundCount As System.Windows.Forms.TextBox
    Friend WithEvents OKcount As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents PSTNtoGSMStepps As TextBox
    Friend WithEvents GSMtoPSTNStepps As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents GSMtoPSTNDundee As TextBox
    Friend WithEvents PSTNtoGSMDundee As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Listboxcli As ListBox
End Class
