<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form4PABX
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.EXOPSTN_GSM_Mismatch = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenPABXLogButton = New System.Windows.Forms.Button()
        Me.OpenResultsFile = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.OpenRDILogButton = New System.Windows.Forms.Button()
        Me.SelectedpathPABXlog = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.SelectedpathRDIlog = New System.Windows.Forms.TextBox()
        Me.ProgressBarRDILog = New System.Windows.Forms.ProgressBar()
        Me.ProgressRDILabel = New System.Windows.Forms.TextBox()
        Me.maxrdiNumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.ProgressPABXLabel = New System.Windows.Forms.TextBox()
        Me.ProgressBarPABXLog = New System.Windows.Forms.ProgressBar()
        CType(Me.maxrdiNumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'EXOPSTN_GSM_Mismatch
        '
        Me.EXOPSTN_GSM_Mismatch.Location = New System.Drawing.Point(335, 348)
        Me.EXOPSTN_GSM_Mismatch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.EXOPSTN_GSM_Mismatch.Name = "EXOPSTN_GSM_Mismatch"
        Me.EXOPSTN_GSM_Mismatch.Size = New System.Drawing.Size(140, 46)
        Me.EXOPSTN_GSM_Mismatch.TabIndex = 246
        Me.EXOPSTN_GSM_Mismatch.Text = "GO"
        Me.EXOPSTN_GSM_Mismatch.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog1
        '
        '
        'OpenPABXLogButton
        '
        Me.OpenPABXLogButton.Location = New System.Drawing.Point(500, 128)
        Me.OpenPABXLogButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.OpenPABXLogButton.Name = "OpenPABXLogButton"
        Me.OpenPABXLogButton.Size = New System.Drawing.Size(140, 46)
        Me.OpenPABXLogButton.TabIndex = 250
        Me.OpenPABXLogButton.Text = "Select PABX Log"
        Me.OpenPABXLogButton.UseVisualStyleBackColor = True
        '
        'OpenResultsFile
        '
        Me.OpenResultsFile.Location = New System.Drawing.Point(335, 411)
        Me.OpenResultsFile.Name = "OpenResultsFile"
        Me.OpenResultsFile.Size = New System.Drawing.Size(140, 49)
        Me.OpenResultsFile.TabIndex = 252
        Me.OpenResultsFile.Text = "Open Results"
        Me.OpenResultsFile.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'OpenRDILogButton
        '
        Me.OpenRDILogButton.Location = New System.Drawing.Point(123, 128)
        Me.OpenRDILogButton.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.OpenRDILogButton.Name = "OpenRDILogButton"
        Me.OpenRDILogButton.Size = New System.Drawing.Size(140, 46)
        Me.OpenRDILogButton.TabIndex = 253
        Me.OpenRDILogButton.Text = "Select RDI Log"
        Me.OpenRDILogButton.UseVisualStyleBackColor = True
        '
        'SelectedpathPABXlog
        '
        Me.SelectedpathPABXlog.Location = New System.Drawing.Point(383, 178)
        Me.SelectedpathPABXlog.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SelectedpathPABXlog.Name = "SelectedpathPABXlog"
        Me.SelectedpathPABXlog.Size = New System.Drawing.Size(365, 22)
        Me.SelectedpathPABXlog.TabIndex = 254
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "OpenFileDialog1"
        '
        'SelectedpathRDIlog
        '
        Me.SelectedpathRDIlog.Location = New System.Drawing.Point(11, 178)
        Me.SelectedpathRDIlog.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SelectedpathRDIlog.Name = "SelectedpathRDIlog"
        Me.SelectedpathRDIlog.Size = New System.Drawing.Size(365, 22)
        Me.SelectedpathRDIlog.TabIndex = 255
        '
        'ProgressBarRDILog
        '
        Me.ProgressBarRDILog.Location = New System.Drawing.Point(12, 245)
        Me.ProgressBarRDILog.Name = "ProgressBarRDILog"
        Me.ProgressBarRDILog.Size = New System.Drawing.Size(364, 23)
        Me.ProgressBarRDILog.TabIndex = 256
        '
        'ProgressRDILabel
        '
        Me.ProgressRDILabel.Location = New System.Drawing.Point(13, 206)
        Me.ProgressRDILabel.Name = "ProgressRDILabel"
        Me.ProgressRDILabel.Size = New System.Drawing.Size(363, 22)
        Me.ProgressRDILabel.TabIndex = 257
        '
        'maxrdiNumericUpDown1
        '
        Me.maxrdiNumericUpDown1.Location = New System.Drawing.Point(13, 290)
        Me.maxrdiNumericUpDown1.Name = "maxrdiNumericUpDown1"
        Me.maxrdiNumericUpDown1.Size = New System.Drawing.Size(120, 22)
        Me.maxrdiNumericUpDown1.TabIndex = 259
        '
        'ProgressPABXLabel
        '
        Me.ProgressPABXLabel.Location = New System.Drawing.Point(382, 206)
        Me.ProgressPABXLabel.Name = "ProgressPABXLabel"
        Me.ProgressPABXLabel.Size = New System.Drawing.Size(363, 22)
        Me.ProgressPABXLabel.TabIndex = 260
        '
        'ProgressBarPABXLog
        '
        Me.ProgressBarPABXLog.Location = New System.Drawing.Point(384, 245)
        Me.ProgressBarPABXLog.Name = "ProgressBarPABXLog"
        Me.ProgressBarPABXLog.Size = New System.Drawing.Size(364, 23)
        Me.ProgressBarPABXLog.TabIndex = 261
        '
        'Form4PABX
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 517)
        Me.Controls.Add(Me.ProgressBarPABXLog)
        Me.Controls.Add(Me.ProgressPABXLabel)
        Me.Controls.Add(Me.maxrdiNumericUpDown1)
        Me.Controls.Add(Me.ProgressRDILabel)
        Me.Controls.Add(Me.ProgressBarRDILog)
        Me.Controls.Add(Me.SelectedpathRDIlog)
        Me.Controls.Add(Me.SelectedpathPABXlog)
        Me.Controls.Add(Me.OpenRDILogButton)
        Me.Controls.Add(Me.OpenResultsFile)
        Me.Controls.Add(Me.OpenPABXLogButton)
        Me.Controls.Add(Me.EXOPSTN_GSM_Mismatch)
        Me.Name = "Form4PABX"
        Me.Text = "PABX Log"
        CType(Me.maxrdiNumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents EXOPSTN_GSM_Mismatch As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents OpenPABXLogButton As Button
    Friend WithEvents OpenResultsFile As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents OpenRDILogButton As Button
    Friend WithEvents SelectedpathPABXlog As TextBox
    Friend WithEvents OpenFileDialog2 As OpenFileDialog
    Friend WithEvents SelectedpathRDIlog As TextBox
    Friend WithEvents ProgressBarRDILog As ProgressBar
    Friend WithEvents ProgressRDILabel As TextBox
    Friend WithEvents maxrdiNumericUpDown1 As NumericUpDown
    Friend WithEvents ProgressPABXLabel As TextBox
    Friend WithEvents ProgressBarPABXLog As ProgressBar
End Class
