Imports System.IO
Imports System.Data.Odbc

Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        runsql()


    End Sub

    Sub runsql()
            Dim MyDataAdapter As New Odbc.OdbcDataAdapter

            Dim MyDataSet As New DataSet

        '        Dim MySelect As String = "SELECT name,description,stdref,phonenumber from outstation;"
        '       Dim StrConn As String = "DSN=SWDL;SERVICE=DYNAMICLOGICA:RTRDB1,DYNAMICLOGICB:RTRDB1;UID=SYSTEM;PWD=SYSTTS09"
        Dim MySelect As String = "select description,plantarea from eventhistory where TIMESTAMP >=  NOW () - HOURS (1) AND description like '%Good Incoming Call (TELEM-COM01:COM16)';"

        Dim StrConn As String = "DSN=SWLD;SERVICE=LDA:LD,LDB:LD;UID=SYSTEM;PWD=SYSTTS09"

            Dim MyConn As New Odbc.OdbcConnection(StrConn)

            MyConn.Open()

            MyDataAdapter.SelectCommand = New Odbc.OdbcCommand(MySelect, MyConn)

            MyDataAdapter.Fill(MyDataSet)

            Me.BindingSource1.DataMember = MyDataSet.Tables(0).TableName

            Me.BindingSource1.DataSource = MyDataSet

            Me.DataGridView1.DataSource = Me.BindingSource1

            MyConn.Close()
            MyConn.Dispose()
            MyDataSet.Dispose()
            MyDataAdapter.Dispose()


        End Sub


        Private Sub dataGridView1_CellClick(ByVal sender As Object,
    ByVal e As DataGridViewCellEventArgs) _
    Handles DataGridView1.CellClick
            Try
                Dim selection As String
                selection = (DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                '       MessageBox.Show(selection)
            Catch ex As Exception
            End Try
        End Sub


        Private Sub BindingSource1_CurrentChanged(sender As Object, e As EventArgs) Handles BindingSource1.CurrentChanged

        End Sub


    'Private Sub FilterButton_Click_1(sender As Object, e As EventArgs) Handles FilterButton.Click
    '    '    Me.BindingSource1.Filter = "phonenumber Like'%" & TextBoxfilter.Text & "%'"
    '    Dim filter1 As String
    '    Dim filter2 As String
    '    filter1 = "phonenumber Like'*" & TextBoxfilter.Text & "*'"
    '    filter2 = "stdref Like'%" & TextBoxfilter.Text & "%'"
    '    Me.BindingSource1.Filter = filter1 'Or filter2 '

    '    DataGridView1.Refresh()
    'End Sub
End Class

