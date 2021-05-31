Public Class mydebugs


    Private Sub testaddress()
        Dim myteststring As String
        Dim Searchfor As String
        Dim exoaddress As String
        Dim myposition As Integer
        Dim mystringlength As Integer
        myteststring = "  00:04:34.64     COM26: Tx: 3C6F2810F100A63E (8)"
        Searchfor = "10F100"
        If InStr(1, myteststring, Searchfor) Then
            mystringlength = Len(myteststring)
            myposition = InStr(1, myteststring, Searchfor)

            exoaddress = Mid(myteststring, myposition - 4, 4)

            MessageBox.Show(myposition & myteststring)
            MessageBox.Show(exoaddress)
        End If
    End Sub

End Class


