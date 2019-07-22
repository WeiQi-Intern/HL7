Public Class EVN
    Public Shared evn As Queue(Of String) = New Queue(Of String)
    Public Shared evnComponents As String() = {"Event Type Code", "Recorded Date/Time", "Date/Time Planned Event", "Event Reason Code", "Operator ID", "Event Occured", "Event Facility"}
    Public Shared toBeParsedEvn As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseEVN(ByVal message As String, ByVal id As Integer)
        evn.Clear()
        toBeParsedEvn.Clear()
        'parserInterface.DataGridView1.Rows.Clear()
        parserInterface.ComboBox1.Refresh()
        'MessageBox.Show("what's good")
        'MessageBox.Show(message)

        Dim rawEvn As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawEvn.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            evn.Enqueue(rawEvn(i))
        Next
        For idx As Integer = 0 To 6
            Dim outputRow As String() = New String() {id, "EVN", "Event Type", evnComponents(idx), evn(idx + 1), "NULL", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NULL"}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            toBeParsedEvn.Enqueue(outputRow)
        Next
    End Function
End Class
