Public Class NTE
    Public Shared nte As Queue(Of String) = New Queue(Of String)
    Public Shared nteComponents As String() = {"Set ID – Notes and Comments", "Source of Comment", "Comment"}
    Public Shared toBeParsedNte As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseNTE(ByVal message As String, ByVal id As Integer)
        'STEP 1: clear history
        nte.Clear()
        toBeParsedNte.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2: split info
        Dim rawNte As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawNte.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            nte.Enqueue(rawNte(i))
        Next

        'No comparison to sql
        For idx As Integer = 0 To 2
            Dim outputRow As String() = New String() {id, "NTE", "Notes and Comments (Observation Specific)", nteComponents(idx), nte(idx + 1), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            toBeParsedNte.Enqueue(outputRow)
        Next
    End Function
End Class
