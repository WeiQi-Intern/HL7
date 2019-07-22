Public Class OBX
    Public Shared obx As Queue(Of String) = New Queue(Of String)
    Public Shared obxComponents As String() = {"Set ID – OBX", "Value Type", "Observation Identifier", "Observation Sub-ID", "Observation Value", "Units", "References Range", "Abnormal Flags", "Probability", "Nature of Abnormal Test", "Observation Result Status", "Effective Date of Reference Range", "User Defined Access Checks", "Date/Time of the Observation", "Producer's ID", "Responsible Observer", "Observation Method", "Equipment Instance Identifier", "Date/Time of the Analysis"}
    Public Shared toBeParsedObx As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseOBX(ByVal message As String, ByVal id As Integer)
        obx.Clear()
        toBeParsedObx.Clear()
        'parserInterface.DataGridView1.Rows.Clear()
        parserInterface.ComboBox1.Refresh()
        'MessageBox.Show("what's good")
        'MessageBox.Show(message)

        Dim rawObx As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawObx.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            obx.Enqueue(rawObx(i))
        Next
        For idx As Integer = 0 To 18
            Dim outputRow As String() = New String() {id, "OBX", "Observation", obxComponents(idx), obx(idx + 1), "NULL", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NULL"}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            toBeParsedObx.Enqueue(outputRow)
        Next
    End Function
End Class
