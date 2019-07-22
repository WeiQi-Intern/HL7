Public Class OBR
    Public Shared obr As Queue(Of String) = New Queue(Of String)
    Public Shared obrComponents As String() = {"Set ID OBR", "Placer Order Number", "Filler Order Number", "Universal Service ID", "Priority – OBR", "Requested Date/Time", "Observation Date/Time", "Observation End Date/Time", "Collection Volume", "Collector Identifier", "Specimen Action Code", "Danger Code", "Relevant Clinical Information", "Specimen Received Date/Time", "Specimen Source", "Ordering Provider", "Order Callback Phone Number", "Placer Field 1", "Placer Field 2", "Filler Field 1", "Filler Field 2", "Results Rpt/Status Chng - Date/Time", "Charge to Practice", "Diagnostic Serv Sect ID", "Result Status", "Parent Result", "Quantity/Timing", "Result Copies To", "Parent", "Transportation Mode", "Reason for Study", "Principal Result Interpreter", "Assistant Result Interpreter", "Technician", "Transcriptionist", "Scheduled Date/Time", "Number of Sample Containers", "Transport Logistics of Collected Sample", "Collector's Comment", "Transport Arrangement Responsibility", "Transport Arranged", "Escort Required", "Planned Patient Transport Comment", "Procedure Code", "Procedure Code Modifier", "Placer Supplemental Service Information", "Filler Supplemental Service Information", "Medically Necessary Duplicate Procedure Reason", "Result Handling"}
    Public Shared toBeParsedObr As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseOBR(ByVal message As String, ByVal id As Integer)
        obr.Clear()
        toBeParsedObr.Clear()
        'parserInterface.DataGridView1.Rows.Clear()
        parserInterface.ComboBox1.Refresh()
        'MessageBox.Show("what's good")
        'MessageBox.Show(message)

        Dim rawObr As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawObr.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            obr.Enqueue(rawObr(i))
        Next
        For idx As Integer = 0 To 48
            Dim outputRow As String() = New String() {id, "OBR", "Observation Reports", obrComponents(idx), obr(idx + 1), "NULL", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NULL"}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            toBeParsedObr.Enqueue(outputRow)
        Next
    End Function
End Class
