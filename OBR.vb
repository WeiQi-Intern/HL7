Public Class OBR
    Public Shared obr As Queue(Of String) = New Queue(Of String)
    Public Shared obrComponents As String() = {"Set ID OBR", "Placer Order Number", "Filler Order Number", "Universal Service ID", "Priority – OBR", "Requested Date/Time", "Observation Date/Time", "Observation End Date/Time", "Collection Volume", "Collector Identifier", "Specimen Action Code", "Danger Code", "Relevant Clinical Information", "Specimen Received Date/Time", "Specimen Source", "Ordering Provider", "Order Callback Phone Number", "Placer Field 1", "Placer Field 2", "Filler Field 1", "Filler Field 2", "Results Rpt/Status Chng - Date/Time", "Charge to Practice", "Diagnostic Serv Sect ID", "Result Status", "Parent Result", "Quantity/Timing", "Result Copies To", "Parent", "Transportation Mode", "Reason for Study", "Principal Result Interpreter", "Assistant Result Interpreter", "Technician", "Transcriptionist", "Scheduled Date/Time", "Number of Sample Containers", "Transport Logistics of Collected Sample", "Collector's Comment", "Transport Arrangement Responsibility", "Transport Arranged", "Escort Required", "Planned Patient Transport Comment", "Procedure Code", "Procedure Code Modifier", "Placer Supplemental Service Information", "Filler Supplemental Service Information", "Medically Necessary Duplicate Procedure Reason", "Result Handling"}
    Public Shared toBeParsedObr As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseOBR(ByVal message As String, ByVal id As Integer)
        'step1
        obr.Clear()
        toBeParsedObr.Clear()
        parserInterface.ComboBox1.Refresh()

        'step2
        Dim rawObr As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawObr.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            obr.Enqueue(rawObr(i))
        Next

        'step3
        '0 set ID - OBR - O;SI
        dataTypeSIoptional(id, obr(1), "OBR", 0, "Invalid set ID")

        '1 placer order number ~ 5 requested date/time - O
        For idx As Integer = 1 To 5
            outputConverted("OBR", id, idx, idx + 1, "")
        Next

        '6 observation date/time - O;DTM
        checkDTMoptional(id, obr(7), "OBR", 6)

        '7 obervation end date/time ~ 48 result handling - O
        For idx As Integer = 7 To 48
            outputConverted("OBR", id, idx, idx + 1, "")
        Next
    End Function
End Class
