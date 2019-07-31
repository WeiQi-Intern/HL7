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

        '0 Set ID - R;SI
        If nte(1) = "" Then
            outputError("NTE", id, 0, 1, "Set ID is REQUIRED")
        Else
            dataTypeSIrequired(id, nte(1), "NTE", 0, "Invalid set ID")
        End If

        '1 source of comment - O [excluded]
        outputConverted("NTE", id, 1, 2, "")

        '2 comment - R;FT
        usageRequired(id, nte(3), "NTE", 2, "Comment is REQUIRED")
    End Function
End Class
