Imports System.Data.SqlClient
Public Class PV1
    Public Shared pv1 As Queue(Of String) = New Queue(Of String)
    Public Shared pv1Components As String() = {"Set ID – PV1", "Patient Class", "Assigned Patient Location", "Admission Type", "Pre-admit Number", "Prior Patient Location", "Attending Doctor", "Referring Doctor", "Consulting Doctor", "Hospital Service", "Temporary Location", "Pre-admit Test Indicator", "Re-admission Indicator", "Admit Source", "Ambulatory Status", "VIP Indicator", "Admitting Doctor", "Patient Type", "Visit Number", "Financial Class", "Charge Price Indicator", "Courtesy Code", "Credit Rating", "Contract Code", "Contract Effective Date", "Contract Amount", "Contract Period", "Interest Code", "Transfer to Bad Debt Code", "Transfer to Bad Debt Date", "Bad Debt Agency Code", "Bad Debt Transfer Amount", "Bad Debt Recovery Amount", "Delete Account Indicator", "Delete Account Date", "Discharge Disposition", "Discharged to Location", "Diet Type", "Servicing Facility", "Bed Status", "Account Status", "Pending Location", "Prior Temporary Location", "Admit Date/Time", "Discharge Date/Time", "Current Patient Balance", "Total Charges", "Total Adjustments", "Total Payments", "Alternate Visit ID", "Visit Indicator", "Other Healthcare Provider"}
    Public Shared toBeParsedPv1 As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parsePV1(ByVal message As String, ByVal id As Integer)
        'STEP 1 
        pv1.Clear()
        toBeParsedPv1.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2
        Dim rawPv1 As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawPv1.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            pv1.Enqueue(rawPv1(i))
        Next

        'STEP 3

        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0
        outputConverted("PV1", id, 0, 1, "")

        '1 patient class
        Dim pClass As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'patientClass'")
                sqlCheckTwoCols(con, cmd, pClass)
            End Using
        End Using
        If pClass.Contains(pv1(1)) Then
            outputConverted("PV1", id, 1, 2, pClass(pClass.IndexOf(pv1(2)) + 1))
        Else
            outputError("PV1", id, 1, 2, "Invalid patient class")
        End If

        '2
        outputConverted("PV1", id, 2, 3, "")

        '3 admission type
        Dim adType As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'admissionType'")
                sqlCheckTwoCols(con, cmd, adType)
            End Using
        End Using
        If adType.Contains(pv1(4)) Then
            outputConverted("PV1", id, 3, 4, adType(adType.IndexOf(pv1(4)) + 1))
        Else
            outputError("PV1", id, 3, 4, "Invalid admission type")
        End If

        '4 to 12
        For idx As Integer = 4 To 12
            outputConverted("PV1", id, idx, idx + 1, "")
        Next

        '13 admit source
        Dim adSource As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'admitSource'")
                sqlCheckTwoCols(con, cmd, adSource)
            End Using
        End Using
        If adSource.Contains(pv1(14)) Then
            outputConverted("PV1", id, 13, 14, adSource(adSource.IndexOf(pv1(14)) + 1))
        Else
            outputError("PV1", id, 13, 14, "Invalid admission type")
        End If

        '14 ambulatory source
        Dim amSource As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'ambulatorySource'")
                sqlCheckTwoCols(con, cmd, amSource)
            End Using
        End Using
        If amSource.Contains(pv1(15)) Then
            outputConverted("PV1", id, 14, 15, amSource(amSource.IndexOf(pv1(15)) + 1))
        Else
            outputError("PV1", id, 14, 15, "Invalid ambulatory source")
        End If

        '15 to 34
        For idx As Integer = 15 To 34
            outputConverted("PV1", id, idx, idx + 1, "")
        Next

        '35 discharge disposition
        Dim dd As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'dischargeDisposition'")
                sqlCheckTwoCols(con, cmd, dd)
            End Using
        End Using
        If dd.Contains(pv1(36)) Then
            outputConverted("PV1", id, 35, 36, dd(dd.IndexOf(pv1(36)) + 1))
        Else
            outputError("PV1", id, 35, 36, "Invalid discharge disposition")
        End If
        'stop here left pv1, PID, OBX, OBR
        '36 to 51
        For idx As Integer = 36 To 51
            outputConverted("PV1", id, idx, idx + 1, "")
        Next
    End Function
End Class
