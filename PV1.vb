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

        '0 Set ID - PID - O;SI
        dataTypeSIoptional(id, pv1(1), "PV1", 0, "Invalid set ID")

        '1 Patient class - O;IS
        Dim pClass As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'patientClass'")
                sqlCheckTwoCols(con, cmd, pClass)
            End Using
        End Using
        If pClass.Contains(pv1(2)) Then
            outputConverted("PV1", id, 1, 2, pClass(pClass.IndexOf(pv1(2)) + 1))
        Else
            outputError("PV1", id, 1, 2, "Invalid patient class")
        End If

        '2 Assigned patient location - O;PL
        outputConverted("PV1", id, 2, 3, "")

        '3 Admission type - O;IS
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

        '4 Pre-admit number ~ 12 Re-admission indicator - O
        For idx As Integer = 4 To 12
            outputConverted("PV1", id, idx, idx + 1, "")
        Next

        '13 Admit source - O;IS
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

        '14 Ambulatory status - O;IS
        Dim amSource As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'ambulatoryStatus'")
                sqlCheckTwoCols(con, cmd, amSource)
            End Using
        End Using
        If amSource.Contains(pv1(15)) Then
            outputConverted("PV1", id, 14, 15, amSource(amSource.IndexOf(pv1(15)) + 1))
        Else
            outputError("PV1", id, 14, 15, "Invalid ambulatory source")
        End If

        '15 VIP indicator ~ 34 Deleted account date - O
        For idx As Integer = 15 To 34
            outputConverted("PV1", id, idx, idx + 1, "")
        Next

        '35 Discharge disposition - O;IS
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

        '36 Discharged to Location ~ 42 Prior Temporary Location - O
        For idx As Integer = 36 To 42
            outputConverted("PV1", id, idx, idx + 1, "")
        Next

        '43 Admit date/time ~ 44 Discharge date/time - O
        For idx As Integer = 43 To 44
            checkDTMoptional(id, pv1(idx + 1), "PV1", idx)
        Next

        '45 Current patient balance - O
        outputConverted("PV1", id, 45, 46, "")

        '46 Total charges - O;NM
        dataTypeNM(id, pv1(47), "PV1", 46, "Invalid total charges")

        '47 Total adjustments ~ 51 Other healthcare provider - O
        For idx As Integer = 47 To 51
            outputConverted("PV1", id, idx, idx + 1, "")
        Next
    End Function
End Class
