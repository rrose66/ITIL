Attribute VB_Name = "Development"
Option Compare Database
Option Explicit
Sub main()
    Dim fso As Object 'File System Object
    Set fso = CreateObject("scripting.filesystemobject")
'    Set MYdb = CurrentDb
'    Set UALdb = OpenDatabase("C:\Users\RROSE66\Documents\ITIL\Wiz_User_Action_Log.accdb", 0)
'    Set dbUAL = OpenDatabase("C:\Users\RROSE66\Documents\ITIL\Wiz_UALs.accdb", 0)

    Dim dtCurrent As Date
    Dim dtLast As Date
    
    Dim intSR As Integer
    Dim intSecondsLeft As Double
    Dim lngDurSec As Long
    Dim intFC As Double 'File Counter
    Dim intLC As Integer 'Last Column
    Dim intRC As Integer 'Record Counter
    
    Dim strDays As String * 3
    Dim strDur As String
    Dim strWV As String * 8
    Dim strHours As String * 2
    Dim strMinutes As String * 2
    Dim strSeconds As String * 2
    Dim strTempDate As String
    
    Dim rstSource As Recordset
    Dim rstWiz_UAL_Date As DAO.Recordset
    Set rstWiz_UAL_Date = dbUAL.OpenRecordset("Wiz_UAL_Date", dbOpenTable)
    
    Set xlsWB = xlsApp.ActiveWorkbook
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    Set rstSource = CurrentDb.OpenRecordset("NeededUALs")
    rstSource.MoveFirst
    While Not rstSource.EOF
        Set xlsWB = xlsApp.Workbooks.Open("C:\Users\RROSE66\Documents\Archives\Import_Files\Source\" & rstSource.Fields("FileName").Value)
        xlsWB.UpdateLinks = xlUpdateLinksNever
        Set xlsWS = xlsWB.Worksheets("User Action Log")
         'Now load the User Action Log
'        Set xlsWS = xlsWB.Worksheets("User Action Log")
        xlsWS.Select
        strWV = xlsWS.Cells(1, 1).Value
        intER = 1
        Dim lngER As Long
        lngER = 1
        Do Until IsEmpty(xlsWS.Range("A" & lngER, "A" & lngER).Value)
            lngER = lngER + 1
        Loop
        lngER = lngER - 1
        Dim lngUALRC As Long
        lngUALRC = lngER
        'Format(dTime, "hh:mm:ss")
        For intSR = 3 To lngER
            If Mid(xlsWS.Cells(intSR, 4).Value, 1, 5) = "CDSID" Then
                'Stop
                rstWiz_UAL_Date.Index = "CDSID"
                rstWiz_UAL_Date.Seek "=", Trim(Mid(xlsWS.Cells(intSR, 4).Value, 8, 10))
                If rstWiz_UAL_Date.NoMatch Then
                    'Stop
                    strCMD = ""
'                    strCMD = "insert into Wiz_UAL_Date (ID, UAL_DateTime, CDSID, DateIssued, FileTimeStamp, TextDateFormat) select "
                    strCMD = "insert into Wiz_UAL_Date (ID, UAL_DateTime, CDSID, DateIssued, FileTimeStamp) select "
                    strCMD = strCMD & GetNextUAL_Date_ID()
                    strCMD = strCMD & ", " & Chr(34) & xlsWS.Cells(intSR, 1).Value & Chr(34)
                    strCMD = strCMD & ", " & Chr(34) & Mid(xlsWS.Cells(intSR, 4).Value, 8, 10) & Chr(34)
                    strCMD = strCMD & ", #" & GetDateIssued(xlsWB.Name) & "#"
                    strCMD = strCMD & ", #" & GetFileTimeStamp("C:\Users\RROSE66\Documents\Archives\Import_Files\Source\" & xlsWB.Name) & "#"
'                    strCMD = strCMD & ", " & Chr(34) & GetTextDateFormat(strCDSID, strUAL_DateTime, FileTimeStamp) & Chr(34)
                    dbUAL.Execute (strCMD)
                End If
            End If
'            'convert text to date
'            '1 '11-Aug-15_11:29:08 AM
'            '2 '21.Aug.2015_10:03:48
'            '3 '07.Sep.2015_12:01:08 PM
'            '4 '12.08.2015_11:29:05 ddmmyyyy done
'            '5 '19.8.2015_10:46:28 ddmyyyy
'            '6 '9.9.15_13:30:23
'            '7 '10/8/2015_PM 4:55:09 ddmyyyy all these dates are not getting trapped here
'            '8 '11/08/2015_02:36:09 p.m. ddmmyyyy done
'            '            'target format is mmddyyyy hhmmss
'            '
'            '1 will always have 2 dashes
'            '2 will always have 2 dots and 3 characters of text but no AM or PM
'            '3 will always have 2 dots and 3 characters of text and AM or PM
'            '
'            'all date with no characters should first be reformatted as two digit month and day and 4 digit year
'            'To determine month and day position of dates with 2 dots, reference the file timestamp
'            'target format is mmddyyyy hhmmss
'            If InStr(1, xlsWS.Range("A" & intSR).Value, "/") > 0 And InStr(1, xlsWS.Range("A" & intSR).Value, ".") = 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "_AM") = 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "_PM") = 0 Then
'                dtCurrent = Replace(xlsWS.Range("A" & intSR).Value, "_", " ")
'            'still need to develop
'    '            ElseIf (InStr(1, xlsWS.Range("A" & intSR).Value, "/") = 0 And InStr(1, xlsWS.Range("A" & intSR).Value, ".") = 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "_AM") > 0) Or (InStr(1, xlsWS.Range("A" & intSR).Value, "/") > 0 And InStr(1, xlsWS.Range("A" & intSR).Value, ".") = 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "_PM") > 0) Then
'    '            '10/8/2015_PM 4:55:09
'    '                strCMD = xlsWS.Range("A" & intSR).Value
'    '                strCMD = IIf(InStr(4, strCMD, "/") = 5, Mid(strCMD, 1, 3) & "0" & Mid(strCMD, 4, 20), strCMD)
'    '                strCMD = Replace(strCMD, "_", " ")
'    '                dtCurrent = IIf(InStr(xlsWS.Range("A" & intSR).Value, "_AM") > 0, Replace(Replace(xlsWS.Range("A" & intSR).Value, "_", " "), "_AM", ""), Replace(Replace(xlsWS.Range("A" & intSR).Value, "_", " "), "_PM", ""))
'            ElseIf InStr(xlsWS.Range("A" & intSR).Value, "-") > 0 And IsNumeric(Mid(xlsWS.Range("A" & intSR).Value, 4, 1)) Then
'                dtCurrent = Mid(xlsWS.Range("A" & intSR).Value, 6, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 9, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 4) & " " & Mid(xlsWS.Range("A" & intSR).Value, 12, 8)
'            ElseIf InStr(xlsWS.Range("A" & intSR).Value, ".") > 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "/") > 0 Then
'                '11/08/2015_02:36:09 p.m.
'                dtCurrent = Replace(Replace(xlsWS.Range("A" & intSR).Value, "_", " "), ".", "")
'            ElseIf InStr(xlsWS.Range("A" & intSR).Value, ".") > 0 And InStr(1, xlsWS.Range("A" & intSR).Value, "/") = 0 Then
'                '12.08.2015_11:29:05
'                dtCurrent = Mid(xlsWS.Range("A" & intSR).Value, 4, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 7, 4) & " " & Mid(xlsWS.Range("A" & intSR).Value, 12, 8)
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Jan" Then
'                dtCurrent = CDate("01/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Feb" Then
'                dtCurrent = CDate("02/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Mar" Then
'                dtCurrent = CDate("03/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Apr" Then
'                dtCurrent = CDate("04/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "May" Then
'                dtCurrent = CDate("05/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Jun" Then
'                dtCurrent = CDate("06/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Jul" Then
'                dtCurrent = CDate("07/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Aug" Then
'                dtCurrent = CDate("08/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Sep" Then
'                dtCurrent = CDate("09/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Oct" Then
'                dtCurrent = CDate("10/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Nov" Then
'                dtCurrent = CDate("11/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            ElseIf Mid(xlsWS.Range("A" & intSR).Value, 4, 3) = "Dec" Then
'                dtCurrent = CDate("12/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 8, 2) & " " & Mid(xlsWS.Range("A" & intSR).Value, 11, 11))
'            Else
'                dtCurrent = Mid(xlsWS.Range("A" & intSR).Value, 4, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 1, 2) & "/" & Mid(xlsWS.Range("A" & intSR).Value, 7, 4) & " " & Mid(xlsWS.Range("A" & intSR).Value, 12, 8)
'            End If
'            'calculate and format the duration
'            'duration is based on days:hours:minutes:seconds
'            'the initial calculation is in seconds
'            'days is simply seconds \ (60 seconds in a minute * 60 minutes in a hour * 24 hours in a day) or 86400
'            'hours minutes and seconds those units of time left over from each calculation
'            If intSR = 3 Then
'                'the first value will always be this since it takes two action log entries to calculate a duration
'                strDur = "000:00:00:00"
'            Else
'                'Calculate the total seconds taken
'                strDur = Abs(DateDiff("s", dtCurrent, dtLast))
'                intSecondsLeft = strDur
'                lngDurSec = strDur
'                'Calculate the rounded days taken
'                strDays = Round(intSecondsLeft / 86400)
'                'now format the string with padded zeroes
'                If strDays < 99 And strDays > 9 Then
'                    strDays = "0" & strDays
'                ElseIf strDays < 10 Then
'                    strDays = "00" & Trim(strDays)
'                End If
'                intSecondsLeft = intSecondsLeft - (strDays * 86400)
'                'convert left over seconds to hours
'                strHours = Round(intSecondsLeft / (60 * 60))
'                intSecondsLeft = intSecondsLeft - (strHours * 60 * 60)
'                'now format the string with padded zeroes
'                If strHours < 10 Then
'                    strHours = "0" & Trim(strHours)
'                End If
'                'convert remaining seconds to minutes left
'                strMinutes = Abs(Round(intSecondsLeft / 60))
'                If strMinutes < 10 Then
'                    strMinutes = "0" & Trim(strMinutes)
'                Else
'                    strMinutes = Trim(strMinutes)
'                End If
'                'calculate seconds left
'                intSecondsLeft = intSecondsLeft - (strMinutes * 60)
'                strSeconds = Abs(intSecondsLeft)
'                If strSeconds < 10 Then
'                    strSeconds = "0" & Trim(strSeconds)
'                Else
'                    strSeconds = Trim(strSeconds)
'                End If
'                strDur = Trim(strDays) & ":" & Trim(strHours) & ":" & Trim(strMinutes) & ":" & Trim(strSeconds)
'            End If
'             'update the date table
'            Stop
'            'now apply the new data model
'            'first find the ID
'            'this will be the ID with a subject of CTQ Analaysis and no related request or incident or PBI or Rally Item
'            'Stop 'not tested
'            strCMD = ""
'            strCMD = "insert into Wiz_User_Action_Log ("
'            strCMD = strCMD & "XLS_Row, AI_Calc, AI_Done, Action, Location, Field_Value, Wiz_ID, AI_Duration, AI_Dur_Sec) "
'            strCMD = strCMD & "select "
'            strCMD = strCMD & intSR
'            strCMD = strCMD & ", " & Chr(34) & dtLast & Chr(34)
'            strCMD = strCMD & ", " & Chr(34) & xlsWS.Range("A" & intSR).Value & Chr(34)
'            strCMD = strCMD & ", " & Chr(34) & xlsWS.Range("B" & intSR).Value & Chr(34)
'            strCMD = strCMD & ", " & Chr(34) & Replace(xlsWS.Range("C" & intSR).Value, Chr(34), "'") & Chr(34)
'            strCMD = strCMD & ", " & Chr(34) & Replace(xlsWS.Range("D" & intSR).Value, Chr(34), "'") & Chr(34)
'            strCMD = strCMD & ", " & rstSource.Fields(0).Value
'            strCMD = strCMD & ", " & Chr(34) & strDur & Chr(34)
'            strCMD = strCMD & ", " & lngDurSec
'            UALdb.Execute (strCMD)
'            dtLast = dtCurrent
        Next intSR
         'now normalize
'         strCMD = ""
'         strCMD = strCMD & "insert into Wiz_UALs (Wiz_ID) select " & rstSource.Fields(0).Value
'         ITILdb.Execute (strCMD)
         xlsWB.Close
        Set xlsWB = Nothing
        rstSource.MoveNext
    Wend
End Sub

Function GetNextUAL_Date_ID() As Long
    Set rstQ = Nothing
    Set rstQ = CurrentDb.OpenRecordset("NextUAL_Date_ID")
    GetNextUAL_Date_ID = rstQ.Fields(0).Value
    rstQ.Close
    Set rstQ = Nothing
End Function
Function GetDateIssued(FileName As String) As Date
    Set rstQ = Nothing
    strCMD = ""
    strCMD = strCMD & "SELECT ITIL_Wiz_Tabs.WorkSheet_Name "
    strCMD = strCMD & "FROM (ITIL_Wiz_IDs INNER JOIN ITIL_Wiz_Tabs ON ITIL_Wiz_IDs.Wiz_ID = ITIL_Wiz_Tabs.Wiz_ID) "
    strCMD = strCMD & "INNER JOIN ITIL_Wiz_File_Names ON ITIL_Wiz_IDs.FN_ID = ITIL_Wiz_File_Names.FN_ID "
    strCMD = strCMD & "where FileName = " & Chr(34) & FileName & Chr(34)
    strCMD = strCMD & " and Worksheet_name = " & Chr(34) & "ID-Overall" & Chr(34) & ";"
    Set rstQ = CurrentDb.OpenRecordset(strCMD)
    If rstQ.RecordCount > 0 Then
        GetDateIssued = xlsWB.Worksheets("ID-Overall").Cells(1, 5).Value
    Else
        GetDateIssued = Null
        'Stop 'need to find another reliable source for the date issued
    End If
End Function
Function GetFileTimeStamp(FileName As String) As Date
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileName)
    GetFileTimeStamp = f.DateCreated
End Function
Function GetTextDateFormat(strCDSID As String, strUAL_DateTime As String, FileTimeStamp As Date) As String
    
End Function
