Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports VB = Microsoft.VisualBasic
Public Class Form1
    ' code v.1.2.9.d3-local and remote PC by Loukas Samaras (for Intel and AMD)
    ' version changes
    ' 1. parallel for decoding
    ' 2. avcoid loading previous tweets on listbox
    Public Property uu As String
    Public cancelFlag As Boolean = False
    Public ClickTimes = 0
    Public Listbox2 As ListBox
    Public CurrentMinute
    Public LastHour, LastHourInt
    Public starttime, endtime, interval
    Public starttime1, starttime2, starttime3, starttime4, starttime5, starttime6, starttime7, starttime8, starttime9, starttime10
    Public endtime1, endtime2, endtime3, endtime4, endtime5, endtime6, endtime7, endtime8, endtime9, endtime10
    Public interval1, interval2, interval3, interval4, interval5, interval6, interval7, interval8, interval9, interval10
    Public Path1, path2, path3, path4, path5 As String
    Public counter = 0
    Public button1Times As Integer = 0
    Public crowsave2 As Integer = 0
    Public TotalTweets As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load '
        CheckForIllegalCrossThreadCalls = False ' all threads can access the User Interface (form
        'Path1 = "D:\Python\OutputStreaming.txt"
        With Chart1 ' set chart initial values
            .Titles.Add("weekly Influenza estimation")
            .Series(0).Name = "Influenza"
            .ChartAreas(0).AxisX.MajorGrid.Enabled = False
            .ChartAreas(0).AxisY.MajorGrid.Enabled = False
            .ChartAreas(0).AxisX.MinorGrid.Enabled = False
            .ChartAreas(0).AxisY.MinorGrid.Enabled = False
            .Series.Add(1)
            .Series(1).Name = "Prediction"
            .Series(1).Color = Color.Red
            .Series(1).ChartType = SeriesChartType.Line
        End With
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Source.Text = "Remote" Then
            Path1 = "\\WINSERVER\D-Server\Python\OutputStreaming.txt"
        Else
            Path1 = "D:\Python\OutputStreaming.txt"
        End If
        'Task 1. set imnitial values---------------
        Dim startTime, endTime, interval
        button1Times = button1Times + 1
        If button1Times > 2 Then GoTo UpdateTask
        startTime = Now
        cancelFlag = False
        Button2.Select()
        counter = 0 : crow = 0 : cLine = 0
        'ListBox1.Items.Clear()
        Button2.Select()
        endtime1 = Now
Start_results:
        'Dim th As New Thread(AddressOf ReadFile)
        'th.Start()
        'th.Join(1)

        starttime2 = Now
        'Task 2:read the file once and display data-----------------------------
        Call ReadFile2()
        endtime2 = Now

        starttime3 = Now
        ' task 3. decode -------------------------------------------------------
        Call ModuleDecoding.decoder()
        endtime3 = Now

        starttime4 = Now
        ' task 4.write translated files-----------------------------------------
        Call ModuleWriteFiles.WriteFile()
        endtime4 = Now

        starttime5 = Now
        ' Task 5. Populate Listbox1---------------------------------------------
        TotalTweets = crow
        'Parallel.For(0, TotalTweets, Sub(ii)
        If CheckBoxShowAll.Checked = True Then
            For ii = 1 To crow
                ListBox1.Items.Add(ii & Chr(9) & data(ii, 1) & Chr(9) & data(ii, 2) & Chr(9) & data(ii, 3)) ' show data in listbox
            Next    '--------------------------------
        End If
        'End Sub)
        endtime5 = Now

                               starttime6 = Now
                               ' Task 6. hold first and last dates/time--------------------------------
                               LastHour = Mid(data(crow, 1), 12, 2) : If LastHour = "" Then LastHour = 0
                               LastHourInt = CInt(LastHour)
                               If crow > 0 Then ' if lines found
                                   StartYear = Mid(data(1, 1), 1, 4) - 2018
                                   StartMonth = Mid(data(1, 1), 6, 2)
                                   StartDay = CInt(Mid(data(1, 1), 9, 2))
                                   StartWeek = 1 + Int((30.4 * (StartMonth - 1) + StartDay) / 7.019)
                                   StartHour = CInt(Mid(data(1, 1), 12, 2))
                                   StartMinute = CInt(Mid(data(1, 1), 15, 2))
                                   EndYear = Mid(data(crow, 1), 1, 4) - 2018
                                   EndMonth = CInt(Mid(data(crow, 1), 6, 2))
                                   EndDay = CInt(Mid(data(crow, 1), 9, 2))
                                   EndWeek = 1 + Int((30.4 * (EndMonth - 1) + EndDay) / 7.019)
                                   EndHour = CInt(Mid(data(crow, 1), 12, 2))
                                   EndMinute = CInt(Mid(data(crow, 1), 15, 2))
                               Else
                                   StartYear = 0
                                   StartMonth = 0
                                   StartDay = 0
                                   StartWeek = 0
                                   StartHour = 0
                                   EndYear = 0
                                   EndMonth = 0
                                   EndDay = 0
                                   EndWeek = 0
                                   EndHour = 0
                                   EndMinute = 0
                               End If
                               'store values to textboxes
                               StartYearBox.Text = StartYear
                               StartMonthBox.Text = StartMonth.ToString
                               StartWeekBox.Text = StartWeek.ToString
                               StartDayBox.Text = StartDay.ToString
                               StartHourBox.Text = StartHour.ToString
                               If StartMinute Is Nothing Then StartMinute = 0
                               StartMinuteBox.Text = StartMinute.ToString
                               EndYearBox.Text = EndYear
                               EndMonthBox.Text = EndMonth.ToString
                               'EndWeekBox.Text = EndWeek.ToString
                               EndDayBox.Text = EndDay.ToString
                               EndHourBox.Text = EndHour.ToString
                               EndMinuteBox.Text = EndMinute.ToString

                               endtime6 = Now
                               starttime7 = Now
                               'Task 7. Call Module stats-----------------------------
                               Call ModuleStats.Calculations()
                               endtime7 = Now

                               starttime8 = Now
                               'Task 8. Populate Listbox2-----------------------------
                               Call PopulateListbox2()
                               endtime8 = Now

                               starttime9 = Now
                               'Task 9. fill chart------------------------------------
                               Call ChartFill()
                               endtime9 = Now

                               starttime10 = Now
                               ' Task 10. store values and set row and line vars as zero------------------
                               crowSave = crow : cLineSave = cLine
                               counter = counter + 1
                               endtime10 = Now

                               TextBoxRead.Text = counter
                               TextBoxTweets.Text = crow
                               endTime = Now
                               interval = (endTime - startTime).totalseconds
                               TextBoxLoadTime.Text = interval.ToString
                               ' end of load procedures
                               ' calculate time of each task (in sec)
                               interval1 = (endtime1 - startTime).totalseconds : TextBoxTask1.Text = interval1.ToString
                               interval2 = (endtime2 - starttime2).totalseconds : TextBoxTask2.Text = interval2.ToString
                               interval3 = (endtime3 - starttime3).totalseconds : TextBoxTask3.Text = interval3.ToString
                               interval4 = (endtime4 - starttime4).totalseconds : TextBoxTask4.Text = interval4.ToString
                               interval5 = (endtime5 - starttime5).totalseconds : TextBoxTask5.Text = interval5.ToString
                               interval6 = (endtime6 - starttime6).totalseconds : TextBoxTask6.Text = interval6.ToString
                               interval7 = (endtime7 - starttime7).totalseconds : TextBoxTask7.Text = interval7.ToString
                               interval8 = (endtime8 - starttime8).totalseconds : TextBoxTask8.Text = interval8.ToString
                               interval9 = (endtime9 - starttime9).totalseconds : TextBoxTask9.Text = interval9.ToString
                               interval10 = (endtime10 - starttime10).totalseconds : TextBoxTask10.Text = interval10.ToString
UpdateTask:

                               ' Task 11. Update data continously-------------------------------------------
                               Call UpdateRead()

                           End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cancelFlag = True
        MsgBox("Program stopped")
        'MsgBox("crow=" & crow & "crowsave=" & crowSave & "crowsave2" & crowsave2)
        total = 0
        Button1.Select()
        'crowSave = crowsave2
        crowSave = crowsave2
    End Sub

    Sub PopulateListbox2()
        If Listbox2.Items.Count > 0 Then
            Listbox2.Items.Clear()
            Listbox2.Refresh()
        End If
        'delete all rows
        total = 0
        Listbox2.Items.Add("Year" & Chr(9) & "Month" & Chr(9) & "Week" & Chr(9) & "Day" & Chr(9) & "Count" & Chr(9) & "Totall") ' show column labels in listbox
        ' add stats to listbox2
        For i = 0 To EndYear ' for every year
            For k = 1 To 12 ' for every month
                For m = 1 To 31 ' for every day
                    If dataDay(i, k, m) > 0 Then ' show counts >1
                        total = total + dataDay(i, k, m)
                        Listbox2.Items.Add(i + 2018 & Chr(9) & k & Chr(9) & WeekCalculation(i, k, m) & Chr(9) & m & Chr(9) _
                                & dataDay(i, k, m) & Chr(9) & total) ' show data in listbox
                    End If
                Next m
            Next k
        Next i
        ' focus on the last tweet in listbox1
        'Listbox2.SelectedIndex = Listbox2.Items.Count - 1
        'ListBox1.Focus()
    End Sub
    Sub ChartFill()
        With Chart1
            .Series(0).Points.Clear()
            .Series(0).Name = "Influenza"
            .ChartAreas(0).AxisX.MajorTickMark.Enabled = False
            .ChartAreas(0).AxisX.Title = "weeks"
            .ChartAreas(0).AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None
        End With
        Chart1.ChartAreas(0).AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None

        Dim i As Integer = TimeYear  ' current year
        Dim k As Integer = TimeMonth ' current month

        For i = 1 To TotalWeeks + 1 ' for every week
            Chart1.Series(0).Points.AddXY(StatsWeek(i, 11), Prediction(i))
        Next i
        Dim PointCount = 0
        For Each s In Chart1.Series
            For Each p As DataPoint In s.Points
                PointCount = PointCount + 1
                If PointCount > TotalWeeks - 1 Then
                    p.Color = Color.Red
                End If
            Next
        Next
    End Sub
    Public Sub wait(ByVal ms As Integer)
        Using wh As New ManualResetEvent(False)
            wh.WaitOne(ms)
        End Using
    End Sub
    Sub ReadFile()
        'Dim startTime, endTime, interval
        'starttime = Now
        crow = 0 : cLine = 0
        Using MyReader As New FileIO.TextFieldParser(Path1)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    crow = crow + 1

                    Dim currentField As String
                    For Each currentField In currentRow
                        cLine = cLine + 1
                        data(crow, cLine) = currentField
                    Next
                    cLine = 0

                Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
            MyReader.Close()
        End Using
    End Sub
    Sub ReadFile2()
        'Dim startTime, endTime, interval
        'starttime = Now
        crow = 0
        Dim Lines() As String = IO.File.ReadAllLines(Path1)

        'MsgBox(Lines(1).ToString)
        For Each line As String In Lines
            If line <> "" Then
                crow += 1
                Dim strArray As String() = line.Split(",") ' split fields by tab (chr(9)
                For i = 0 To 2
                    data(crow, i + 1) = strArray(i).ToString
                Next
            End If
        Next
    End Sub
    Sub UpdateRead()
        While cancelFlag = False ' execute reader continiuously if button start is pressed
            crow = 0 : cLine = 0
            'Application.DoEvents()
            Button2.Select()
            'If interval >= 10 Then ' do not wait 10 seconds to reduce cpu usage and temparatures
            starttime = DateTime.Now()
            counter = counter + 1 ' how many reads
            TextBoxRead.Text = counter
            Using MyReader As New FileIO.TextFieldParser(Path1)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        crow = crow + 1
                        'If cRow = cRowSave Then GoTo notWrite

                        Dim currentField As String
                        For Each currentField In currentRow
                            cLine = cLine + 1
                            data(crow, cLine) = currentField

                        Next
                        cLine = 0
                        Call ModuleDecoding.decoder()
                    Catch ex As Microsoft.VisualBasic.
                                    FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                    Application.DoEvents()
                End While
                MyReader.Close()
            End Using
            'total = 0
            Dim LastMinute ' store last minute
            LastMinute = Mid(data(crowSave, 1), 15, 2)
            LastHour = Mid(data(crow, 1), 12, 2) : If LastHour = "" Then LastHour = 0
            LastHourInt = CInt(LastHour) ' *************************************

            'check if file has changed
            If crowSave <> crow Then ' if new rows added
                'add to the list box the new rows
                For i = crowSave + 1 To crow
                    ListBox1.Items.Add(i & Chr(9) & data(i, 1) & Chr(9) & data(i, 2) & Chr(9) & data(i, 3))
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    Call ModuleStats.Calculations()
                    'Call ModuleStats.CalculationTotal() ' if new entry found, recalculate stats
                    Call PopulateListbox2()
                    Call ChartFill()
                    EndYear = Mid(data(crow, 1), 1, 4) - 2018 : EndYearBox.Text = EndYear
                    EndMonth = CInt(Mid(data(crow, 1), 6, 2)) : EndMonthBox.Text = EndMonth
                    EndDay = CInt(Mid(data(crow, 1), 9, 2)) : EndDayBox.Text = EndDay
                    EndWeek = 1 + Int(((30.4 * EndMonth) + EndDay) / 7.019)  'EndWeekBox.Text = EndWeek
                    EndHour = CInt(Mid(data(crow, 1), 12, 2)) : EndHourBox.Text = EndHour
                    EndMinute = CInt(Mid(data(crow, 1), 15, 2)) : EndMinuteBox.Text = EndMinute
                Next i
                Call ModuleWriteFiles.UpdateWrite()
                crowsave2 = crow ' store current number of tweets to continue if breaks
                crowSave = crow : cLineSave = cLine
                TextBoxTweets.Text = crow
                TextBoxRead.Text = counter
            End If
            Application.DoEvents()
            wait(5000)
        End While
    End Sub

    Private Sub Form1_Disposed(sender As Object, e As EventArgs) Handles MyBase.Disposed
        cancelFlag = 1
    End Sub
    Public Sub wait(ByVal seconds As Single)
        ' wait for some seconds
        Static start As Single
        start = VB.Timer()
        Do While VB.Timer() < start + seconds
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub
End Class

Module ModuleDecoding
    Public crow = 0
    Public cLine = 0
    Public ReplaceString As String
    Public total
    Public crowSave, cLineSave
    Public Decode_from, Decodede_to As Integer
    Sub decoder()
        Decode_from = crowSave + 1 : Decodede_to = crow
        Parallel.For(Decode_from, Decodede_to + 1, Sub(i)
                                                       'For i = crowSave + 1 To crow
                                                       data(i, 3) = data(i, 3).replace("\xce", "")
                                                       data(i, 3) = data(i, 3).replace("\xcf", "")
                                                       data(i, 3) = data(i, 3).replace("b'", "")
                                                       data(i, 3) = data(i, 3).replace("\x91", "Α")
                                                       data(i, 3) = data(i, 3).replace("\x92", "Β")
                                                       data(i, 3) = data(i, 3).replace("\x93", "Γ")
                                                       data(i, 3) = data(i, 3).replace("\x94", "Δ")
                                                       data(i, 3) = data(i, 3).replace("\x95", "Ε")
                                                       data(i, 3) = data(i, 3).replace("\x96", "Ζ")
                                                       data(i, 3) = data(i, 3).replace("\x97", "Η")
                                                       data(i, 3) = data(i, 3).replace("\x98", "Θ")
                                                       data(i, 3) = data(i, 3).replace("\x99", "Ι")
                                                       data(i, 3) = data(i, 3).replace("\x9a", "Κ")
                                                       data(i, 3) = data(i, 3).replace("\x9b", "Λ")
                                                       data(i, 3) = data(i, 3).replace("\x9c", "Μ")
                                                       data(i, 3) = data(i, 3).replace("\x9d", "Ν")
                                                       data(i, 3) = data(i, 3).replace("\x9e", "Ξ")
                                                       data(i, 3) = data(i, 3).replace("\x9f", "Ο")
                                                       data(i, 3) = data(i, 3).replace("\xa0", "Π")
                                                       data(i, 3) = data(i, 3).replace("\xa1", "Ρ")
                                                       data(i, 3) = data(i, 3).replace("\xa2", "?")
                                                       data(i, 3) = data(i, 3).replace("\xa3", "Σ")
                                                       data(i, 3) = data(i, 3).replace("\xa4", "Τ")
                                                       data(i, 3) = data(i, 3).replace("\xa5", "Υ")
                                                       data(i, 3) = data(i, 3).replace("\xa6", "φ")
                                                       data(i, 3) = data(i, 3).replace("\xa7", "χ")
                                                       data(i, 3) = data(i, 3).replace("\xa8", "Ψ")
                                                       data(i, 3) = data(i, 3).replace("\xa9", "Ω")
                                                       data(i, 3) = data(i, 3).replace("\xac", "ά")
                                                       data(i, 3) = data(i, 3).replace("\xad", "έ")
                                                       data(i, 3) = data(i, 3).replace("\xae", "ή")
                                                       data(i, 3) = data(i, 3).replace("\xaf", "ί")
                                                       data(i, 3) = data(i, 3).replace("\xb1", "α")
                                                       data(i, 3) = data(i, 3).replace("\xb2", "β")
                                                       data(i, 3) = data(i, 3).replace("\xb3", "γ")
                                                       data(i, 3) = data(i, 3).replace("\xb4", "δ")
                                                       data(i, 3) = data(i, 3).replace("\xb5", "ε")
                                                       data(i, 3) = data(i, 3).replace("\xb6", "ζ")
                                                       data(i, 3) = data(i, 3).replace("\xb7", "η")
                                                       data(i, 3) = data(i, 3).replace("\xb8", "θ")
                                                       data(i, 3) = data(i, 3).replace("\xb9", "ι")
                                                       data(i, 3) = data(i, 3).replace("\xba", "κ")
                                                       data(i, 3) = data(i, 3).replace("\xbb", "λ")
                                                       data(i, 3) = data(i, 3).replace("\xbc", "μ")
                                                       data(i, 3) = data(i, 3).replace("\xbd", "ν")
                                                       data(i, 3) = data(i, 3).replace("\xbe", "ξ")
                                                       data(i, 3) = data(i, 3).replace("\xbf", "ο")
                                                       data(i, 3) = data(i, 3).replace("\x80", "π")
                                                       data(i, 3) = data(i, 3).replace("\x81", "ρ")
                                                       data(i, 3) = data(i, 3).replace("\x82", "ς")
                                                       data(i, 3) = data(i, 3).replace("\x83", "σ")
                                                       data(i, 3) = data(i, 3).replace("\x84", "τ")
                                                       data(i, 3) = data(i, 3).replace("\x85", "υ")
                                                       data(i, 3) = data(i, 3).replace("\x86", "φ")
                                                       data(i, 3) = data(i, 3).replace("\x87", "χ")
                                                       data(i, 3) = data(i, 3).replace("\x88", "ψ")
                                                       data(i, 3) = data(i, 3).replace("\x89", "ω")
                                                       data(i, 3) = data(i, 3).replace("\x8d", "ύ")
                                                       data(i, 3) = data(i, 3).replace("\x8e", "ώ")
                                                       data(i, 3) = data(i, 3).replace("\x8c", "ό")
                                                       data(i, 3) = data(i, 3).replace("\x90", "ϊ")
                                                       data(i, 3) = data(i, 3).replace("\xe2", "ύ")
                                                       data(i, 3) = data(i, 3).replace("\n", "")
                                                       'Next i
                                                   End Sub)
    End Sub
End Module

Imports System
Imports System.IO
Imports System.Text
Module ModuleWriteFiles

    Public TimeStamp
    Public ClickTimes
    Sub WriteFile()
        If ClickTimes > 0 Then GoTo endsub
        ' create txt files
        Dim path As String = "D:\Python\data\DataGreek.txt"
        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
            End Using
        End If
        ' open files to write
        Dim objWriter As New System.IO.StreamWriter(path)
        For i = 1 To crow
            ' write values
            objWriter.WriteLine(data(i, 1) & Chr(9) & data(i, 2) & Chr(9) & data(i, 3)) ' store tweets
        Next i
        objWriter.Close()
        'time of last entry
        'TimeMinute = Mid(data(cRow, 1), 15, 2) ' last minute recording
        TimeStamp = data(cRow, 1)
endsub:
    End Sub
    Sub UpdateWrite()
        If ClickTimes > 0 Then GoTo endsub
        For i = cRowSave + 1 To cRow
            ' update txt file
            Dim path As String = "D:\Python\Data\DataGreek.txt"
            Using sw As StreamWriter = File.AppendText(path)
                sw.WriteLine(data(i, 1) & Chr(9) & data(i, 2) & Chr(9) & data(i, 3))
            End Using
        Next i
endsub:
    End Sub

    Sub CalculationTotal()
        Dim OutputMonth, OutputDay, OutputWeek
        Dim CheckDay As Integer = 28 'check day for February
        ' create file
        Dim path As String = "D:\Python\data\StatsDay.txt"
        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
            End Using
        End If
        ' open files to write
        Dim objWriter2 As New System.IO.StreamWriter(path)
        ' calculate sums and write
        objWriter2.WriteLine("yyyy-mm-dd" & Chr(9) & "week" & Chr(9) & "tweets" & Chr(9) & "estimate") ' write headers

        For i = StartYear To EndYear ' for every year
            For k = 1 To 12 ' for every month
                For m = 1 To 31 ' for every day
                    If k < 10 Then OutputMonth = "0" & k Else OutputMonth = k 'If dataDay(i, k, m) Is Nothing Then dataDay(i, k, m) = 0
                    If m < 10 Then OutputDay = "0" & m Else OutputDay = m
                    If WeekCalculation(i, k, m) < 10 Then OutputWeek = "0" & WeekCalculation(i, k, m) Else OutputWeek = WeekCalculation(i, k, m)
                    ' check start
                    If i = 0 And k < 12 Or (i = 0 And k = 12 And m < 13) Then GoTo Notwrite ' start writing
                    If i = EndYear And k > EndMonth Or (i = EndYear And k = EndMonth And m > EndDay) Then GoTo Notwrite ' stop writing
                    If (k = 4 Or k = 6 Or k = 9 Or k = 11) And m = 31 Then GoTo Notwrite ' stop writing for months with 30 days
                    If (i + 2018) Mod 4 = 0 Then CheckDay = 29 Else CheckDay = 28 ' check February for 29 δαυσ
                    If k = 2 And m > CheckDay Then GoTo Notwrite ' check February of 28 days
                    objWriter2.WriteLine(i + 2018 & "-" & OutputMonth & "-" & OutputDay & Chr(9) & OutputWeek & Chr(9) & dataDay(i, k, m) & Chr(9) & dataDay2(i, k, m)) ' store tweets stats
Notwrite:
                Next m
            Next k
        Next i
        objWriter2.Close()
    End Sub

    Sub WriteDailylyStats2()
        If ClickTimes > 0 Then Exit Sub
        ' create txt files
        Dim path As String = "D:\Python\data\StatsDay2.txt"
        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
            End Using
        End If
        ' open files to write
        Dim objWriter3 As New System.IO.StreamWriter(path)
        objWriter3.WriteLine("count" & Chr(9) & "yyyy-mm-dd" & Chr(9) & "week" & Chr(9) & "tweets" & Chr(9) & "est1" & Chr(9) & "year" & Chr(9) &
                             "tweets2" & Chr(9) & "est2") ' store tweets
        For i = 2 To cRow2
            ' write values
            objWriter3.WriteLine(StatsDay1(i, 0) & Chr(9) & StatsDay1(i, 1) & Chr(9) & StatsDay1(i, 2) & Chr(9) &
                                 StatsDay1(i, 3) & Chr(9) & StatsDay1(i, 4) & Chr(9) & StatsDay1(i, 8) & Chr(9) &
                                 StatsDay1(i, 5) & Chr(9) & StatsDay1(i, 6)) ' store results
        Next i
        objWriter3.Close()

    End Sub

    Sub WriteWeeklyStats()
        If ClickTimes > 0 Then Exit Sub
        ' create txt files 
        Dim path As String = "D:\Python\data\StatsWeek.txt"
        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
            End Using
        End If
        ' open files to write
        Dim objWriter4 As New System.IO.StreamWriter(path)
        objWriter4.WriteLine("count" & Chr(9) & "year" & Chr(9) & "week" & Chr(9) & "tweets" & Chr(9) & "est1" & Chr(9) & "label" &
                             Chr(9) & "est2" & Chr(9) & "tweets2" & Chr(9) & "ARIMA" & Chr(9) & " R(1)" & Chr(9) & "R(2)") ' store tweets
        For i = 1 To WeekCount
            ' write values
            objWriter4.WriteLine(StatsWeek(i, 0) & Chr(9) & StatsWeek(i, 1) & Chr(9) & StatsWeek(i, 2) & Chr(9) & StatsWeek(i, 3) & Chr(9) &
                StatsWeek(i, 4) & Chr(9) & StatsWeek(i, 11) & Chr(9) & StatsWeek(i, 6) & Chr(9) & StatsWeek(i, 5) & Chr(9) & Prediction(i) &
                Chr(9) & Pearson1(i) & Chr(9) & Pearson2(i)) ' store tweets
        Next i
        objWriter4.Close()

    End Sub
End Module
