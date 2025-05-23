' ===========================================
' 📄 Services/ScheduleStatisticsDisplayService.vb
' -------------------------------------------
' განრიგის სტატისტიკური ინფორმაციის ჩვენების სერვისი
' პასუხისმგებელია GBSumInf გრუპბოქსის ლეიბლების განახლებაზე
' ===========================================
Imports System.Drawing
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Services
Imports System.Globalization
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის სტატისტიკური ინფორმაციის ჩვენების სერვისი
    ''' პასუხისმგებელია GBSumInf პანელში ყველა ლეიბლის განახლებაზე
    ''' </summary>
    Public Class ScheduleStatisticsDisplayService

        ' მონაცემთა სერვისი
        Private ReadOnly dataService As IDataService

        ' UI ლეიბლები სტატისტიკისთვის
        Private ReadOnly statisticsLabels As Dictionary(Of String, Label)

        ''' <summary>
        ''' სტატისტიკის შედეგების მოდელი
        ''' </summary>
        Public Class SessionStatistics
            ''' <summary>სეანსების ზოგადი სტატისტიკა</summary>
            Public Property TotalSessions As Integer = 0
            Public Property TotalDurationMinutes As Integer = 0

            ''' <summary>შესრულებული სეანსები</summary>
            Public Property CompletedSessions As Integer = 0
            Public Property CompletedDurationMinutes As Integer = 0

            ''' <summary>აღდგენილი სეანსები</summary>
            Public Property RestoredSessions As Integer = 0
            Public Property RestoredDurationMinutes As Integer = 0

            ''' <summary>გაცდენა არასაპატიო</summary>
            Public Property MissedUnexcusedSessions As Integer = 0
            Public Property MissedUnexcusedDurationMinutes As Integer = 0

            ''' <summary>გაცდენა საპატიო</summary>
            Public Property MissedExcusedSessions As Integer = 0
            Public Property MissedExcusedDurationMinutes As Integer = 0

            ''' <summary>პროგრამით გატარება</summary>
            Public Property AutoProcessedSessions As Integer = 0
            Public Property AutoProcessedDurationMinutes As Integer = 0

            ''' <summary>გაუქმებული სეანსები</summary>
            Public Property CancelledSessions As Integer = 0
            Public Property CancelledDurationMinutes As Integer = 0

            ''' <summary>დაგეგმილი (მომავალი)</summary>
            Public Property PlannedFutureSessions As Integer = 0
            Public Property PlannedFutureDurationMinutes As Integer = 0

            ''' <summary>დაგეგმილი (ვადაგასული)</summary>
            Public Property PlannedOverdueSessions As Integer = 0
            Public Property PlannedOverdueDurationMinutes As Integer = 0

            ''' <summary>უნიკალური ბენეფიციარები</summary>
            Public Property UniqueBeneficiaries As Integer = 0

            ''' <summary>უნიკალური თერაპევტები</summary>
            Public Property UniqueTherapists As Integer = 0

            ''' <summary>
            ''' საათების გამოთვლა (წუთები/60)
            ''' </summary>
            Public Function GetHours(minutes As Integer) As Double
                Return Math.Round(minutes / 60.0, 2)
            End Function

            ''' <summary>
            ''' 30-წუთიანი სეანსების რაოდენობა
            ''' </summary>
            Public Function GetSessions30(minutes As Integer) As Double
                Return Math.Round(minutes / 30.0, 2)
            End Function

            ''' <summary>
            ''' 60-წუთიანი სეანსების რაოდენობა
            ''' </summary>
            Public Function GetSessions60(minutes As Integer) As Double
                Return Math.Round(minutes / 60.0, 2)
            End Function
        End Class

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="service">მონაცემთა სერვისი</param>
        ''' <param name="groupBox">GBSumInf გრუპბოქსი</param>
        Public Sub New(service As IDataService, groupBox As GroupBox)
            dataService = service
            statisticsLabels = New Dictionary(Of String, Label)()

            ' ლეიბლების ძებნა და რეგისტრაცია
            InitializeLabels(groupBox)
        End Sub

        ''' <summary>
        ''' ლეიბლების ძებნა და რეგისტრაცია GBSumInf გრუპბოქსში
        ''' </summary>
        ''' <param name="groupBox">GBSumInf გრუპბოქსი</param>
        Private Sub InitializeLabels(groupBox As GroupBox)
            Try
                Debug.WriteLine("ScheduleStatisticsDisplayService: ლეიბლების ძებნა დაიწყო")

                If groupBox Is Nothing Then
                    Debug.WriteLine("ScheduleStatisticsDisplayService: GBSumInf გრუპბოქსი არ არის მოწოდებული")
                    Return
                End If

                ' ყველა საჭირო ლეიბლის სახელების სია
                Dim labelNames As String() = {
                    "lsr", "lsm", "ls30", "ls60",
                    "lshess", "lshesm", "lshes30", "lshes60",
                    "las", "lam", "la30", "la60",
                    "lgas", "lgam", "lga30", "lga60",
                    "lgss", "lgsm", "lgs30", "lgs60",
                    "lpgs", "lpgm", "lpg30", "lpg60",
                    "lgs", "lgm", "lg30", "lg60",
                    "lds", "ldm", "ld30", "ld60",
                    "ldvs", "ldvm", "ldv30", "ldv60",
                    "lbenes", "lpers"
                }

                ' ლეიბლების ძებნა გრუპბოქსში
                For Each labelName In labelNames
                    Dim foundLabel = FindLabelRecursive(groupBox, labelName)
                    If foundLabel IsNot Nothing Then
                        statisticsLabels(labelName) = foundLabel
                        Debug.WriteLine($"ScheduleStatisticsDisplayService: ნაპოვნია ლეიბლი '{labelName}'")
                    Else
                        Debug.WriteLine($"ScheduleStatisticsDisplayService: ვერ მოიძებნა ლეიბლი '{labelName}'")
                    End If
                Next

                Debug.WriteLine($"ScheduleStatisticsDisplayService: მთლიანად ნაპოვნია {statisticsLabels.Count} ლეიბლი {labelNames.Length}-დან")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: InitializeLabels შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ლეიბლის რეკურსიული ძებნა კონტროლის იერარქიაში
        ''' </summary>
        ''' <param name="parent">მშობელი კონტროლი</param>
        ''' <param name="labelName">ლეიბლის სახელი</param>
        ''' <returns>ნაპოვნი ლეიბლი ან Nothing</returns>
        Private Function FindLabelRecursive(parent As Control, labelName As String) As Label
            Try
                ' ჯერ პირდაპირ შვილებში ვეძებთ
                For Each ctrl As Control In parent.Controls
                    If TypeOf ctrl Is Label AndAlso
                       String.Equals(ctrl.Name, labelName, StringComparison.OrdinalIgnoreCase) Then
                        Return DirectCast(ctrl, Label)
                    End If
                Next

                ' მერე რეკურსიულად შვილების შვილებში
                For Each ctrl As Control In parent.Controls
                    Dim found = FindLabelRecursive(ctrl, labelName)
                    If found IsNot Nothing Then
                        Return found
                    End If
                Next

                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' სტატისტიკის გამოთვლა გაფილტრული მონაცემებისთვის
        ''' </summary>
        ''' <param name="filteredData">გაფილტრული მონაცემები</param>
        ''' <returns>სტატისტიკის შედეგი</returns>
        Public Function CalculateStatistics(filteredData As List(Of IList(Of Object))) As SessionStatistics
            Try
                Debug.WriteLine($"ScheduleStatisticsDisplayService: სტატისტიკის გამოთვლა {filteredData.Count} ჩანაწერისთვის")

                Dim stats As New SessionStatistics()
                Dim uniqueBeneficiaries As New HashSet(Of String)()
                Dim uniqueTherapists As New HashSet(Of String)()
                Dim currentTime As DateTime = DateTime.Now

                ' ყოველი ჩანაწერის დამუშავება
                For Each row In filteredData
                    Try
                        ' ძირითადი ვალიდაცია
                        If row.Count < 13 Then Continue For

                        ' სტატუსის მიღება
                        Dim status As String = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")

                        ' ხანგრძლივობის მიღება
                        Dim duration As Integer = 60 ' ნაგულისხმევი
                        If row.Count > 6 AndAlso IsNumeric(row(6)) Then
                            duration = CInt(row(6))
                        End If

                        ' სესიის თარიღის მიღება
                        Dim sessionDate As DateTime = DateTime.Now
                        If row.Count > 5 Then
                            DateTime.TryParse(row(5).ToString(), sessionDate)
                        End If

                        ' სულ სეანსები
                        stats.TotalSessions += 1
                        stats.TotalDurationMinutes += duration

                        ' სტატუსის მიხედვით კატეგორიზაცია
                        Select Case status
                            Case "შესრულებული"
                                stats.CompletedSessions += 1
                                stats.CompletedDurationMinutes += duration

                            Case "აღდგენა"
                                stats.RestoredSessions += 1
                                stats.RestoredDurationMinutes += duration

                            Case "გაცდენა არასაპატიო"
                                stats.MissedUnexcusedSessions += 1
                                stats.MissedUnexcusedDurationMinutes += duration

                            Case "გაცდენა საპატიო"
                                stats.MissedExcusedSessions += 1
                                stats.MissedExcusedDurationMinutes += duration

                            Case "პროგრამით გატარება"
                                stats.AutoProcessedSessions += 1
                                stats.AutoProcessedDurationMinutes += duration

                            Case "გაუქმება", "გაუქმებული"
                                stats.CancelledSessions += 1
                                stats.CancelledDurationMinutes += duration

                            Case "დაგეგმილი"
                                ' დაგეგმილი სეანსებისთვის ვამოწმებთ ვადაგადაცილებას
                                If sessionDate <= currentTime Then
                                    ' ვადაგასული
                                    stats.PlannedOverdueSessions += 1
                                    stats.PlannedOverdueDurationMinutes += duration
                                Else
                                    ' მომავალი
                                    stats.PlannedFutureSessions += 1
                                    stats.PlannedFutureDurationMinutes += duration
                                End If
                        End Select

                        ' უნიკალური ბენეფიციარების დათვლა
                        If row.Count > 4 Then
                            Dim beneficiaryName = $"{row(3).ToString().Trim()} {row(4).ToString().Trim()}"
                            If Not String.IsNullOrWhiteSpace(beneficiaryName.Trim()) Then
                                uniqueBeneficiaries.Add(beneficiaryName)
                            End If
                        End If

                        ' უნიკალური თერაპევტების დათვლა
                        If row.Count > 8 AndAlso Not String.IsNullOrWhiteSpace(row(8).ToString()) Then
                            uniqueTherapists.Add(row(8).ToString().Trim())
                        End If

                    Catch ex As Exception
                        Debug.WriteLine($"ScheduleStatisticsDisplayService: მწკრივის დამუშავების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next

                ' უნიკალური მონაწილეების რაოდენობის დაყენება
                stats.UniqueBeneficiaries = uniqueBeneficiaries.Count
                stats.UniqueTherapists = uniqueTherapists.Count

                Debug.WriteLine($"ScheduleStatisticsDisplayService: სტატისტიკა გამოთვლილია - სულ: {stats.TotalSessions}, შესრულებული: {stats.CompletedSessions}")
                Return stats

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: CalculateStatistics შეცდომა: {ex.Message}")
                Return New SessionStatistics()
            End Try
        End Function

        ''' <summary>
        ''' სტატისტიკის ჩვენება ლეიბლებში
        ''' </summary>
        ''' <param name="stats">სტატისტიკის შედეგი</param>
        Public Sub DisplayStatistics(stats As SessionStatistics)
            Try
                Debug.WriteLine("ScheduleStatisticsDisplayService: სტატისტიკის ჩვენება ლეიბლებში")

                If stats Is Nothing Then
                    Debug.WriteLine("ScheduleStatisticsDisplayService: სტატისტიკა არის Nothing")
                    Return
                End If

                ' სულ სეანსები
                UpdateLabel("lsr", stats.TotalSessions.ToString())
                UpdateLabel("lsm", $"{stats.TotalDurationMinutes} წთ")
                UpdateLabel("ls30", stats.GetSessions30(stats.TotalDurationMinutes).ToString("F1"))
                UpdateLabel("ls60", stats.GetSessions60(stats.TotalDurationMinutes).ToString("F1"))

                ' შესრულებული სეანსები
                UpdateLabel("lshess", stats.CompletedSessions.ToString())
                UpdateLabel("lshesm", $"{stats.CompletedDurationMinutes} წთ")
                UpdateLabel("lshes30", stats.GetSessions30(stats.CompletedDurationMinutes).ToString("F1"))
                UpdateLabel("lshes60", stats.GetSessions60(stats.CompletedDurationMinutes).ToString("F1"))

                ' აღდგენილი სეანსები
                UpdateLabel("las", stats.RestoredSessions.ToString())
                UpdateLabel("lam", $"{stats.RestoredDurationMinutes} წთ")
                UpdateLabel("la30", stats.GetSessions30(stats.RestoredDurationMinutes).ToString("F1"))
                UpdateLabel("la60", stats.GetSessions60(stats.RestoredDurationMinutes).ToString("F1"))

                ' გაცდენა არასაპატიო
                UpdateLabel("lgas", stats.MissedUnexcusedSessions.ToString())
                UpdateLabel("lgam", $"{stats.MissedUnexcusedDurationMinutes} წთ")
                UpdateLabel("lga30", stats.GetSessions30(stats.MissedUnexcusedDurationMinutes).ToString("F1"))
                UpdateLabel("lga60", stats.GetSessions60(stats.MissedUnexcusedDurationMinutes).ToString("F1"))

                ' გაცდენა საპატიო
                UpdateLabel("lgss", stats.MissedExcusedSessions.ToString())
                UpdateLabel("lgsm", $"{stats.MissedExcusedDurationMinutes} წთ")
                UpdateLabel("lgs30", stats.GetSessions30(stats.MissedExcusedDurationMinutes).ToString("F1"))
                UpdateLabel("lgs60", stats.GetSessions60(stats.MissedExcusedDurationMinutes).ToString("F1"))

                ' პროგრამით გატარება
                UpdateLabel("lpgs", stats.AutoProcessedSessions.ToString())
                UpdateLabel("lpgm", $"{stats.AutoProcessedDurationMinutes} წთ")
                UpdateLabel("lpg30", stats.GetSessions30(stats.AutoProcessedDurationMinutes).ToString("F1"))
                UpdateLabel("lpg60", stats.GetSessions60(stats.AutoProcessedDurationMinutes).ToString("F1"))

                ' გაუქმებული სეანსები
                UpdateLabel("lgs", stats.CancelledSessions.ToString())
                UpdateLabel("lgm", $"{stats.CancelledDurationMinutes} წთ")
                UpdateLabel("lg30", stats.GetSessions30(stats.CancelledDurationMinutes).ToString("F1"))
                UpdateLabel("lg60", stats.GetSessions60(stats.CancelledDurationMinutes).ToString("F1"))

                ' დაგეგმილი მომავალი
                UpdateLabel("lds", stats.PlannedFutureSessions.ToString())
                UpdateLabel("ldm", $"{stats.PlannedFutureDurationMinutes} წთ")
                UpdateLabel("ld30", stats.GetSessions30(stats.PlannedFutureDurationMinutes).ToString("F1"))
                UpdateLabel("ld60", stats.GetSessions60(stats.PlannedFutureDurationMinutes).ToString("F1"))

                ' დაგეგმილი ვადაგასული
                UpdateLabel("ldvs", stats.PlannedOverdueSessions.ToString())
                UpdateLabel("ldvm", $"{stats.PlannedOverdueDurationMinutes} წთ")
                UpdateLabel("ldv30", stats.GetSessions30(stats.PlannedOverdueDurationMinutes).ToString("F1"))
                UpdateLabel("ldv60", stats.GetSessions60(stats.PlannedOverdueDurationMinutes).ToString("F1"))

                ' უნიკალური მონაწილეები
                UpdateLabel("lbenes", stats.UniqueBeneficiaries.ToString())
                UpdateLabel("lpers", stats.UniqueTherapists.ToString())

                Debug.WriteLine("ScheduleStatisticsDisplayService: ყველა ლეიბლი განახლდა")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: DisplayStatistics შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' კონკრეტული ლეიბლის განახლება
        ''' </summary>
        ''' <param name="labelName">ლეიბლის სახელი</param>
        ''' <param name="value">ახალი მნიშვნელობა</param>
        Private Sub UpdateLabel(labelName As String, value As String)
            Try
                If statisticsLabels.ContainsKey(labelName) AndAlso statisticsLabels(labelName) IsNot Nothing Then
                    statisticsLabels(labelName).Text = value
                    Debug.WriteLine($"ScheduleStatisticsDisplayService: ლეიბლი '{labelName}' განახლდა: '{value}'")
                Else
                    Debug.WriteLine($"ScheduleStatisticsDisplayService: ლეიბლი '{labelName}' ვერ მოიძებნა განსახლებისთვის")
                End If
            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: UpdateLabel შეცდომა '{labelName}': {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ყველა ლეიბლის გასუფთავება
        ''' </summary>
        Public Sub ClearAllLabels()
            Try
                Debug.WriteLine("ScheduleStatisticsDisplayService: ყველა ლეიბლის გასუფთავება")

                For Each kvp In statisticsLabels
                    If kvp.Value IsNot Nothing Then
                        kvp.Value.Text = "0"
                    End If
                Next

                Debug.WriteLine("ScheduleStatisticsDisplayService: ყველა ლეიბლი გასუფთავებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: ClearAllLabels შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სტატისტიკის ასინქრონული განახლება
        ''' </summary>
        ''' <param name="filteredData">გაფილტრული მონაცემები</param>
        Public Sub UpdateStatisticsAsync(filteredData As List(Of IList(Of Object)))
            Try
                ' ფონზე სტატისტიკის გამოთვლა
                System.Threading.Tasks.Task.Run(Sub()
                                                    Try
                                                        Dim stats = CalculateStatistics(filteredData)

                                                        ' UI Thread-ზე ლეიბლების განახლება
                                                        If statisticsLabels.Count > 0 AndAlso statisticsLabels.Values.First() IsNot Nothing Then
                                                            statisticsLabels.Values.First().Invoke(Sub()
                                                                                                       DisplayStatistics(stats)
                                                                                                   End Sub)
                                                        End If

                                                    Catch ex As Exception
                                                        Debug.WriteLine($"ScheduleStatisticsDisplayService: UpdateStatisticsAsync ფონური შეცდომა: {ex.Message}")
                                                    End Try
                                                End Sub)

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsDisplayService: UpdateStatisticsAsync შეცდომა: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace