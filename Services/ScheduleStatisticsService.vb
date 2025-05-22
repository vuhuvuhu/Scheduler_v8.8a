' ===========================================
' 📄 Services/ScheduleStatisticsService.vb
' -------------------------------------------
' განრიგის სტატისტიკისა და შემაჯამებელი ინფორმაციის სერვისი
' UC_Schedule-დან გატანილი სტატისტიკის ლოგიკა
' ===========================================
Imports System.Globalization
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის სტატისტიკისა და შემაჯამებელი ინფორმაციის სერვისი
    ''' პასუხისმგებელია რიცხვობრივი მონაცემების დამუშავება და ანალიზი
    ''' </summary>
    Public Class ScheduleStatisticsService

        ' მონაცემთა სერვისი
        Private ReadOnly dataService As IDataService

        ''' <summary>
        ''' სტატისტიკის შედეგების კლასი
        ''' </summary>
        Public Class StatisticsResult
            ''' <summary>სულ სესიების რაოდენობა</summary>
            Public Property TotalSessions As Integer

            ''' <summary>შესრულებული სესიების რაოდენობა</summary>
            Public Property CompletedSessions As Integer

            ''' <summary>დაგეგმილი სესიების რაოდენობა</summary>
            Public Property PlannedSessions As Integer

            ''' <summary>გაუქმებული სესიების რაოდენობა</summary>
            Public Property CancelledSessions As Integer

            ''' <summary>ვადაგადაცილებული სესიების რაოდენობა</summary>
            Public Property OverdueSessions As Integer

            ''' <summary>სულ ფასი (შესრულებული სესიებისთვის)</summary>
            Public Property TotalRevenue As Decimal

            ''' <summary>საშუალო სესიის ფასი</summary>
            Public Property AveragePrice As Decimal

            ''' <summary>უნიკალური ბენეფიციარების რაოდენობა</summary>
            Public Property UniqueBeneficiaries As Integer

            ''' <summary>უნიკალური თერაპევტების რაოდენობა</summary>
            Public Property UniqueTherapists As Integer

            ''' <summary>სტატუსების მიხედვით დანაწილება</summary>
            Public Property StatusBreakdown As Dictionary(Of String, Integer)

            ''' <summary>თერაპიის ტიპების მიხედვით დანაწილება</summary>
            Public Property TherapyTypeBreakdown As Dictionary(Of String, Integer)

            ''' <summary>ყველაზე აქტიური თერაპევტი</summary>
            Public Property MostActiveTherapist As String

            ''' <summary>ყველაზე აქტიური ბენეფიციარი</summary>
            Public Property MostActiveBeneficiary As String

            ''' <summary>
            ''' კონსტრუქტორი - საწყისი მნიშვნელობებით
            ''' </summary>
            Public Sub New()
                StatusBreakdown = New Dictionary(Of String, Integer)()
                TherapyTypeBreakdown = New Dictionary(Of String, Integer)()
                MostActiveTherapist = ""
                MostActiveBeneficiary = ""
            End Sub

            ''' <summary>
            ''' შესრულების პროცენტი
            ''' </summary>
            Public ReadOnly Property CompletionRate As Double
                Get
                    If TotalSessions > 0 Then
                        Return Math.Round((CompletedSessions / TotalSessions) * 100, 2)
                    End If
                    Return 0
                End Get
            End Property

            ''' <summary>
            ''' გაუქმების პროცენტი
            ''' </summary>
            Public ReadOnly Property CancellationRate As Double
                Get
                    If TotalSessions > 0 Then
                        Return Math.Round((CancelledSessions / TotalSessions) * 100, 2)
                    End If
                    Return 0
                End Get
            End Property
        End Class

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="service">მონაცემთა სერვისი</param>
        Public Sub New(service As IDataService)
            dataService = service
        End Sub

        ''' <summary>
        ''' სტატისტიკის გამოთვლა მოცემული ფილტრის კრიტერიუმებისთვის
        ''' </summary>
        ''' <param name="criteria">ფილტრის კრიტერიუმები</param>
        ''' <returns>სტატისტიკის შედეგი</returns>
        Public Function CalculateStatistics(criteria As ScheduleDataProcessor.FilterCriteria) As StatisticsResult
            Try
                Debug.WriteLine("ScheduleStatisticsService: სტატისტიკის გამოთვლა დაიწყო")

                Dim result As New StatisticsResult()

                ' ყველა მონაცემის მიღება
                Dim allRows As IList(Of IList(Of Object)) = dataService.GetData("DB-Schedule!A2:O")

                If allRows Is Nothing OrElse allRows.Count = 0 Then
                    Debug.WriteLine("ScheduleStatisticsService: მონაცემები არ მოიძებნა")
                    Return result
                End If

                ' ფილტრირებული მონაცემების მიღება
                Dim filteredRows = FilterDataByCriteria(allRows, criteria)
                Debug.WriteLine($"ScheduleStatisticsService: გაფილტრულია {filteredRows.Count} ჩანაწერი {allRows.Count}-დან")

                ' ძირითადი სტატისტიკის გამოთვლა
                CalculateBasicStatistics(filteredRows, result)

                ' დეტალური ანალიზის გამოთვლა
                CalculateDetailedAnalysis(filteredRows, result)

                Debug.WriteLine($"ScheduleStatisticsService: სტატისტიკა გამოთვლილია - სულ: {result.TotalSessions}, შესრულებული: {result.CompletedSessions}")
                Return result

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: CalculateStatistics შეცდომა: {ex.Message}")
                Return New StatisticsResult()
            End Try
        End Function

        ''' <summary>
        ''' მონაცემების ფილტრაცია კრიტერიუმების მიხედვით
        ''' </summary>
        ''' <param name="allRows">ყველა ჩანაწერი</param>
        ''' <param name="criteria">ფილტრის კრიტერიუმები</param>
        ''' <returns>ფილტრირებული ჩანაწერები</returns>
        Private Function FilterDataByCriteria(allRows As IList(Of IList(Of Object)), criteria As ScheduleDataProcessor.FilterCriteria) As List(Of IList(Of Object))
            Dim filtered As New List(Of IList(Of Object))

            Try
                For Each row As IList(Of Object) In allRows
                    If IsRowMatchingCriteria(row, criteria) Then
                        filtered.Add(row)
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: FilterDataByCriteria შეცდომა: {ex.Message}")
            End Try

            Return filtered
        End Function

        ''' <summary>
        ''' ჩანაწერის შესაბამისობის შემოწმება კრიტერიუმებთან
        ''' </summary>
        Private Function IsRowMatchingCriteria(row As IList(Of Object), criteria As ScheduleDataProcessor.FilterCriteria) As Boolean
            Try
                ' ძირითადი ვალიდაცია
                If row.Count < 7 Then Return False
                If String.IsNullOrWhiteSpace(row(0).ToString()) Then Return False
                If String.IsNullOrWhiteSpace(row(5).ToString()) Then Return False

                ' თარიღის ფილტრი
                Dim sessionDate As DateTime
                Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                If Not DateTime.TryParseExact(row(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, sessionDate) Then
                    If Not DateTime.TryParse(row(5).ToString().Trim(), sessionDate) Then
                        Return False
                    End If
                End If

                If sessionDate.Date < criteria.DateFrom.Date OrElse sessionDate.Date > criteria.DateTo.Date Then
                    Return False
                End If

                ' სტატუსის ფილტრი
                If criteria.SelectedStatuses.Count > 0 Then
                    Dim rowStatus As String = If(row.Count > 12, row(12).ToString().Trim(), "")
                    If Not criteria.SelectedStatuses.Any(Function(s) String.Equals(s, rowStatus, StringComparison.OrdinalIgnoreCase)) Then
                        Return False
                    End If
                End If

                Return True

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: IsRowMatchingCriteria შეცდომა: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' ძირითადი სტატისტიკის გამოთვლა
        ''' </summary>
        Private Sub CalculateBasicStatistics(filteredRows As List(Of IList(Of Object)), result As StatisticsResult)
            Try
                result.TotalSessions = filteredRows.Count

                Dim totalRevenue As Decimal = 0
                Dim completedCount As Integer = 0
                Dim plannedCount As Integer = 0
                Dim cancelledCount As Integer = 0
                Dim overdueCount As Integer = 0
                Dim currentTime As DateTime = DateTime.Now

                For Each row In filteredRows
                    Try
                        ' სტატუსის მიღება
                        Dim status As String = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")

                        ' სტატუსების დათვლა
                        Select Case status
                            Case "შესრულებული"
                                completedCount += 1
                                ' შესრულებული სესიების ფასის დამატება
                                If row.Count > 11 Then
                                    Dim price As Decimal
                                    If Decimal.TryParse(row(11).ToString().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, price) Then
                                        totalRevenue += price
                                    End If
                                End If
                            Case "დაგეგმილი"
                                plannedCount += 1
                                ' ვადაგადაცილების შემოწმება
                                Dim sessionDate As DateTime
                                If DateTime.TryParse(row(5).ToString(), sessionDate) AndAlso sessionDate < currentTime Then
                                    overdueCount += 1
                                End If
                            Case "გაუქმება", "გაუქმებული"
                                cancelledCount += 1
                        End Select

                    Catch ex As Exception
                        Debug.WriteLine($"ScheduleStatisticsService: ძირითადი სტატისტიკის მწკრივის დამუშავების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next

                ' შედეგების შენახვა
                result.CompletedSessions = completedCount
                result.PlannedSessions = plannedCount
                result.CancelledSessions = cancelledCount
                result.OverdueSessions = overdueCount
                result.TotalRevenue = totalRevenue
                result.AveragePrice = If(completedCount > 0, Math.Round(totalRevenue / completedCount, 2), 0)

                Debug.WriteLine($"ScheduleStatisticsService: ძირითადი სტატისტიკა - სულ: {result.TotalSessions}, შემოსავალი: {result.TotalRevenue}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: CalculateBasicStatistics შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' დეტალური ანალიზის გამოთვლა
        ''' </summary>
        Private Sub CalculateDetailedAnalysis(filteredRows As List(Of IList(Of Object)), result As StatisticsResult)
            Try
                Dim uniqueBeneficiaries As New HashSet(Of String)()
                Dim uniqueTherapists As New HashSet(Of String)()
                Dim therapistCounts As New Dictionary(Of String, Integer)()
                Dim beneficiaryCounts As New Dictionary(Of String, Integer)()

                For Each row In filteredRows
                    Try
                        ' ბენეფიციარის სრული სახელი
                        If row.Count > 4 Then
                            Dim beneficiaryName = $"{row(3).ToString().Trim()} {row(4).ToString().Trim()}"
                            uniqueBeneficiaries.Add(beneficiaryName)

                            ' ბენეფიციარის აქტივობის დათვლა
                            If beneficiaryCounts.ContainsKey(beneficiaryName) Then
                                beneficiaryCounts(beneficiaryName) += 1
                            Else
                                beneficiaryCounts(beneficiaryName) = 1
                            End If
                        End If

                        ' თერაპევტი
                        If row.Count > 8 AndAlso Not String.IsNullOrWhiteSpace(row(8).ToString()) Then
                            Dim therapistName = row(8).ToString().Trim()
                            uniqueTherapists.Add(therapistName)

                            ' თერაპევტის აქტივობის დათვლა
                            If therapistCounts.ContainsKey(therapistName) Then
                                therapistCounts(therapistName) += 1
                            Else
                                therapistCounts(therapistName) = 1
                            End If
                        End If

                        ' სტატუსების დანაწილება
                        If row.Count > 12 Then
                            Dim status = row(12).ToString().Trim()
                            If Not String.IsNullOrWhiteSpace(status) Then
                                If result.StatusBreakdown.ContainsKey(status) Then
                                    result.StatusBreakdown(status) += 1
                                Else
                                    result.StatusBreakdown(status) = 1
                                End If
                            End If
                        End If

                        ' თერაპიის ტიპების დანაწილება
                        If row.Count > 9 Then
                            Dim therapyType = row(9).ToString().Trim()
                            If Not String.IsNullOrWhiteSpace(therapyType) Then
                                If result.TherapyTypeBreakdown.ContainsKey(therapyType) Then
                                    result.TherapyTypeBreakdown(therapyType) += 1
                                Else
                                    result.TherapyTypeBreakdown(therapyType) = 1
                                End If
                            End If
                        End If

                    Catch ex As Exception
                        Debug.WriteLine($"ScheduleStatisticsService: დეტალური ანალიზის მწკრივის დამუშავების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next

                ' შედეგების შენახვა
                result.UniqueBeneficiaries = uniqueBeneficiaries.Count
                result.UniqueTherapists = uniqueTherapists.Count

                ' ყველაზე აქტიური თერაპევტი
                If therapistCounts.Count > 0 Then
                    Dim mostActiveTherapist = therapistCounts.OrderByDescending(Function(kvp) kvp.Value).First()
                    result.MostActiveTherapist = $"{mostActiveTherapist.Key} ({mostActiveTherapist.Value} სესია)"
                End If

                ' ყველაზე აქტიური ბენეფიციარი
                If beneficiaryCounts.Count > 0 Then
                    Dim mostActiveBeneficiary = beneficiaryCounts.OrderByDescending(Function(kvp) kvp.Value).First()
                    result.MostActiveBeneficiary = $"{mostActiveBeneficiary.Key} ({mostActiveBeneficiary.Value} სესია)"
                End If

                Debug.WriteLine($"ScheduleStatisticsService: დეტალური ანალიზი - ბენეფიციარები: {result.UniqueBeneficiaries}, თერაპევტები: {result.UniqueTherapists}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: CalculateDetailedAnalysis შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სტატისტიკის ტექსტური განმარტება
        ''' </summary>
        ''' <param name="stats">სტატისტიკის შედეგი</param>
        ''' <returns>ტექსტური განმარტება</returns>
        Public Shared Function FormatStatistics(stats As StatisticsResult) As String
            Try
                Dim report As New System.Text.StringBuilder()

                report.AppendLine("📊 განრიგის სტატისტიკა")
                report.AppendLine("".PadRight(50, "="c))
                report.AppendLine()

                ' ძირითადი მონაცემები
                report.AppendLine("🔢 ძირითადი მონაცემები:")
                report.AppendLine($"   • სულ სესიები: {stats.TotalSessions}")
                report.AppendLine($"   • შესრულებული: {stats.CompletedSessions} ({stats.CompletionRate}%)")
                report.AppendLine($"   • დაგეგმილი: {stats.PlannedSessions}")
                report.AppendLine($"   • გაუქმებული: {stats.CancelledSessions} ({stats.CancellationRate}%)")
                report.AppendLine($"   • ვადაგადაცილებული: {stats.OverdueSessions}")
                report.AppendLine()

                ' ფინანსური მონაცემები
                report.AppendLine("💰 ფინანსური მონაცემები:")
                report.AppendLine($"   • სულ შემოსავალი: {stats.TotalRevenue:F2} ლარი")
                report.AppendLine($"   • საშუალო ფასი: {stats.AveragePrice:F2} ლარი")
                report.AppendLine()

                ' მონაწილეები
                report.AppendLine("👥 მონაწილეები:")
                report.AppendLine($"   • უნიკალური ბენეფიციარები: {stats.UniqueBeneficiaries}")
                report.AppendLine($"   • უნიკალური თერაპევტები: {stats.UniqueTherapists}")
                If Not String.IsNullOrEmpty(stats.MostActiveTherapist) Then
                    report.AppendLine($"   • ყველაზე აქტიური თერაპევტი: {stats.MostActiveTherapist}")
                End If
                If Not String.IsNullOrEmpty(stats.MostActiveBeneficiary) Then
                    report.AppendLine($"   • ყველაზე აქტიური ბენეფიციარი: {stats.MostActiveBeneficiary}")
                End If
                report.AppendLine()

                ' სტატუსების დანაწილება
                If stats.StatusBreakdown.Count > 0 Then
                    report.AppendLine("📈 სტატუსების მიხედვით:")
                    For Each kvp In stats.StatusBreakdown.OrderByDescending(Function(x) x.Value)
                        Dim percentage = If(stats.TotalSessions > 0, Math.Round((kvp.Value / stats.TotalSessions) * 100, 1), 0)
                        report.AppendLine($"   • {kvp.Key}: {kvp.Value} ({percentage}%)")
                    Next
                    report.AppendLine()
                End If

                ' თერაპიის ტიპების დანაწილება
                If stats.TherapyTypeBreakdown.Count > 0 Then
                    report.AppendLine("🩺 თერაპიის ტიპების მიხედვით:")
                    For Each kvp In stats.TherapyTypeBreakdown.OrderByDescending(Function(x) x.Value)
                        Dim percentage = If(stats.TotalSessions > 0, Math.Round((kvp.Value / stats.TotalSessions) * 100, 1), 0)
                        report.AppendLine($"   • {kvp.Key}: {kvp.Value} ({percentage}%)")
                    Next
                End If

                Return report.ToString()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: FormatStatistics შეცდომა: {ex.Message}")
                Return "სტატისტიკის ფორმატირების შეცდომა"
            End Try
        End Function

        ''' <summary>
        ''' სწრაფი სტატისტიკის მიღება (ძირითადი მონაცემები მხოლოდ)
        ''' </summary>
        ''' <param name="criteria">ფილტრის კრიტერიუმები</param>
        ''' <returns>ძირითადი სტატისტიკა</returns>
        Public Function GetQuickStatistics(criteria As ScheduleDataProcessor.FilterCriteria) As (TotalSessions As Integer, CompletedSessions As Integer, TotalRevenue As Decimal)
            Try
                Dim allRows As IList(Of IList(Of Object)) = dataService.GetData("DB-Schedule!A2:O")
                If allRows Is Nothing Then Return (0, 0, 0)

                Dim filteredRows = FilterDataByCriteria(allRows, criteria)
                Dim totalSessions = filteredRows.Count
                Dim completedSessions = 0
                Dim totalRevenue As Decimal = 0

                For Each row In filteredRows
                    Try
                        Dim status = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")
                        If status = "შესრულებული" Then
                            completedSessions += 1
                            If row.Count > 11 Then
                                Dim price As Decimal
                                If Decimal.TryParse(row(11).ToString().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, price) Then
                                    totalRevenue += price
                                End If
                            End If
                        End If
                    Catch
                        Continue For
                    End Try
                Next

                Return (totalSessions, completedSessions, totalRevenue)

            Catch ex As Exception
                Debug.WriteLine($"ScheduleStatisticsService: GetQuickStatistics შეცდომა: {ex.Message}")
                Return (0, 0, 0)
            End Try
        End Function
    End Class
End Namespace