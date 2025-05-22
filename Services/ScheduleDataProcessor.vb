' ===========================================
' 📄 Services/ScheduleDataProcessor.vb
' -------------------------------------------
' განრიგის მონაცემების დამუშავებისა და ფილტრაციის სერვისი
' UC_Schedule-დან გატანილი ფუნქციონალი
' ===========================================
Imports System.Globalization
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის მონაცემების დამუშავებისა და ფილტრაციის სერვისი
    ''' პასუხისმგებელია მონაცემების წამოღება, ფილტრაცია, დალაგება და გვერდებად დაყოფაზე
    ''' </summary>
    Public Class ScheduleDataProcessor

        ' მონაცემთა სერვისი
        Private ReadOnly dataService As IDataService

        ' ქეშირება უკეთესი წარმადობისთვის
        Private lastFilterHash As String = ""
        Private cachedFilteredData As List(Of IList(Of Object)) = Nothing
        Private lastCacheTime As DateTime = DateTime.MinValue
        Private Const CACHE_DURATION_MINUTES As Integer = 2

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="service">მონაცემთა სერვისი</param>
        Public Sub New(service As IDataService)
            dataService = service
        End Sub

        ''' <summary>
        ''' ღირებულების კლასი ფილტრაციის კრიტერიუმებისთვის
        ''' </summary>
        Public Class FilterCriteria
            ''' <summary>საწყისი თარიღი</summary>
            Public Property DateFrom As Date

            ''' <summary>საბოლოო თარიღი</summary>
            Public Property DateTo As Date

            ''' <summary>მონიშნული სტატუსების სია</summary>
            Public Property SelectedStatuses As List(Of String)

            ''' <summary>ბენეფიციარის სახელის ფილტრი</summary>
            Public Property BeneficiaryName As String

            ''' <summary>ბენეფიციარის გვარის ფილტრი</summary>
            Public Property BeneficiarySurname As String

            ''' <summary>თერაპევტის ფილტრი</summary>
            Public Property TherapistName As String

            ''' <summary>თერაპიის ტიპის ფილტრი</summary>
            Public Property TherapyType As String

            ''' <summary>სივრცის ფილტრი</summary>
            Public Property Space As String

            ''' <summary>დაფინანსების ფილტრი</summary>
            Public Property Funding As String

            ''' <summary>
            ''' კონსტრუქტორი ცარიელი ფილტრებით
            ''' </summary>
            Public Sub New()
                DateFrom = DateTime.Today.AddDays(-30)
                DateTo = DateTime.Today
                SelectedStatuses = New List(Of String)()
                BeneficiaryName = ""
                BeneficiarySurname = ""
                TherapistName = ""
                TherapyType = ""
                Space = ""
                Funding = ""
            End Sub

            ''' <summary>
            ''' ფილტრის ჰეშის გენერაცია ქეშირებისთვის
            ''' </summary>
            Public Function GetHashCode() As String
                Dim hashBuilder As New System.Text.StringBuilder()
                hashBuilder.Append($"{DateFrom:yyyyMMdd}_{DateTo:yyyyMMdd}_")
                hashBuilder.Append($"{String.Join(",", SelectedStatuses)}_")
                hashBuilder.Append($"{BeneficiaryName}_{BeneficiarySurname}_")
                hashBuilder.Append($"{TherapistName}_{TherapyType}_")
                hashBuilder.Append($"{Space}_{Funding}")
                Return hashBuilder.ToString()
            End Function
        End Class

        ''' <summary>
        ''' გვერდების მონაცემების კლასი
        ''' </summary>
        Public Class PagedResult
            ''' <summary>მიმდინარე გვერდის მონაცემები</summary>
            Public Property Data As List(Of IList(Of Object))

            ''' <summary>მიმდინარე გვერდის ნომერი</summary>
            Public Property CurrentPage As Integer

            ''' <summary>სულ გვერდების რაოდენობა</summary>
            Public Property TotalPages As Integer

            ''' <summary>სულ ჩანაწერების რაოდენობა</summary>
            Public Property TotalRecords As Integer

            ''' <summary>გვერდის ზომა</summary>
            Public Property PageSize As Integer
        End Class

        ''' <summary>
        ''' ფილტრაციის შედეგების მიღება გვერდებად დაყოფით
        ''' </summary>
        ''' <param name="criteria">ფილტრაციის კრიტერიუმები</param>
        ''' <param name="pageNumber">გვერდის ნომერი (1-დან იწყება)</param>
        ''' <param name="pageSize">გვერდის ზომა</param>
        ''' <returns>გვერდების შედეგი</returns>
        Public Function GetFilteredSchedule(criteria As FilterCriteria, pageNumber As Integer, pageSize As Integer) As PagedResult
            Try
                Debug.WriteLine($"ScheduleDataProcessor: GetFilteredSchedule - გვერდი {pageNumber}, ზომა {pageSize}")

                ' ფილტრაციის შედეგების მიღება
                Dim filteredData = GetFilteredData(criteria)

                ' გვერდების გამოთვლა
                Dim totalRecords = filteredData.Count
                Dim totalPages = Math.Max(1, Math.Ceiling(totalRecords / pageSize))
                Dim validPageNumber = Math.Max(1, Math.Min(pageNumber, totalPages))

                ' მიმდინარე გვერდის მონაცემების მიღება
                Dim pageData = filteredData.Skip((validPageNumber - 1) * pageSize).Take(pageSize).ToList()

                Debug.WriteLine($"ScheduleDataProcessor: სულ {totalRecords} ჩანაწერი, {totalPages} გვერდი, მიმდინარე გვერდი {validPageNumber}")

                Return New PagedResult With {
                    .Data = pageData,
                    .CurrentPage = validPageNumber,
                    .TotalPages = totalPages,
                    .TotalRecords = totalRecords,
                    .PageSize = pageSize
                }

            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: GetFilteredSchedule შეცდომა: {ex.Message}")
                Return New PagedResult With {
                    .Data = New List(Of IList(Of Object))(),
                    .CurrentPage = 1,
                    .TotalPages = 1,
                    .TotalRecords = 0,
                    .PageSize = pageSize
                }
            End Try
        End Function

        ''' <summary>
        ''' ფილტრაციის მონაცემების მიღება ქეშირებით
        ''' </summary>
        ''' <param name="criteria">ფილტრაციის კრიტერიუმები</param>
        ''' <returns>ფილტრირებული მონაცემები</returns>
        Private Function GetFilteredData(criteria As FilterCriteria) As List(Of IList(Of Object))
            Try
                ' ქეშის შემოწმება
                Dim currentHash = criteria.GetHashCode()
                Dim cacheAge = DateTime.Now.Subtract(lastCacheTime).TotalMinutes

                If currentHash = lastFilterHash AndAlso
                   cachedFilteredData IsNot Nothing AndAlso
                   cacheAge < CACHE_DURATION_MINUTES Then
                    Debug.WriteLine("ScheduleDataProcessor: გამოყენებულია ქეშირებული მონაცემები")
                    Return cachedFilteredData
                End If

                ' ახალი მონაცემების მიღება
                Dim filteredData = FilterScheduleData(criteria)

                ' ქეშის განახლება
                lastFilterHash = currentHash
                cachedFilteredData = filteredData
                lastCacheTime = DateTime.Now

                Return filteredData

            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: GetFilteredData შეცდომა: {ex.Message}")
                Return New List(Of IList(Of Object))()
            End Try
        End Function

        ''' <summary>
        ''' განრიგის მონაცემების ფილტრაცია
        ''' </summary>
        ''' <param name="criteria">ფილტრაციის კრიტერიუმები</param>
        ''' <returns>ფილტრირებული მონაცემები</returns>
        Private Function FilterScheduleData(criteria As FilterCriteria) As List(Of IList(Of Object))
            Try
                If dataService Is Nothing Then
                    Debug.WriteLine("ScheduleDataProcessor: dataService არ არის ინიციალიზებული")
                    Return New List(Of IList(Of Object))()
                End If

                ' Google Sheets-დან ყველა ჩანაწერის წამოღება
                Dim allRows As IList(Of IList(Of Object)) = dataService.GetData("DB-Schedule!A2:O")
                Dim filteredRows As New List(Of IList(Of Object))

                If allRows Is Nothing Then
                    Debug.WriteLine("ScheduleDataProcessor: მონაცემები ვერ ჩამოიტვირთა")
                    Return filteredRows
                End If

                Debug.WriteLine($"ScheduleDataProcessor: დამუშავება - სულ {allRows.Count} ჩანაწერი")

                ' ყოველი ჩანაწერის გადარჩევა
                For Each row As IList(Of Object) In allRows
                    If IsRowMatchingCriteria(row, criteria) Then
                        filteredRows.Add(row)
                    End If
                Next

                ' დალაგება თარიღის მიხედვით (ახლიდან ძველისკენ)
                filteredRows = SortByDateDescending(filteredRows)

                Debug.WriteLine($"ScheduleDataProcessor: ფილტრაციის შემდეგ დარჩა {filteredRows.Count} ჩანაწერი")
                Return filteredRows

            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: FilterScheduleData შეცდომა: {ex.Message}")
                Return New List(Of IList(Of Object))()
            End Try
        End Function

        ''' <summary>
        ''' ჩანაწერის შესაბამისობის შემოწმება ფილტრის კრიტერიუმებთან
        ''' </summary>
        ''' <param name="row">შესამოწმებელი ჩანაწერი</param>
        ''' <param name="criteria">ფილტრის კრიტერიუმები</param>
        ''' <returns>True თუ ჩანაწერი შეესაბამება კრიტერიუმებს</returns>
        Private Function IsRowMatchingCriteria(row As IList(Of Object), criteria As FilterCriteria) As Boolean
            Try
                ' ძირითადი ვალიდაცია
                If row.Count < 7 Then Return False ' მინიმუმ უნდა არსებობდეს N-დან ხანგძლიობამდე
                If String.IsNullOrWhiteSpace(row(0).ToString()) Then Return False ' N ცარიელი
                If String.IsNullOrWhiteSpace(row(5).ToString()) Then Return False ' თარიღი ცარიელი
                If String.IsNullOrWhiteSpace(row(3).ToString()) OrElse String.IsNullOrWhiteSpace(row(4).ToString()) Then Return False ' სახელი ან გვარი ცარიელი

                ' თარიღის ფილტრი
                If Not IsDateInRange(row(5).ToString(), criteria.DateFrom, criteria.DateTo) Then Return False

                ' სტატუსის ფილტრი
                If Not IsStatusMatching(row, criteria.SelectedStatuses) Then Return False

                ' ბენეფიციარის სახელის ფილტრი
                If Not IsTextFieldMatching(row, 3, criteria.BeneficiaryName) Then Return False

                ' ბენეფიციარის გვარის ფილტრი
                If Not IsTextFieldMatching(row, 4, criteria.BeneficiarySurname) Then Return False

                ' თერაპევტის ფილტრი
                If Not IsTextFieldMatching(row, 8, criteria.TherapistName) Then Return False

                ' თერაპიის ტიპის ფილტრი
                If Not IsTextFieldMatching(row, 9, criteria.TherapyType) Then Return False

                ' სივრცის ფილტრი
                If Not IsTextFieldMatching(row, 10, criteria.Space) Then Return False

                ' დაფინანსების ფილტრი
                If Not IsTextFieldMatching(row, 13, criteria.Funding) Then Return False

                Return True

            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: IsRowMatchingCriteria შეცდომა: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' თარიღის დიაპაზონში მოქცევის შემოწმება
        ''' </summary>
        Private Function IsDateInRange(dateString As String, dateFrom As Date, dateTo As Date) As Boolean
            Try
                Dim sessionDate As DateTime
                Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                If Not DateTime.TryParseExact(dateString.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, sessionDate) Then
                    If Not DateTime.TryParse(dateString.Trim(), sessionDate) Then
                        Return False
                    End If
                End If

                Return sessionDate.Date >= dateFrom.Date AndAlso sessionDate.Date <= dateTo.Date
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' სტატუსის შესაბამისობის შემოწმება
        ''' </summary>
        Private Function IsStatusMatching(row As IList(Of Object), selectedStatuses As List(Of String)) As Boolean
            Try
                If selectedStatuses.Count = 0 Then Return True

                Dim rowStatus As String = If(row.Count > 12, row(12).ToString().Trim(), "")

                For Each selectedStatus In selectedStatuses
                    If String.Equals(rowStatus, selectedStatus, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next

                Return False
            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' ტექსტური ველის შესაბამისობის შემოწმება
        ''' </summary>
        Private Function IsTextFieldMatching(row As IList(Of Object), columnIndex As Integer, filterValue As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(filterValue) OrElse filterValue = "ყველა" Then Return True

                If row.Count <= columnIndex Then Return True

                Dim rowValue As String = row(columnIndex).ToString().Trim()
                Return String.Equals(rowValue, filterValue, StringComparison.OrdinalIgnoreCase)
            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' მონაცემების დალაგება თარიღის მიხედვით კლებადობით
        ''' </summary>
        Private Function SortByDateDescending(data As List(Of IList(Of Object))) As List(Of IList(Of Object))
            Try
                data.Sort(Function(a, b)
                              Try
                                  Dim dta, dtb As DateTime
                                  Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                                  ' b თარიღის პარსინგი (ახალი)
                                  If Not DateTime.TryParseExact(b(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dtb) Then
                                      If Not DateTime.TryParse(b(5).ToString().Trim(), dtb) Then
                                          dtb = Date.MinValue
                                      End If
                                  End If

                                  ' a თარიღის პარსინგი (ძველი)
                                  If Not DateTime.TryParseExact(a(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dta) Then
                                      If Not DateTime.TryParse(a(5).ToString().Trim(), dta) Then
                                          dta = Date.MinValue
                                      End If
                                  End If

                                  ' კლებადობით დალაგება (ახლიდან ძველისკენ)
                                  Return dtb.CompareTo(dta)
                              Catch
                                  Return 0
                              End Try
                          End Function)

                Return data
            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: SortByDateDescending შეცდომა: {ex.Message}")
                Return data
            End Try
        End Function

        ''' <summary>
        ''' ქეშის გასუფთავება - ახალი მონაცემების ძაღლი ჩატვირთვისთვის
        ''' </summary>
        Public Sub ClearCache()
            lastFilterHash = ""
            cachedFilteredData = Nothing
            lastCacheTime = DateTime.MinValue
            Debug.WriteLine("ScheduleDataProcessor: ქეში გასუფთავებულია")
        End Sub

        ''' <summary>
        ''' უნიკალური მნიშვნელობების მიღება კონკრეტული სვეტისთვის ფილტრების ComboBox-ების შესავსებად
        ''' </summary>
        ''' <param name="columnIndex">სვეტის ინდექსი</param>
        ''' <returns>უნიკალური მნიშვნელობების სია</returns>
        Public Function GetUniqueColumnValues(columnIndex As Integer) As List(Of String)
            Try
                Dim uniqueValues As New HashSet(Of String)()
                Dim allRows As IList(Of IList(Of Object)) = dataService.GetData("DB-Schedule!A2:O")

                If allRows IsNot Nothing Then
                    For Each row In allRows
                        If row.Count > columnIndex AndAlso row(columnIndex) IsNot Nothing Then
                            Dim value = row(columnIndex).ToString().Trim()
                            If Not String.IsNullOrWhiteSpace(value) Then
                                uniqueValues.Add(value)
                            End If
                        End If
                    Next
                End If

                Dim result = uniqueValues.ToList()
                result.Sort()
                result.Insert(0, "ყველა") ' პირველ ადგილას "ყველა" ვარიანტი

                Debug.WriteLine($"ScheduleDataProcessor: სვეტი {columnIndex}-ისთვის ნაპოვნია {result.Count - 1} უნიკალური მნიშვნელობა")
                Return result

            Catch ex As Exception
                Debug.WriteLine($"ScheduleDataProcessor: GetUniqueColumnValues შეცდომა: {ex.Message}")
                Return New List(Of String) From {"ყველა"}
            End Try
        End Function

        ''' <summary>
        ''' ყველა ხელმისაწვდომი სტატუსის მიღება
        ''' </summary>
        ''' <returns>სტატუსების სია</returns>
        Public Function GetAllStatuses() As List(Of String)
            Return GetUniqueColumnValues(12) ' M სვეტი - სტატუსი
        End Function

        ''' <summary>
        ''' ყველა ბენეფიციარის სახელის მიღება
        ''' </summary>
        ''' <returns>სახელების სია</returns>
        Public Function GetAllBeneficiaryNames() As List(Of String)
            Return GetUniqueColumnValues(3) ' D სვეტი - სახელი
        End Function

        ''' <summary>
        ''' ყველა ბენეფიციარის გვარის მიღება
        ''' </summary>
        ''' <returns>გვარების სია</returns>
        Public Function GetAllBeneficiarySurnames() As List(Of String)
            Return GetUniqueColumnValues(4) ' E სვეტი - გვარი
        End Function

        ''' <summary>
        ''' ყველა თერაპევტის სახელის მიღება
        ''' </summary>
        ''' <returns>თერაპევტების სია</returns>
        Public Function GetAllTherapists() As List(Of String)
            Return GetUniqueColumnValues(8) ' I სვეტი - თერაპევტი
        End Function

        ''' <summary>
        ''' ყველა თერაპიის ტიპის მიღება
        ''' </summary>
        ''' <returns>თერაპიის ტიპების სია</returns>
        Public Function GetAllTherapyTypes() As List(Of String)
            Return GetUniqueColumnValues(9) ' J სვეტი - თერაპია
        End Function

        ''' <summary>
        ''' ყველა სივრცის მიღება
        ''' </summary>
        ''' <returns>სივრცეების სია</returns>
        Public Function GetAllSpaces() As List(Of String)
            Return GetUniqueColumnValues(10) ' K სვეტი - სივრცე
        End Function

        ''' <summary>
        ''' ყველა დაფინანსების ტიპის მიღება
        ''' </summary>
        ''' <returns>დაფინანსების ტიპების სია</returns>
        Public Function GetAllFundingTypes() As List(Of String)
            Return GetUniqueColumnValues(13) ' N სვეტი - დაფინანსება
        End Function
    End Class
End Namespace