' ===========================================
' 📄 Models/SessionModel.vb
' -------------------------------------------
' სესიის მოდელი - შეიცავს სესიის ინფორმაციას ავტორისა და რედაქტირების თარიღის ჩათვლით
' შესწორება: რედაქტირების თარიღისა და კომენტარების სწორი დამუშავება v8.2 ვერსიის მსგავსად
' ===========================================
Imports System.ComponentModel
Imports System.Text
Imports System.Globalization

Namespace Scheduler_v8_8a.Models

    ''' <summary>
    ''' SessionModel - სესიის მონაცემების შენახვა და INotifyPropertyChanged ინტერფეისის იმპლემენტაცია
    ''' </summary>
    Public Class SessionModel
        Implements INotifyPropertyChanged

        ' ძირითადი მონაცემების ველები
        Private _id As Integer
        Private _beneficiaryName As String
        Private _beneficiarySurname As String
        Private _dateTime As DateTime
        Private _duration As Integer
        Private _isGroup As Boolean
        Private _therapistName As String
        Private _therapyType As String
        Private _space As String
        Private _price As Decimal
        Private _status As String
        Private _funding As String
        Private _comments As String
        Private _author As String
        Private _lastEditDate As DateTime?

        ''' <summary>სესიის ID</summary>
        Public Property Id As Integer
            Get
                Return _id
            End Get
            Set(value As Integer)
                If _id <> value Then
                    _id = value
                    OnPropertyChanged(NameOf(Id))
                End If
            End Set
        End Property

        ''' <summary>ბენეფიციარის სახელი</summary>
        Public Property BeneficiaryName As String
            Get
                Return _beneficiaryName
            End Get
            Set(value As String)
                If _beneficiaryName <> value Then
                    _beneficiaryName = value
                    OnPropertyChanged(NameOf(BeneficiaryName))
                    OnPropertyChanged(NameOf(FullName))
                End If
            End Set
        End Property

        ''' <summary>ბენეფიციარის გვარი</summary>
        Public Property BeneficiarySurname As String
            Get
                Return _beneficiarySurname
            End Get
            Set(value As String)
                If _beneficiarySurname <> value Then
                    _beneficiarySurname = value
                    OnPropertyChanged(NameOf(BeneficiarySurname))
                    OnPropertyChanged(NameOf(FullName))
                End If
            End Set
        End Property

        ''' <summary>სრული სახელი (სახელი + გვარი)</summary>
        Public ReadOnly Property FullName As String
            Get
                Return $"{_beneficiaryName} {_beneficiarySurname}"
            End Get
        End Property

        ''' <summary>სესიის თარიღი და დრო</summary>
        Public Property DateTime As DateTime
            Get
                Return _dateTime
            End Get
            Set(value As DateTime)
                If _dateTime <> value Then
                    _dateTime = value
                    OnPropertyChanged(NameOf(DateTime))
                    OnPropertyChanged(NameOf(FormattedDateTime))
                End If
            End Set
        End Property

        ''' <summary>ფორმატირებული თარიღი და დრო (dd.MM.yyyy HH:mm)</summary>
        Public ReadOnly Property FormattedDateTime As String
            Get
                Return _dateTime.ToString("dd.MM.yyyy HH:mm")
            End Get
        End Property

        ''' <summary>სესიის ხანგრძლივობა წუთებში</summary>
        Public Property Duration As Integer
            Get
                Return _duration
            End Get
            Set(value As Integer)
                If _duration <> value Then
                    _duration = value
                    OnPropertyChanged(NameOf(Duration))
                End If
            End Set
        End Property

        ''' <summary>არის თუ არა ჯგუფური სესია</summary>
        Public Property IsGroup As Boolean
            Get
                Return _isGroup
            End Get
            Set(value As Boolean)
                If _isGroup <> value Then
                    _isGroup = value
                    OnPropertyChanged(NameOf(IsGroup))
                End If
            End Set
        End Property

        ''' <summary>თერაპევტის სახელი</summary>
        Public Property TherapistName As String
            Get
                Return _therapistName
            End Get
            Set(value As String)
                If _therapistName <> value Then
                    _therapistName = value
                    OnPropertyChanged(NameOf(TherapistName))
                End If
            End Set
        End Property

        ''' <summary>თერაპიის ტიპი</summary>
        Public Property TherapyType As String
            Get
                Return _therapyType
            End Get
            Set(value As String)
                If _therapyType <> value Then
                    _therapyType = value
                    OnPropertyChanged(NameOf(TherapyType))
                End If
            End Set
        End Property

        ''' <summary>სივრცე/ოთახი სადაც ტარდება სესია</summary>
        Public Property Space As String
            Get
                Return _space
            End Get
            Set(value As String)
                If _space <> value Then
                    _space = value
                    OnPropertyChanged(NameOf(Space))
                End If
            End Set
        End Property

        ''' <summary>სესიის ფასი</summary>
        Public Property Price As Decimal
            Get
                Return _price
            End Get
            Set(value As Decimal)
                If _price <> value Then
                    _price = value
                    OnPropertyChanged(NameOf(Price))
                End If
            End Set
        End Property

        ''' <summary>სესიის სტატუსი (შესრულებული, გაუქმება, დაგეგმილი)</summary>
        Public Property Status As String
            Get
                Return _status
            End Get
            Set(value As String)
                If _status <> value Then
                    _status = value
                    OnPropertyChanged(NameOf(Status))
                End If
            End Set
        End Property

        ''' <summary>დაფინანსების ტიპი</summary>
        Public Property Funding As String
            Get
                Return _funding
            End Get
            Set(value As String)
                If _funding <> value Then
                    _funding = value
                    OnPropertyChanged(NameOf(Funding))
                End If
            End Set
        End Property

        ''' <summary>დამატებითი კომენტარები სესიაზე</summary>
        Public Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                If _comments <> value Then
                    _comments = value
                    OnPropertyChanged(NameOf(Comments))
                End If
            End Set
        End Property

        ''' <summary>ჩანაწერის ავტორი</summary>
        Public Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                If _author <> value Then
                    _author = value
                    OnPropertyChanged(NameOf(Author))
                End If
            End Set
        End Property

        ''' <summary>ბოლო რედაქტირების თარიღი</summary>
        Public Property LastEditDate As DateTime?
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime?)
                If _lastEditDate <> value Then
                    _lastEditDate = value
                    OnPropertyChanged(NameOf(LastEditDate))
                    OnPropertyChanged(NameOf(FormattedLastEditDate))
                End If
            End Set
        End Property

        ''' <summary>ფორმატირებული რედაქტირების თარიღი</summary>
        Public ReadOnly Property FormattedLastEditDate As String
            Get
                If _lastEditDate.HasValue Then
                    Return _lastEditDate.Value.ToString("dd.MM.yyyy HH:mm")
                Else
                    Return "-"
                End If
            End Get
        End Property

        ''' <inheritdoc/>
        Public Event PropertyChanged As PropertyChangedEventHandler _
            Implements INotifyPropertyChanged.PropertyChanged

        ''' <summary>
        ''' იძახებს PropertyChanged ივენთს როდესაც თვისება იცვლება
        ''' </summary>
        Protected Sub OnPropertyChanged(propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი - ცარიელი
        ''' </summary>
        Public Sub New()
            ' საწყისი მნიშვნელობების ინიციალიზაცია
            _beneficiaryName = String.Empty
            _beneficiarySurname = String.Empty
            _dateTime = DateTime.Now
            _duration = 60
            _isGroup = False
            _therapistName = String.Empty
            _therapyType = String.Empty
            _space = String.Empty
            _price = 0
            _status = "დაგეგმილი"
            _funding = "კერძო"
            _comments = String.Empty
            _author = String.Empty
            _lastEditDate = Nothing
        End Sub

        ''' <summary>
        ''' შექმნის SessionModel-ს Google Sheets-დან მოცემული მწკრივიდან
        ''' გამართული ვერსია v8.2-ის მსგავსი რედაქტირების თარიღის დამუშავებით
        ''' </summary>
        ''' <param name="rowData">მწკრივი Google Sheets-დან</param>
        ''' <returns>SessionModel შევსებული მონაცემებით</returns>
        Public Shared Function FromSheetRow(rowData As IList(Of Object)) As SessionModel
            ' შემოწმება მონაცემების რაოდენობის
            If rowData Is Nothing OrElse rowData.Count < 13 Then
                Throw New ArgumentException("არასაკმარისი მონაცემები SessionModel-ისთვის")
            End If

            Try
                Dim session As New SessionModel()

                ' A სვეტი: ID-ის პარსინგი
                Dim idStr = rowData(0).ToString().Trim()
                Dim sessionId As Integer = 0
                If Not String.IsNullOrWhiteSpace(idStr) AndAlso Integer.TryParse(idStr, sessionId) Then
                    session.Id = sessionId
                Else
                    session.Id = 0
                End If

                ' B სვეტი: რედაქტირების თარიღი - გამართული ვერსია v8.2-ის მსგავსად
                If rowData.Count > 1 AndAlso rowData(1) IsNot Nothing Then
                    Dim editDateStr = rowData(1).ToString().Trim()

                    If Not String.IsNullOrWhiteSpace(editDateStr) Then
                        Try
                            ' Google Sheets-ის ყველაზე გავრცელებული ფორმატები
                            Dim editFormats As String() = {
                                "dd.MM.yyyy HH:mm", "d.M.yyyy HH:mm", "dd.M.yyyy HH:mm", "d.MM.yyyy HH:mm",
                                "dd.MM.yyyy H:mm", "d.M.yyyy H:mm", "dd.M.yyyy H:mm", "d.MM.yyyy H:mm",
                                "dd.MM.yy HH:mm", "d.M.yy HH:mm", "dd.M.yy HH:mm", "d.MM.yy HH:mm",
                                "dd.MM.yy H:mm", "d.M.yy H:mm", "dd.M.yy H:mm", "d.MM.yy H:mm",
                                "d/M/yyyy H:mm:ss", "dd/M/yyyy H:mm:ss", "d/MM/yyyy H:mm:ss", "dd/MM/yyyy H:mm:ss",
                                "d/M/yyyy HH:mm:ss", "dd/M/yyyy HH:mm:ss", "d/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss",
                                "d/M/yyyy H:mm", "dd/M/yyyy H:mm", "d/MM/yyyy H:mm", "dd/MM/yyyy H:mm",
                                "d/M/yyyy HH:mm", "dd/M/yyyy HH:mm", "d/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm",
                                "MM/dd/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm",
                                "dd.MM.yyyy", "d.M.yyyy", "dd.M.yyyy", "d.MM.yyyy",
                                "dd.MM.yy", "d.M.yy", "dd.M.yy", "d.MM.yy"
                            }

                            Dim parsedEditDate As DateTime

                            ' ზუსტი ფორმატით პარსინგის მცდელობა
                            If DateTime.TryParseExact(editDateStr, editFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, parsedEditDate) Then
                                session.LastEditDate = parsedEditDate
                                Debug.WriteLine($"SessionModel: რედაქტირების თარიღი დაპარსილია (ზუსტი): {parsedEditDate:dd.MM.yyyy HH:mm}")
                                ' ზოგადი პარსინგის მცდელობა
                            ElseIf DateTime.TryParse(editDateStr, parsedEditDate) Then
                                session.LastEditDate = parsedEditDate
                                Debug.WriteLine($"SessionModel: რედაქტირების თარიღი დაპარსილია (ზოგადი): {parsedEditDate:dd.MM.yyyy HH:mm}")
                            Else
                                session.LastEditDate = Nothing
                                Debug.WriteLine($"SessionModel: რედაქტირების თარიღის პარსინგი ვერ მოხერხდა: '{editDateStr}'")
                            End If
                        Catch ex As Exception
                            session.LastEditDate = Nothing
                            Debug.WriteLine($"SessionModel: რედაქტირების თარიღის პარსინგის შეცდომა: {ex.Message}")
                        End Try
                    Else
                        session.LastEditDate = Nothing
                    End If
                Else
                    session.LastEditDate = Nothing
                End If

                ' C სვეტი: ავტორი
                session.Author = If(rowData.Count > 2, rowData(2).ToString().Trim(), String.Empty)

                ' D და E სვეტები: ბენეფიციარის სახელი და გვარი
                session.BeneficiaryName = If(rowData.Count > 3, rowData(3).ToString().Trim(), String.Empty)
                session.BeneficiarySurname = If(rowData.Count > 4, rowData(4).ToString().Trim(), String.Empty)

                ' F სვეტი: სესიის თარიღი - გამართული ვერსია
                If rowData.Count > 5 AndAlso rowData(5) IsNot Nothing Then
                    Dim dateTimeStr = rowData(5).ToString().Trim()

                    If Not String.IsNullOrWhiteSpace(dateTimeStr) Then
                        Try
                            ' Google Sheets-ის სესიის თარიღის ფორმატები
                            Dim sessionFormats As String() = {
                                "dd.MM.yyyy HH:mm", "d.M.yyyy HH:mm", "dd.M.yyyy HH:mm", "d.MM.yyyy HH:mm",
                                "dd.MM.yyyy H:mm", "d.M.yyyy H:mm", "dd.M.yyyy H:mm", "d.MM.yyyy H:mm",
                                "dd.MM.yy HH:mm", "d.M.yy HH:mm", "dd.M.yy HH:mm", "d.MM.yy HH:mm",
                                "dd.MM.yy H:mm", "d.M.yy H:mm", "dd.M.yy H:mm", "d.MM.yy H:mm",
                                "d/M/yyyy H:mm:ss", "dd/M/yyyy H:mm:ss", "d/MM/yyyy H:mm:ss", "dd/MM/yyyy H:mm:ss",
                                "d/M/yyyy HH:mm:ss", "dd/M/yyyy HH:mm:ss", "d/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss",
                                "d/M/yyyy H:mm", "dd/M/yyyy H:mm", "d/MM/yyyy H:mm", "dd/MM/yyyy H:mm",
                                "d/M/yyyy HH:mm", "dd/M/yyyy HH:mm", "d/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm",
                                "MM/dd/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm",
                                "dd.MM.yyyy", "d.M.yyyy", "dd.M.yyyy", "d.MM.yyyy",
                                "dd.MM.yy", "d.M.yy", "dd.M.yy", "d.MM.yy"
                            }

                            Dim parsedDate As DateTime

                            ' ზუსტი ფორმატით პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, sessionFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                                session.DateTime = parsedDate
                                ' ზოგადი პარსინგის მცდელობა
                            ElseIf DateTime.TryParse(dateTimeStr, parsedDate) Then
                                session.DateTime = parsedDate
                            Else
                                ' თუ ვერ დავამუშავეთ, გამოვიყენოთ default თარიღი
                                session.DateTime = New DateTime(1900, 1, 1)
                                Debug.WriteLine($"SessionModel: სესიის თარიღის პარსინგი ვერ მოხერხდა: '{dateTimeStr}'")
                            End If
                        Catch ex As Exception
                            session.DateTime = New DateTime(1900, 1, 1)
                            Debug.WriteLine($"SessionModel: სესიის თარიღის პარსინგის შეცდომა: {ex.Message}")
                        End Try
                    Else
                        session.DateTime = New DateTime(1900, 1, 1)
                    End If
                Else
                    session.DateTime = New DateTime(1900, 1, 1)
                End If

                ' G სვეტი: ხანგრძლივობა
                If rowData.Count > 6 AndAlso rowData(6) IsNot Nothing Then
                    Dim durationStr = rowData(6).ToString().Trim()
                    Dim duration As Integer = 60
                    If Integer.TryParse(durationStr, duration) Then
                        session.Duration = duration
                    Else
                        session.Duration = 60
                    End If
                Else
                    session.Duration = 60
                End If

                ' H სვეტი: ჯგუფური
                If rowData.Count > 7 AndAlso rowData(7) IsNot Nothing Then
                    Dim groupStr = rowData(7).ToString().Trim().ToLower()
                    session.IsGroup = (groupStr = "true" OrElse groupStr = "1" OrElse groupStr = "yes")
                Else
                    session.IsGroup = False
                End If

                ' I სვეტი: თერაპევტი
                session.TherapistName = If(rowData.Count > 8, rowData(8).ToString().Trim(), String.Empty)

                ' J სვეტი: თერაპიის ტიპი
                session.TherapyType = If(rowData.Count > 9, rowData(9).ToString().Trim(), String.Empty)

                ' K სვეტი: სივრცე
                session.Space = If(rowData.Count > 10, rowData(10).ToString().Trim(), String.Empty)

                ' L სვეტი: ფასი
                If rowData.Count > 11 AndAlso rowData(11) IsNot Nothing Then
                    Dim priceStr = rowData(11).ToString().Replace(",", ".").Trim()
                    Dim price As Decimal = 0
                    If Decimal.TryParse(priceStr, NumberStyles.Any, CultureInfo.InvariantCulture, price) Then
                        session.Price = price
                    Else
                        session.Price = 0
                    End If
                Else
                    session.Price = 0
                End If

                ' M სვეტი: სტატუსი
                session.Status = If(rowData.Count > 12, rowData(12).ToString().Trim(), "დაგეგმილი")

                ' N სვეტი: დაფინანსება
                session.Funding = If(rowData.Count > 13, rowData(13).ToString().Trim(), String.Empty)

                ' O სვეტი: კომენტარი - გამართული დამუშავება
                If rowData.Count > 14 AndAlso rowData(14) IsNot Nothing Then
                    Dim commentStr = rowData(14).ToString().Trim()
                    ' თუ კომენტარი არის "0" ან ცარიელი, ვუყენებთ ცარიელ სტრიქონს
                    If commentStr = "0" OrElse String.IsNullOrWhiteSpace(commentStr) Then
                        session.Comments = String.Empty
                    Else
                        session.Comments = commentStr
                    End If
                Else
                    session.Comments = String.Empty
                End If

                Return session

            Catch ex As Exception
                Throw New Exception($"შეცდომა SessionModel-ის შექმნისას: {ex.Message}", ex)
            End Try
        End Function

        ''' <summary>
        ''' გარდაქმნის SessionModel-ს მწკრივად Google Sheets-ისთვის
        ''' გამართული ვერსია რედაქტირების თარიღისა და კომენტარების სწორი ჩაწერისთვის
        ''' </summary>
        ''' <returns>ობიექტების სია, რომელიც შეიძლება გადაეცეს Sheets API-ს</returns>
        Public Function ToSheetRow() As List(Of Object)
            Dim rowData As New List(Of Object)

            ' A სვეტი: ID
            rowData.Add(Id)

            ' B სვეტი: რედაქტირების თარიღი - ახლა ვაყენებთ მიმდინარე თარიღს
            rowData.Add(DateTime.Now.ToString("dd.MM.yyyy HH:mm"))

            ' C სვეტი: ავტორი
            rowData.Add(Author)

            ' D სვეტი: ბენეფიციარის სახელი
            rowData.Add(BeneficiaryName)

            ' E სვეტი: ბენეფიციარის გვარი
            rowData.Add(BeneficiarySurname)

            ' F სვეტი: სესიის თარიღი
            rowData.Add(DateTime.ToString("dd.MM.yyyy HH:mm"))

            ' G სვეტი: ხანგრძლივობა
            rowData.Add(Duration)

            ' H სვეტი: ჯგუფური
            rowData.Add(IsGroup)

            ' I სვეტი: თერაპევტი
            rowData.Add(TherapistName)

            ' J სვეტი: თერაპიის ტიპი
            rowData.Add(TherapyType)

            ' K სვეტი: სივრცე
            rowData.Add(Space)

            ' L სვეტი: ფასი
            rowData.Add(Price)

            ' M სვეტი: სტატუსი
            rowData.Add(Status)

            ' N სვეტი: დაფინანსება
            rowData.Add(Funding)

            ' O სვეტი: კომენტარი
            rowData.Add(If(String.IsNullOrWhiteSpace(Comments), "", Comments))

            Return rowData
        End Function

        ''' <summary>
        ''' არის თუ არა სესია ვადაგადაცილებული
        ''' </summary>
        Public ReadOnly Property IsOverdue As Boolean
            Get
                Try
                    ' მიმდინარე თარიღი და დრო
                    Dim currentDate = DateTime.Today
                    Dim currentDateTime = DateTime.Now

                    ' სესიის თარიღი (მხოლოდ თარიღის ნაწილი, დროის გარეშე)
                    Dim sessionDate = Me.DateTime.Date

                    ' სტატუსის შემოწმება - მხოლოდ "დაგეგმილი" სტატუსისთვის
                    Dim normalizedStatus = Me.Status.Trim().ToLower()
                    Dim isPlanned = (normalizedStatus = "დაგეგმილი" OrElse normalizedStatus = "დაგეგმილი ")

                    ' შემოწმება 1: თუ სესიის თარიღი წარსულში - უკვე გასული დღეა
                    Dim isPastDate = sessionDate < currentDate

                    ' შემოწმება 2: თუ სესიის თარიღი დღევანდელია, მაგრამ მისი დრო უკვე გასულია
                    Dim isTodayPastTime = (sessionDate = currentDate) AndAlso (Me.DateTime < currentDateTime)

                    ' ვადაგადაცილებულია, თუ:
                    ' 1. სტატუსი არის "დაგეგმილი" და
                    ' 2. ან სესიის თარიღი უკვე გასულია, ან დღევანდელი სესიის დრო გასულია
                    Dim result = isPlanned AndAlso (isPastDate OrElse isTodayPastTime)

                    Return result
                Catch ex As Exception
                    Debug.WriteLine($"SessionModel.IsOverdue: შეცდომა - {ex.Message}")
                    Return False
                End Try
            End Get
        End Property
    End Class
End Namespace