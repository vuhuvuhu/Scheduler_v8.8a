' ===========================================
' 📄 Models/SessionModel.vb
' -------------------------------------------
' სესიის მოდელი - შეიცავს სესიის ინფორმაციას ავტორისა და რედაქტირების თარიღის ჩათვლით
' შესწორება: რედაქტირების თარიღისა და კომენტარების ნორმალური დამუშავება (დებაგის გარეშე)
' ===========================================
Imports System.ComponentModel
Imports System.Text

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
                    OnPropertyChanged(NameOf(FullName)) ' ასევე განაახლე FullName თვისება
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
                    OnPropertyChanged(NameOf(FullName)) ' ასევე განაახლე FullName თვისება
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
                    OnPropertyChanged(NameOf(FormattedDateTime)) ' განახლდეს ფორმატირებული თარიღიც
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

        ''' <summary>ფორმატირებული რედაქტირების თარიღი - შესწორებული ვერსია</summary>
        Public ReadOnly Property FormattedLastEditDate As String
            Get
                If _lastEditDate.HasValue Then
                    Return _lastEditDate.Value.ToString("dd.MM.yyyy HH:mm")
                Else
                    Return "" ' ცარიელი სტრიქონი, რომელიც UC_Schedule-ში "-"-ად შეიცვლება
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
            _duration = 60 ' ნაგულისხმები 60 წუთი
            _isGroup = False
            _therapistName = String.Empty
            _therapyType = String.Empty
            _space = String.Empty
            _price = 0
            _status = "დაგეგმილი" ' ნაგულისხმები სტატუსი
            _funding = "კერძო" ' ნაგულისხმები დაფინანსება
            _comments = String.Empty
            _author = String.Empty ' ნაგულისხმები ავტორი
            _lastEditDate = Nothing ' ნაგულისხმები რედაქტირების თარიღი
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი პარამეტრებით - განახლებული ავტორისა და რედაქტირების თარიღით
        ''' </summary>
        Public Sub New(id As Integer, beneficiaryName As String, beneficiarySurname As String,
                      dateTime As DateTime, duration As Integer, isGroup As Boolean,
                      therapistName As String, therapyType As String, space As String,
                      price As Decimal, status As String, funding As String,
                      Optional author As String = "", Optional lastEditDate As DateTime? = Nothing)
            Me.Id = id
            Me.BeneficiaryName = beneficiaryName
            Me.BeneficiarySurname = beneficiarySurname
            Me.DateTime = dateTime
            Me.Duration = duration
            Me.IsGroup = isGroup
            Me.TherapistName = therapistName
            Me.TherapyType = therapyType
            Me.Space = space
            Me.Price = price
            Me.Status = status
            Me.Funding = funding
            Me.Comments = String.Empty
            Me.Author = author
            Me.LastEditDate = lastEditDate
        End Sub

        ''' <summary>
        ''' შექმნის SessionModel-ს Google Sheets-დან მოცემული მწკრივიდან
        ''' შესწორებული ვერსია რედაქტირების თარიღისა და კომენტარების სწორი დამუშავებისთვის (დებაგის გარეშე)
        ''' </summary>
        ''' <param name="rowData">მწკრივი Google Sheets-დან</param>
        ''' <returns>SessionModel შევსებული მონაცემებით</returns>
        Public Shared Function FromSheetRow(rowData As IList(Of Object)) As SessionModel
            ' შემოწმება მონაცემების რაოდენობის
            If rowData Is Nothing OrElse rowData.Count < 13 Then ' M სვეტისთვის (ინდექსი 12) გვჭირდება მინიმუმ 13 ელემენტი
                Throw New ArgumentException("არასაკმარისი მონაცემები SessionModel-ისთვის")
            End If

            Try
                Dim session As New SessionModel()

                ' A სვეტი: ID-ის პარსინგი
                Dim idStr = rowData(0).ToString().Trim()
                Dim sessionId As Integer = 0

                If Not String.IsNullOrWhiteSpace(idStr) AndAlso Integer.TryParse(idStr, sessionId) Then
                    ' ID წარმატებით დაპარსილია
                Else
                    ' ID-ის პარსინგის შეცდომა, ვიყენებთ 0
                    sessionId = 0
                End If

                session.Id = sessionId

                ' C სვეტი: ავტორი
                session.Author = If(rowData.Count > 2, rowData(2).ToString().Trim(), String.Empty)

                ' B სვეტი: რედაქტირების თარიღი - შესწორებული Google Sheets ფორმატებისთვის
                Dim editDateStr = If(rowData.Count > 1, rowData(1).ToString().Trim(), String.Empty)

                Try
                    If Not String.IsNullOrWhiteSpace(editDateStr) Then
                        ' Google Sheets-ის ფორმატები: "21/5/2025 19:10:00", "21.05.25", "19:10"
                        Dim editFormats As String() = {
                            "dd.MM.yy, HH:mm", "d.MM.yy, HH:mm", "dd.M.yy, HH:mm", "d.M.yy, HH:mm",
                            "dd.MM.yy, H:mm", "d.MM.yy, H:mm", "dd.M.yy, H:mm", "d.M.yy, H:mm",
                            "dd.MM.yyyy, HH:mm", "d.MM.yyyy, HH:mm", "dd.M.yyyy, HH:mm", "d.M.yyyy, HH:mm",
                            "dd.MM.yyyy, H:mm", "d.MM.yyyy, H:mm", "dd.M.yyyy, H:mm", "d.M.yyyy, H:mm",
                            "d/M/yyyy H:mm:ss", "dd/M/yyyy H:mm:ss", "d/MM/yyyy H:mm:ss", "dd/MM/yyyy H:mm:ss",
                            "d/M/yyyy HH:mm:ss", "dd/M/yyyy HH:mm:ss", "d/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss",
                            "d/M/yyyy H:mm", "dd/M/yyyy H:mm", "d/MM/yyyy H:mm", "dd/MM/yyyy H:mm",
                            "d/M/yyyy HH:mm", "dd/M/yyyy HH:mm", "d/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm",
                            "dd.MM.yy", "d.MM.yy", "dd.M.yy", "d.M.yy",
                            "dd.MM.yyyy HH:mm", "d.M.yyyy HH:mm", "dd/MM/yyyy HH:mm", "d/MM/yyyy HH:mm",
                            "dd.MM.yyyy H:mm", "d.M.yyyy H:mm", "dd/MM/yyyy H:mm", "d/MM/yyyy H:mm",
                            "MM/dd/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss", "dd.MM.yyyy", "d.M.yyyy",
                            "HH:mm", "H:mm"
                        }

                        Dim parsedEditDate As DateTime
                        If DateTime.TryParseExact(editDateStr, editFormats,
                                      System.Globalization.CultureInfo.InvariantCulture,
                                      System.Globalization.DateTimeStyles.None,
                                      parsedEditDate) Then
                            ' თუ მხოლოდ დრო იყო (HH:mm), დღევანდელ თარიღს ვუმატებთ
                            If editDateStr.Contains(":") AndAlso Not editDateStr.Contains("/") AndAlso Not editDateStr.Contains(".") Then
                                parsedEditDate = DateTime.Today.Add(parsedEditDate.TimeOfDay)
                            End If
                            session.LastEditDate = parsedEditDate
                        ElseIf DateTime.TryParse(editDateStr, parsedEditDate) Then
                            session.LastEditDate = parsedEditDate
                        Else
                            session.LastEditDate = Nothing
                        End If
                    Else
                        session.LastEditDate = Nothing
                    End If
                Catch ex As Exception
                    session.LastEditDate = Nothing
                End Try

                ' D და E სვეტები: ბენეფიციარის სახელი და გვარი
                session.BeneficiaryName = If(rowData.Count > 3, rowData(3).ToString(), String.Empty)
                session.BeneficiarySurname = If(rowData.Count > 4, rowData(4).ToString(), String.Empty)

                ' F სვეტი: სესიის თარიღი - Google Sheets ფორმატებისთვის შესწორებული
                Dim dateTimeStr = If(rowData.Count > 5, rowData(5).ToString().Trim(), String.Empty)

                ' თარიღის დამუშავების კოდი
                Try
                    If Not String.IsNullOrWhiteSpace(dateTimeStr) Then
                        ' Google Sheets-ის ფორმატები: "21/5/2025 19:10:00", "21.05.25", "19:10"
                        Dim sessionFormats As String() = {
                            "dd.MM.yy, HH:mm", "d.MM.yy, HH:mm", "dd.M.yy, HH:mm", "d.M.yy, HH:mm",
                            "dd.MM.yy, H:mm", "d.MM.yy, H:mm", "dd.M.yy, H:mm", "d.M.yy, H:mm",
                            "dd.MM.yyyy, HH:mm", "d.MM.yyyy, HH:mm", "dd.M.yyyy, HH:mm", "d.M.yyyy, HH:mm",
                            "dd.MM.yyyy, H:mm", "d.MM.yyyy, H:mm", "dd.M.yyyy, H:mm", "d.M.yyyy, H:mm",
                            "d/M/yyyy H:mm:ss", "dd/M/yyyy H:mm:ss", "d/MM/yyyy H:mm:ss", "dd/MM/yyyy H:mm:ss",
                            "d/M/yyyy HH:mm:ss", "dd/M/yyyy HH:mm:ss", "d/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss",
                            "d/M/yyyy H:mm", "dd/M/yyyy H:mm", "d/MM/yyyy H:mm", "dd/MM/yyyy H:mm",
                            "d/M/yyyy HH:mm", "dd/M/yyyy HH:mm", "d/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm",
                            "dd.MM.yy", "d.MM.yy", "dd.M.yy", "d.M.yy",
                            "dd.MM.yyyy HH:mm", "d.M.yyyy HH:mm", "dd/MM/yyyy HH:mm", "d/MM/yyyy HH:mm",
                            "dd.MM.yyyy H:mm", "d.M.yyyy H:mm", "dd/MM/yyyy H:mm", "d/MM/yyyy H:mm",
                            "MM/dd/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss", "dd.MM.yyyy", "d.M.yyyy",
                            "HH:mm", "H:mm"
                        }

                        Dim parsedDate As DateTime
                        Dim successfullyParsed As Boolean = False

                        ' ვცადოთ Google Sheets-ის ფორმატები
                        If DateTime.TryParseExact(dateTimeStr, sessionFormats,
                                      System.Globalization.CultureInfo.InvariantCulture,
                                      System.Globalization.DateTimeStyles.None,
                                      parsedDate) Then
                            session.DateTime = parsedDate
                            successfullyParsed = True
                        End If

                        ' თუ ფორმატები ვერ გაართვეს თავი, ვცადოთ საშუალო მიდგომა
                        If Not successfullyParsed Then
                            If DateTime.TryParse(dateTimeStr, parsedDate) Then
                                session.DateTime = parsedDate
                                successfullyParsed = True
                            End If
                        End If

                        ' თუ მაინც ვერ დავამუშავეთ, გამოვიყენოთ default თარიღი
                        If Not successfullyParsed Then
                            session.DateTime = New DateTime(1900, 1, 1) ' ძალიან ძველი თარიღი
                        End If
                    Else
                        ' თარიღის სტრიქონი ცარიელია
                        session.DateTime = New DateTime(1900, 1, 1)
                    End If
                Catch ex As Exception
                    ' თარიღის დამუშავების შეცდომა
                    session.DateTime = New DateTime(1900, 1, 1)
                End Try

                ' G სვეტი: ხანგრძლივობა
                session.Duration = If(rowData.Count > 6 AndAlso Integer.TryParse(rowData(6).ToString(), Nothing), Integer.Parse(rowData(6).ToString()), 60)

                ' H სვეტი: ჯგუფური
                If rowData.Count > 7 Then
                    Dim groupStr = rowData(7).ToString().Trim().ToLower()
                    session.IsGroup = (groupStr = "true" OrElse groupStr = "1")
                Else
                    session.IsGroup = False
                End If

                ' I სვეტი: თერაპევტი
                session.TherapistName = If(rowData.Count > 8, rowData(8).ToString(), String.Empty)

                ' J სვეტი: თერაპიის ტიპი
                session.TherapyType = If(rowData.Count > 9, rowData(9).ToString(), String.Empty)

                ' K სვეტი: სივრცე
                session.Space = If(rowData.Count > 10, rowData(10).ToString(), String.Empty)

                ' L სვეტი: ფასი
                If rowData.Count > 11 Then
                    Dim priceStr = rowData(11).ToString().Replace(",", ".")
                    Dim price As Decimal = 0
                    If Decimal.TryParse(priceStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, price) Then
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
                session.Funding = If(rowData.Count > 13, rowData(13).ToString(), String.Empty)

                ' O სვეტი: კომენტარი - შესწორებული დამუშავება
                If rowData.Count > 14 Then
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
        ''' შესწორებული ვერსია რედაქტირების თარიღისა და კომენტარების სწორი ჩაწერისთვის
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
            End Get
        End Property
    End Class
End Namespace