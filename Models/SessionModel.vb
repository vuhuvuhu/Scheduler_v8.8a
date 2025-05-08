' ===========================================
' 📄 Models/SessionModel.vb
' -------------------------------------------
' სესიის მოდელი - შეიცავს სესიის ინფორმაციას
' ===========================================
Imports System.ComponentModel

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
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი პარამეტრებით
        ''' </summary>
        Public Sub New(id As Integer, beneficiaryName As String, beneficiarySurname As String,
                      dateTime As DateTime, duration As Integer, isGroup As Boolean,
                      therapistName As String, therapyType As String, space As String,
                      price As Decimal, status As String, funding As String)
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
        End Sub

        ''' <summary>
        ''' შექმნის SessionModel-ს იღებს რა მონაცემთა მასივს Google Sheets-დან
        ''' </summary>
        ''' <param name="rowData">მწკრივი Google Sheets-დან</param>
        ''' <returns>SessionModel შევსებული მონაცემებით</returns>
        Public Shared Function FromSheetRow(rowData As IList(Of Object)) As SessionModel
            ' შემოწმება მონაცემების რაოდენობის
            If rowData Is Nothing OrElse rowData.Count < 12 Then
                Throw New ArgumentException("არასაკმარისი მონაცემები SessionModel-ისთვის")
            End If

            Try
                Dim session As New SessionModel()

                ' ID და სხვა მონაცემების დაყენება - დამატებით შემოწმებით
                Try
                    session.Id = Integer.Parse(rowData(0).ToString().Trim())
                Catch ex As Exception
                    Debug.WriteLine($"ID-ის პარსინგის შეცდომა: {rowData(0)} - {ex.Message}")
                    session.Id = 0 ' ნაგულისხმევი ID
                End Try

                ' სახელები
                Try
                    session.BeneficiaryName = If(rowData.Count > 2, rowData(2).ToString().Trim(), "")
                    session.BeneficiarySurname = If(rowData.Count > 3, rowData(3).ToString().Trim(), "")
                Catch ex As Exception
                    Debug.WriteLine($"სახელის პარსინგის შეცდომა: {ex.Message}")
                    session.BeneficiaryName = ""
                    session.BeneficiarySurname = ""
                End Try

                ' თარიღის და დროის პარსინგი - უფრო მდგრადი მიდგომა
                Try
                    If rowData.Count > 4 Then
                        Dim dateTimeStr = rowData(4).ToString().Trim()
                        Debug.WriteLine($"SessionModel.FromSheetRow: სესიის თარიღი='{dateTimeStr}'")

                        ' თუ ცარიელია, გამოვიყენოთ მიმდინარე დრო
                        If String.IsNullOrEmpty(dateTimeStr) Then
                            session.DateTime = DateTime.Now
                        Else
                            ' ფიქსირებული ფორმატის მიდგომა
                            Try
                                Dim parts = dateTimeStr.Split(" "c)
                                If parts.Length >= 2 Then
                                    Dim datePart = parts(0).Trim()
                                    Dim timePart = parts(1).Trim()

                                    ' თარიღის ნაწილების დაპარსვა (დღე.თვე.წელი)
                                    Dim datePieces = datePart.Split("."c)
                                    If datePieces.Length >= 3 Then
                                        Dim day = Integer.Parse(datePieces(0))
                                        Dim month = Integer.Parse(datePieces(1))
                                        Dim year = Integer.Parse(datePieces(2))

                                        ' თუ წელი 2-ნიშნაა, დავამატოთ 2000
                                        If year < 100 Then
                                            year += 2000
                                        End If

                                        ' დროის ნაწილების დაპარსვა (საათი:წუთი)
                                        Dim timePieces = timePart.Replace(",", "").Split(":"c)
                                        Dim hour = If(timePieces.Length >= 1, Integer.Parse(timePieces(0)), 0)
                                        Dim minute = If(timePieces.Length >= 2, Integer.Parse(timePieces(1)), 0)

                                        ' შევქმნათ თარიღის ობიექტი ამ კომპონენტებით
                                        session.DateTime = New DateTime(year, month, day, hour, minute, 0)
                                    Else
                                        session.DateTime = DateTime.Now
                                    End If
                                Else
                                    session.DateTime = DateTime.Now
                                End If
                            Catch ex As Exception
                                Debug.WriteLine($"თარიღის პარსინგის შეცდომა: {dateTimeStr} - {ex.Message}")
                                session.DateTime = DateTime.Now
                            End Try
                        End If
                    Else
                        session.DateTime = DateTime.Now
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"თარიღის ძირითადი პარსინგის შეცდომა: {ex.Message}")
                    session.DateTime = DateTime.Now
                End Try

                ' დანარჩენი მონაცემების დაყენება - ყველგან Try-Catch დამატებით
                Try
                    session.Duration = If(rowData.Count > 5 AndAlso Not String.IsNullOrEmpty(rowData(5).ToString()),
                                Integer.Parse(rowData(5).ToString().Trim()), 60)
                Catch ex As Exception
                    Debug.WriteLine($"ხანგრძლივობის პარსინგის შეცდომა: {rowData(5)} - {ex.Message}")
                    session.Duration = 60 ' ნაგულისხმევი მნიშვნელობა
                End Try

                Try
                    session.IsGroup = If(rowData.Count > 6 AndAlso Not String.IsNullOrEmpty(rowData(6).ToString()),
                               Boolean.Parse(rowData(6).ToString().Trim()), False)
                Catch ex As Exception
                    Debug.WriteLine($"IsGroup პარსინგის შეცდომა: {rowData(6)} - {ex.Message}")
                    session.IsGroup = False ' ნაგულისხმევი მნიშვნელობა
                End Try

                Try
                    session.TherapistName = If(rowData.Count > 7, rowData(7).ToString().Trim(), "")
                Catch ex As Exception
                    session.TherapistName = ""
                End Try

                Try
                    session.TherapyType = If(rowData.Count > 8, rowData(8).ToString().Trim(), "")
                Catch ex As Exception
                    session.TherapyType = ""
                End Try

                Try
                    session.Space = If(rowData.Count > 9, rowData(9).ToString().Trim(), "")
                Catch ex As Exception
                    session.Space = ""
                End Try

                Try
                    session.Price = If(rowData.Count > 10 AndAlso Not String.IsNullOrEmpty(rowData(10).ToString()),
                            Decimal.Parse(rowData(10).ToString().Trim().Replace(",", ".")), 0)
                Catch ex As Exception
                    Debug.WriteLine($"ფასის პარსინგის შეცდომა: {rowData(10)} - {ex.Message}")
                    session.Price = 0 ' ნაგულისხმევი მნიშვნელობა
                End Try

                Try
                    session.Status = If(rowData.Count > 11, rowData(11).ToString().Trim(), "დაგეგმილი")
                Catch ex As Exception
                    session.Status = "დაგეგმილი" ' ნაგულისხმევი მნიშვნელობა
                End Try

                Try
                    session.Funding = If(rowData.Count > 12, rowData(12).ToString().Trim(), "კერძო")
                Catch ex As Exception
                    session.Funding = "კერძო" ' ნაგულისხმევი მნიშვნელობა
                End Try

                ' კომენტარი თუ არის და სესიებს აქვს საკმარისი სვეტები
                Try
                    session.Comments = If(rowData.Count > 13, rowData(13).ToString().Trim(), "")
                Catch ex As Exception
                    session.Comments = ""
                End Try

                Return session
            Catch ex As Exception
                Debug.WriteLine($"SessionModel.FromSheetRow: შეცდომა - {ex.Message}")
                Throw New Exception($"შეცდომა SessionModel-ის შექმნისას: {ex.Message}", ex)
            End Try
        End Function
        ''' <summary>
        ''' გარდაქმნის SessionModel-ს მწკრივად Google Sheets-ისთვის
        ''' </summary>
        ''' <returns>ობიექტების სია, რომელიც შეიძლება გადაეცეს Sheets API-ს</returns>
        Public Function ToSheetRow() As List(Of Object)
            Dim rowData As New List(Of Object)

            rowData.Add(Id)
            rowData.Add("") ' ცარიელი ადგილი რედაქტირების თარიღისთვის
            rowData.Add(BeneficiaryName)
            rowData.Add(BeneficiarySurname)
            rowData.Add(DateTime.ToString("dd.MM.yyyy HH:mm"))
            rowData.Add(Duration)
            rowData.Add(IsGroup)
            rowData.Add(TherapistName)
            rowData.Add(TherapyType)
            rowData.Add(Space)
            rowData.Add(Price)
            rowData.Add(Status)
            rowData.Add(Funding)
            rowData.Add(Comments)

            Return rowData
        End Function
        '' <summary>
        ''' არის თუ არა სესია ვადაგადაცილებული
        ''' </summary>
        Public ReadOnly Property IsOverdue As Boolean
            Get
                Try
                    ' მიმდინარე თარიღი - მხოლოდ თარიღის კომპონენტი (დრო 00:00)
                    Dim currentDate = DateTime.Today

                    ' სესიის თარიღი (მხოლოდ თარიღის ნაწილი, დროის გარეშე)
                    Dim sessionDate = Me.DateTime.Date

                    ' სტატუსი ზუსტად უნდა იყოს "დაგეგმილი"
                    Dim normalizedStatus = Me.Status.Trim().ToLower()
                    Dim isPlanned = (normalizedStatus = "დაგეგმილი")

                    ' შევადაროთ თარიღები - სესია უნდა იყოს წარსულში
                    Dim isPastDue = sessionDate < currentDate

                    ' ვადაგადაცილებულია, თუ:
                    ' 1. სტატუსი არის "დაგეგმილი" და
                    ' 2. სესიის თარიღი უკვე გასულია
                    Dim result = isPlanned AndAlso isPastDue

                    Debug.WriteLine($"SessionModel.IsOverdue [{Id}]: სტატუსი='{Me.Status}', isPlanned={isPlanned}")
                    Debug.WriteLine($"SessionModel.IsOverdue [{Id}]: სესიის თარიღი={sessionDate:dd.MM.yyyy}, დღეს={currentDate:dd.MM.yyyy}, isPastDue={isPastDue}")
                    Debug.WriteLine($"SessionModel.IsOverdue [{Id}]: შედეგი = {result}")

                    Return result
                Catch ex As Exception
                    Debug.WriteLine($"შეცდომა IsOverdue შემოწმებისას: {ex.Message}")
                    Return False
                End Try
            End Get
        End Property
    End Class
End Namespace