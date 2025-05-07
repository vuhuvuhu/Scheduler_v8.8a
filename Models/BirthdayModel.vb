' ===========================================
' 📄 Models/BirthdayModel.vb
' -------------------------------------------
' დაბადების დღის ინფორმაციის შემცველი მოდელი
' ===========================================
Imports System.ComponentModel

Namespace Scheduler_v8_8a.Models

    ''' <summary>
    ''' BirthdayModel - ინახავს ინფორმაციას დაბადების დღეების შესახებ
    ''' </summary>
    Public Class BirthdayModel
        Implements INotifyPropertyChanged

        ' მონაცემების ველები
        Private _id As Integer
        Private _personName As String
        Private _personSurname As String
        Private _birthDate As DateTime
        Private _age As Integer
        Private _email As String
        Private _phone As String
        Private _notes As String

        ''' <summary>პიროვნების ID</summary>
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

        ''' <summary>პიროვნების სახელი</summary>
        Public Property PersonName As String
            Get
                Return _personName
            End Get
            Set(value As String)
                If _personName <> value Then
                    _personName = value
                    OnPropertyChanged(NameOf(PersonName))
                    OnPropertyChanged(NameOf(FullName)) ' ასევე განაახლებს FullName-ს
                End If
            End Set
        End Property

        ''' <summary>პიროვნების გვარი</summary>
        Public Property PersonSurname As String
            Get
                Return _personSurname
            End Get
            Set(value As String)
                If _personSurname <> value Then
                    _personSurname = value
                    OnPropertyChanged(NameOf(PersonSurname))
                    OnPropertyChanged(NameOf(FullName)) ' ასევე განაახლებს FullName-ს
                End If
            End Set
        End Property

        ''' <summary>სრული სახელი (სახელი + გვარი)</summary>
        Public ReadOnly Property FullName As String
            Get
                Return $"{_personName} {_personSurname}"
            End Get
        End Property

        ''' <summary>დაბადების თარიღი</summary>
        Public Property BirthDate As DateTime
            Get
                Return _birthDate
            End Get
            Set(value As DateTime)
                If _birthDate <> value Then
                    _birthDate = value
                    OnPropertyChanged(NameOf(BirthDate))
                    OnPropertyChanged(NameOf(FormattedBirthDate))
                    OnPropertyChanged(NameOf(Age))
                    OnPropertyChanged(NameOf(DaysUntilBirthday))
                End If
            End Set
        End Property

        ''' <summary>ფორმატირებული დაბადების თარიღი</summary>
        Public ReadOnly Property FormattedBirthDate As String
            Get
                Return _birthDate.ToString("dd.MM.yyyy")
            End Get
        End Property

        ''' <summary>ასაკი (წლებში)</summary>
        Public ReadOnly Property Age As Integer
            Get
                Dim today = DateTime.Today
                Dim currentAge = today.Year - _birthDate.Year ' შეცვლილია ცვლადის სახელი age -> currentAge
                ' შეამოწმე ამ წელს უკვე ქონდა თუ არა დაბადების დღე
                If _birthDate.Date > today.AddYears(-currentAge) Then
                    currentAge -= 1
                End If
                Return currentAge
            End Get
        End Property

        ''' <summary>რამდენი დღე დარჩა დაბადების დღემდე</summary>
        Public ReadOnly Property DaysUntilBirthday As Integer
            Get
                Dim today = DateTime.Today
                Dim nextBirthday = New DateTime(today.Year, _birthDate.Month, _birthDate.Day)

                ' თუ დაბადების დღე უკვე ჩატარდა წელს, ვიყენებთ მომავალი წლის თარიღს
                If nextBirthday < today Then
                    nextBirthday = nextBirthday.AddYears(1)
                End If

                Return (nextBirthday - today).Days
            End Get
        End Property

        ''' <summary>პიროვნების ელ.ფოსტა</summary>
        Public Property Email As String
            Get
                Return _email
            End Get
            Set(value As String)
                If _email <> value Then
                    _email = value
                    OnPropertyChanged(NameOf(Email))
                End If
            End Set
        End Property

        ''' <summary>პიროვნების ტელეფონის ნომერი</summary>
        Public Property Phone As String
            Get
                Return _phone
            End Get
            Set(value As String)
                If _phone <> value Then
                    _phone = value
                    OnPropertyChanged(NameOf(Phone))
                End If
            End Set
        End Property

        ''' <summary>დამატებითი შენიშვნები</summary>
        Public Property Notes As String
            Get
                Return _notes
            End Get
            Set(value As String)
                If _notes <> value Then
                    _notes = value
                    OnPropertyChanged(NameOf(Notes))
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
        ''' ცარიელი კონსტრუქტორი
        ''' </summary>
        Public Sub New()
            ' საწყისი მნიშვნელობები
            _personName = String.Empty
            _personSurname = String.Empty
            _birthDate = DateTime.Today
            _email = String.Empty
            _phone = String.Empty
            _notes = String.Empty
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი პარამეტრებით
        ''' </summary>
        Public Sub New(id As Integer, personName As String, personSurname As String,
                      birthDateValue As DateTime, email As String, phone As String)
            Me.Id = id
            Me.PersonName = personName
            Me.PersonSurname = personSurname
            Me.BirthDate = birthDateValue ' შეცვლილია პარამეტრის სახელი
            Me.Email = email
            Me.Phone = phone
            Me.Notes = String.Empty
        End Sub

        ''' <summary>
        ''' შექმნის BirthdayModel-ს Google Sheets-დან მოცემული მწკრივიდან
        ''' </summary>
        ''' <param name="rowData">მწკრივი Google Sheets-დან</param>
        ''' <returns>BirthdayModel შევსებული მონაცემებით</returns>
        Public Shared Function CreateFromSheetRow(rowData As IList(Of Object)) As BirthdayModel ' შეცვლილია ფუნქციის სახელი
            ' შემოწმება მონაცემების რაოდენობის
            If rowData Is Nothing OrElse rowData.Count < 6 Then
                Throw New ArgumentException("არასაკმარისი მონაცემები BirthdayModel-ისთვის")
            End If

            Try
                Dim birthday As New BirthdayModel()

                ' ID და პიროვნების ინფორმაციის დაყენება
                birthday.Id = Integer.Parse(rowData(0).ToString())
                birthday.PersonName = rowData(1).ToString()
                birthday.PersonSurname = rowData(2).ToString()

                ' დაბადების თარიღის პარსინგი
                Dim birthDateStr = rowData(3).ToString()
                Dim parsedBirthDate As DateTime
                If DateTime.TryParseExact(birthDateStr, "dd.MM.yyyy",
                                        Globalization.CultureInfo.InvariantCulture,
                                        Globalization.DateTimeStyles.None, parsedBirthDate) Then
                    birthday.BirthDate = parsedBirthDate
                End If

                ' საკონტაქტო ინფორმაცია
                birthday.Email = rowData(4).ToString()
                birthday.Phone = rowData(5).ToString()

                ' შენიშვნები თუ არის
                If rowData.Count > 6 Then
                    birthday.Notes = rowData(6).ToString()
                End If

                Return birthday
            Catch ex As Exception
                Throw New Exception("შეცდომა BirthdayModel-ის შექმნისას: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' გარდაქმნის BirthdayModel-ს მწკრივად Google Sheets-ისთვის
        ''' </summary>
        ''' <returns>ობიექტების სია, რომელიც შეიძლება გადაეცეს Sheets API-ს</returns>
        Public Function ToSheetRow() As List(Of Object)
            Dim rowData As New List(Of Object)

            rowData.Add(Id)
            rowData.Add(PersonName)
            rowData.Add(PersonSurname)
            rowData.Add(BirthDate.ToString("dd.MM.yyyy"))
            rowData.Add(Email)
            rowData.Add(Phone)
            rowData.Add(Notes)

            Return rowData
        End Function
    End Class
End Namespace