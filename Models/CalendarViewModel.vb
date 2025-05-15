' ===========================================
' 📄 Models/CalendarViewModel.vb
' -------------------------------------------
' კალენდრის ViewModel - მართავს კალენდრის მონაცემებს
' ===========================================
Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Public Class CalendarViewModel
    Implements INotifyPropertyChanged

    ' მონაცემთა ველები
    Private _selectedMonth As DateTime    ' არჩეული თვე კალენდარზე
    Private _sessions As New ObservableCollection(Of SessionModel)    ' ყველა სესია
    Private _selectedDate As DateTime     ' არჩეული თარიღი
    Private _selectedDaySessions As New ObservableCollection(Of SessionModel)  ' არჩეული დღის სესიები
    Private _userEmail As String = String.Empty  ' მომხმარებლის მეილი

    ' INotifyPropertyChanged ივენთი
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' შერჩეული თვე კალენდარში
    ''' </summary>
    Public Property SelectedMonth As DateTime
        Get
            Return _selectedMonth
        End Get
        Set(value As DateTime)
            If _selectedMonth <> value Then
                _selectedMonth = value
                OnPropertyChanged(NameOf(SelectedMonth))
                OnPropertyChanged(NameOf(MonthName))
                OnPropertyChanged(NameOf(Year))
            End If
        End Set
    End Property

    ''' <summary>
    ''' თვის სახელი (მაგ: "იანვარი")
    ''' </summary>
    Public ReadOnly Property MonthName As String
        Get
            Return _selectedMonth.ToString("MMMM", New Globalization.CultureInfo("ka-GE"))
        End Get
    End Property

    ''' <summary>
    ''' წელი (მაგ: "2023")
    ''' </summary>
    Public ReadOnly Property Year As Integer
        Get
            Return _selectedMonth.Year
        End Get
    End Property

    ''' <summary>
    ''' სესიების კოლექცია
    ''' </summary>
    Public Property Sessions As ObservableCollection(Of SessionModel)
        Get
            Return _sessions
        End Get
        Set(value As ObservableCollection(Of SessionModel))
            _sessions = value
            OnPropertyChanged(NameOf(Sessions))
        End Set
    End Property

    ''' <summary>
    ''' შერჩეული დღე კალენდარში
    ''' </summary>
    Public Property SelectedDate As DateTime
        Get
            Return _selectedDate
        End Get
        Set(value As DateTime)
            If _selectedDate <> value Then
                _selectedDate = value
                OnPropertyChanged(NameOf(SelectedDate))

                ' გავფილტროთ სესიები არჩეული დღისთვის
                LoadSelectedDaySessions()
            End If
        End Set
    End Property

    ''' <summary>
    ''' არჩეული დღის სესიები
    ''' </summary>
    Public Property SelectedDaySessions As ObservableCollection(Of SessionModel)
        Get
            Return _selectedDaySessions
        End Get
        Set(value As ObservableCollection(Of SessionModel))
            _selectedDaySessions = value
            OnPropertyChanged(NameOf(SelectedDaySessions))
        End Set
    End Property

    ''' <summary>
    ''' მომხმარებლის ელ.ფოსტა
    ''' </summary>
    Public Property UserEmail As String
        Get
            Return _userEmail
        End Get
        Set(value As String)
            If _userEmail <> value Then
                _userEmail = value
                OnPropertyChanged(NameOf(UserEmail))
            End If
        End Set
    End Property

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    Public Sub New()
        ' საწყისი მნიშვნელობები
        _selectedMonth = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        _selectedDate = DateTime.Today
        _sessions = New ObservableCollection(Of SessionModel)()
        _selectedDaySessions = New ObservableCollection(Of SessionModel)()
    End Sub

    ''' <summary>
    ''' PropertyChanged ივენთის გამოძახება
    ''' </summary>
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    ''' <summary>
    ''' წინა თვეზე გადასვლა
    ''' </summary>
    Public Sub PreviousMonth()
        SelectedMonth = _selectedMonth.AddMonths(-1)
    End Sub

    ''' <summary>
    ''' შემდეგ თვეზე გადასვლა
    ''' </summary>
    Public Sub NextMonth()
        SelectedMonth = _selectedMonth.AddMonths(1)
    End Sub

    ''' <summary>
    ''' მიმდინარე თვეზე დაბრუნება
    ''' </summary>
    Public Sub CurrentMonth()
        SelectedMonth = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
    End Sub

    ''' <summary>
    ''' სესიების ჩატვირთვა შერჩეული თვისთვის
    ''' </summary>
    Public Sub LoadMonthSessions(sessions As List(Of SessionModel))
        _sessions.Clear()
        If sessions IsNot Nothing Then
            For Each session In sessions
                _sessions.Add(session)
            Next
        End If

        ' განვაახლოთ შერჩეული დღის სესიებიც
        LoadSelectedDaySessions()
    End Sub

    ''' <summary>
    ''' შერჩეული დღის სესიების ჩატვირთვა
    ''' </summary>
    Private Sub LoadSelectedDaySessions()
        _selectedDaySessions.Clear()

        ' გავფილტროთ მხოლოდ შერჩეული დღის სესიები
        For Each session In _sessions
            If session.DateTime.Date = _selectedDate.Date Then
                _selectedDaySessions.Add(session)
            End If
        Next

        OnPropertyChanged(NameOf(SelectedDaySessions))
    End Sub
End Class