' ===========================================
' 📄 Models/HomeViewModel.vb
' -------------------------------------------
' საწყისი გვერდის ViewModel - ინახავს საწყისი გვერდის მონაცემებს
' ===========================================
Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Globalization

Namespace Scheduler_v8_8a.Models

    ''' <summary>
    ''' HomeViewModel ახორციელებს INotifyPropertyChanged და ინახავს საწყისი გვერდის მდგომარეობას
    ''' Properties: Greeting, CurrentTime, PendingSessionsCount, PendingSessions, UpcomingBirthdays, ActiveTasks
    ''' </summary>
    Public Class HomeViewModel
        Implements INotifyPropertyChanged

        ' ძირითადი მონაცემების ველები
        Private _greeting As String = String.Empty
        Private _userName As String = String.Empty
        Private _currentTime As DateTime = DateTime.Now
        Private _pendingSessionsCount As Integer = 0
        Private _pendingSessions As ObservableCollection(Of SessionModel)
        Private _upcomingBirthdays As ObservableCollection(Of BirthdayModel)
        Private _activeTasks As ObservableCollection(Of TaskModel)

        ''' <summary>მისალმების ტექსტი</summary>
        Public Property Greeting As String
            Get
                Return _greeting
            End Get
            Set(value As String)
                If _greeting <> value Then
                    _greeting = value
                    OnPropertyChanged(NameOf(Greeting))
                End If
            End Set
        End Property

        ''' <summary>მომხმარებლის სახელი მისალმებისთვის</summary>
        Public Property UserName As String
            Get
                Return _userName
            End Get
            Set(value As String)
                If _userName <> value Then
                    _userName = value
                    OnPropertyChanged(NameOf(UserName))
                End If
            End Set
        End Property

        ''' <summary>მიმდინარე დრო</summary>
        Public Property CurrentTime As DateTime
            Get
                Return _currentTime
            End Get
            Set(value As DateTime)
                If _currentTime <> value Then
                    _currentTime = value
                    ' დროის ყველა ფორმატის განახლება
                    OnPropertyChanged(NameOf(CurrentTime))
                    OnPropertyChanged(NameOf(FormattedTime))
                    OnPropertyChanged(NameOf(FormattedDate))
                    OnPropertyChanged(NameOf(WeekDayName))
                End If
            End Set
        End Property

        ''' <summary>მიმდინარე დრო ფორმატირებული (HH:mm:ss)</summary>
        Public ReadOnly Property FormattedTime As String
            Get
                Return CurrentTime.ToString("HH:mm:ss")
            End Get
        End Property

        ''' <summary>მიმდინარე თარიღი ფორმატირებული (დღე თვე წელი)</summary>
        Public ReadOnly Property FormattedDate As String
            Get
                ' ქართული ფორმატით თარიღის ფორმირება
                Return String.Format("{0} {1} {2} წელი",
                        CurrentTime.Day,
                        CurrentTime.ToString("MMMM", New CultureInfo("ka")),
                        CurrentTime.Year)
            End Get
        End Property

        ''' <summary>კვირის დღის დასახელება ქართულად</summary>
        Public ReadOnly Property WeekDayName As String
            Get
                Return CurrentTime.ToString("dddd", New CultureInfo("ka"))
            End Get
        End Property

        ''' <summary>მისალმების ტექსტის ფორმირება დროის მიხედვით</summary>
        Public Function GetGreetingByTime() As String
            Dim hour = DateTime.Now.Hour

            ' საათის მიხედვით შესაბამისი მისალმების შერჩევა
            If hour >= 5 AndAlso hour < 12 Then
                Return "დილა მშვიდობისა"
            ElseIf hour >= 12 AndAlso hour < 17 Then
                Return "შუადღე მშვიდობისა"
            ElseIf hour >= 17 AndAlso hour < 21 Then
                Return "საღამო მშვიდობისა"
            Else
                Return "ღამე მშვიდობისა"
            End If
        End Function

        ''' <summary>მოლოდინში არსებული სესიების რაოდენობა</summary>
        Public Property PendingSessionsCount As Integer
            Get
                Return _pendingSessionsCount
            End Get
            Set(value As Integer)
                If _pendingSessionsCount <> value Then
                    _pendingSessionsCount = value
                    OnPropertyChanged(NameOf(PendingSessionsCount))
                End If
            End Set
        End Property

        ''' <summary>მოლოდინში არსებული სესიების სია</summary>
        Public Property PendingSessions As ObservableCollection(Of SessionModel)
            Get
                If _pendingSessions Is Nothing Then
                    _pendingSessions = New ObservableCollection(Of SessionModel)()
                End If
                Return _pendingSessions
            End Get
            Set(value As ObservableCollection(Of SessionModel))
                _pendingSessions = value
                OnPropertyChanged(NameOf(PendingSessions))
            End Set
        End Property

        ''' <summary>მომავალი დაბადების დღეების სია</summary>
        Public Property UpcomingBirthdays As ObservableCollection(Of BirthdayModel)
            Get
                If _upcomingBirthdays Is Nothing Then
                    _upcomingBirthdays = New ObservableCollection(Of BirthdayModel)()
                End If
                Return _upcomingBirthdays
            End Get
            Set(value As ObservableCollection(Of BirthdayModel))
                _upcomingBirthdays = value
                OnPropertyChanged(NameOf(UpcomingBirthdays))
            End Set
        End Property

        ''' <summary>აქტიური დავალებების სია</summary>
        Public Property ActiveTasks As ObservableCollection(Of TaskModel)
            Get
                If _activeTasks Is Nothing Then
                    _activeTasks = New ObservableCollection(Of TaskModel)()
                End If
                Return _activeTasks
            End Get
            Set(value As ObservableCollection(Of TaskModel))
                _activeTasks = value
                OnPropertyChanged(NameOf(ActiveTasks))
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
        ''' კონსტრუქტორი: ინიციალიზება ახდენს ყველა კოლექციის და დროის
        ''' </summary>
        Public Sub New()
            ' კოლექციების ინიციალიზაცია
            _pendingSessions = New ObservableCollection(Of SessionModel)()
            _upcomingBirthdays = New ObservableCollection(Of BirthdayModel)()
            _activeTasks = New ObservableCollection(Of TaskModel)()

            ' დროის ინიციალიზაცია
            _currentTime = DateTime.Now

            ' საწყისი მისალმების დაყენება
            _greeting = GetGreetingByTime()
        End Sub

        ''' <summary>
        ''' ანახლებს მიმდინარე დროს
        ''' გამოსაყენებელია Timer-დან
        ''' </summary>
        Public Sub UpdateTime()
            CurrentTime = DateTime.Now
        End Sub

        ''' <summary>
        ''' მონაცემების განახლება ყველა წყაროდან
        ''' </summary>
        Public Sub RefreshData()
            ' TODO: შეასრულე ყველა მონაცემის განახლება:
            ' - მიმდინარე სესიები
            ' - დაბადების დღეები
            ' - აქტიური დავალებები

            ' დაზუსტება: ეს მეთოდი გამოძახებული იქნება სერვისებიდან ან Form1-დან
        End Sub
    End Class
End Namespace