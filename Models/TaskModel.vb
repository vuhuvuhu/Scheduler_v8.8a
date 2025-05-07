' ===========================================
' 📄 Models/TaskModel.vb
' -------------------------------------------
' დავალების/ამოცანის მოდელი
' ===========================================
Imports System.ComponentModel

Namespace Scheduler_v8_8a.Models

    ''' <summary>
    ''' TaskModel - ინახავს ინფორმაციას დავალებებისა და ამოცანების შესახებ
    ''' </summary>
    Public Class TaskModel
        Implements INotifyPropertyChanged

        ' ძირითადი მონაცემების ველები
        Private _id As Integer
        Private _title As String
        Private _description As String
        Private _assignedTo As String
        Private _assignedBy As String
        Private _creationDate As DateTime
        Private _dueDate As DateTime
        Private _priority As TaskPriority
        Private _status As TaskStatus
        Private _percentComplete As Integer
        Private _category As String
        Private _tags As List(Of String)

        ''' <summary>დავალების პრიორიტეტის ჩამონათვალი</summary>
        Public Enum TaskPriority
            Low = 0
            Medium = 1
            High = 2
            Critical = 3
        End Enum

        ''' <summary>დავალების სტატუსის ჩამონათვალი</summary>
        Public Enum TaskStatus
            NotStarted = 0
            InProgress = 1
            OnHold = 2
            Completed = 3
            Cancelled = 4
        End Enum

        ''' <summary>დავალების ID</summary>
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

        ''' <summary>დავალების სათაური</summary>
        Public Property Title As String
            Get
                Return _title
            End Get
            Set(value As String)
                If _title <> value Then
                    _title = value
                    OnPropertyChanged(NameOf(Title))
                End If
            End Set
        End Property

        ''' <summary>დავალების აღწერა</summary>
        Public Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                If _description <> value Then
                    _description = value
                    OnPropertyChanged(NameOf(Description))
                End If
            End Set
        End Property

        ''' <summary>ვისთვისაც არის დავალება განკუთვნილი</summary>
        Public Property AssignedTo As String
            Get
                Return _assignedTo
            End Get
            Set(value As String)
                If _assignedTo <> value Then
                    _assignedTo = value
                    OnPropertyChanged(NameOf(AssignedTo))
                End If
            End Set
        End Property

        ''' <summary>ვინ გასცა დავალება</summary>
        Public Property AssignedBy As String
            Get
                Return _assignedBy
            End Get
            Set(value As String)
                If _assignedBy <> value Then
                    _assignedBy = value
                    OnPropertyChanged(NameOf(AssignedBy))
                End If
            End Set
        End Property

        ''' <summary>დავალების შექმნის თარიღი</summary>
        Public Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                If _creationDate <> value Then
                    _creationDate = value
                    OnPropertyChanged(NameOf(CreationDate))
                    OnPropertyChanged(NameOf(FormattedCreationDate))
                End If
            End Set
        End Property

        ''' <summary>ფორმატირებული შექმნის თარიღი</summary>
        Public ReadOnly Property FormattedCreationDate As String
            Get
                Return _creationDate.ToString("dd.MM.yyyy HH:mm")
            End Get
        End Property

        ''' <summary>დავალების შესრულების ვადა</summary>
        Public Property DueDate As DateTime
            Get
                Return _dueDate
            End Get
            Set(value As DateTime)
                If _dueDate <> value Then
                    _dueDate = value
                    OnPropertyChanged(NameOf(DueDate))
                    OnPropertyChanged(NameOf(FormattedDueDate))
                    OnPropertyChanged(NameOf(DaysUntilDue))
                    OnPropertyChanged(NameOf(IsOverdue))
                End If
            End Set
        End Property

        ''' <summary>ფორმატირებული შესრულების ვადა</summary>
        Public ReadOnly Property FormattedDueDate As String
            Get
                Return _dueDate.ToString("dd.MM.yyyy HH:mm")
            End Get
        End Property

        ''' <summary>დღეების რაოდენობა შესრულების ვადამდე</summary>
        Public ReadOnly Property DaysUntilDue As Integer
            Get
                Dim today = DateTime.Today
                Return (_dueDate.Date - today).Days
            End Get
        End Property

        ''' <summary>არის თუ არა ვადაგადაცილებული</summary>
        Public ReadOnly Property IsOverdue As Boolean
            Get
                Return DateTime.Now > _dueDate And _status <> TaskStatus.Completed And _status <> TaskStatus.Cancelled
            End Get
        End Property

        ''' <summary>დავალების პრიორიტეტი</summary>
        Public Property Priority As TaskPriority
            Get
                Return _priority
            End Get
            Set(value As TaskPriority)
                If _priority <> value Then
                    _priority = value
                    OnPropertyChanged(NameOf(Priority))
                    OnPropertyChanged(NameOf(PriorityText))
                End If
            End Set
        End Property

        ''' <summary>პრიორიტეტის ტექსტური წარმოდგენა</summary>
        Public ReadOnly Property PriorityText As String
            Get
                Select Case _priority
                    Case TaskPriority.Low
                        Return "დაბალი"
                    Case TaskPriority.Medium
                        Return "საშუალო"
                    Case TaskPriority.High
                        Return "მაღალი"
                    Case TaskPriority.Critical
                        Return "კრიტიკული"
                    Case Else
                        Return "უცნობი"
                End Select
            End Get
        End Property

        ''' <summary>დავალების სტატუსი</summary>
        Public Property Status As TaskStatus
            Get
                Return _status
            End Get
            Set(value As TaskStatus)
                If _status <> value Then
                    _status = value
                    OnPropertyChanged(NameOf(Status))
                    OnPropertyChanged(NameOf(StatusText))
                    OnPropertyChanged(NameOf(IsOverdue)) ' გადაიანგარიშდება
                End If
            End Set
        End Property

        ''' <summary>სტატუსის ტექსტური წარმოდგენა</summary>
        Public ReadOnly Property StatusText As String
            Get
                Select Case _status
                    Case TaskStatus.NotStarted
                        Return "არ დაწყებულა"
                    Case TaskStatus.InProgress
                        Return "მიმდინარე"
                    Case TaskStatus.OnHold
                        Return "შეჩერებული"
                    Case TaskStatus.Completed
                        Return "დასრულებული"
                    Case TaskStatus.Cancelled
                        Return "გაუქმებული"
                    Case Else
                        Return "უცნობი"
                End Select
            End Get
        End Property

        ''' <summary>დავალების შესრულების პროცენტი</summary>
        Public Property PercentComplete As Integer
            Get
                Return _percentComplete
            End Get
            Set(value As Integer)
                ' შევამოწმოთ ლეგიტიმურია თუ არა პროცენტის მნიშვნელობა
                Dim validValue = Math.Max(0, Math.Min(100, value))
                If _percentComplete <> validValue Then
                    _percentComplete = validValue
                    OnPropertyChanged(NameOf(PercentComplete))
                End If
            End Set
        End Property

        ''' <summary>დავალების კატეგორია</summary>
        Public Property Category As String
            Get
                Return _category
            End Get
            Set(value As String)
                If _category <> value Then
                    _category = value
                    OnPropertyChanged(NameOf(Category))
                End If
            End Set
        End Property

        ''' <summary>დავალების ტეგები</summary>
        Public Property Tags As List(Of String)
            Get
                If _tags Is Nothing Then
                    _tags = New List(Of String)()
                End If
                Return _tags
            End Get
            Set(value As List(Of String))
                _tags = value
                OnPropertyChanged(NameOf(Tags))
                OnPropertyChanged(NameOf(TagsText))
            End Set
        End Property

        ''' <summary>ტეგების ტექსტური წარმოდგენა მძიმეებით გამოყოფილი</summary>
        Public ReadOnly Property TagsText As String
            Get
                If _tags Is Nothing OrElse _tags.Count = 0 Then
                    Return String.Empty
                End If
                Return String.Join(", ", _tags)
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
        ''' ცარიელი კონსტრუქტორი
        ''' </summary>
        Public Sub New()
            ' საწყისი მნიშვნელობების ინიციალიზაცია
            _title = String.Empty
            _description = String.Empty
            _assignedTo = String.Empty
            _assignedBy = String.Empty
            _creationDate = DateTime.Now
            _dueDate = DateTime.Now.AddDays(7) ' ნაგულისხმები 7 დღე
            _priority = TaskPriority.Medium
            _status = TaskStatus.NotStarted
            _percentComplete = 0
            _category = String.Empty
            _tags = New List(Of String)()
        End Sub

        ''' <summary>
        ''' კონსტრუქტორი პარამეტრებით
        ''' </summary>
        Public Sub New(id As Integer, title As String, description As String,
                      assignedTo As String, assignedBy As String,
                      dueDate As DateTime, priority As TaskPriority)
            Me.New() ' გამოძახება ცარიელი კონსტრუქტორის
            Me.Id = id
            Me.Title = title
            Me.Description = description
            Me.AssignedTo = assignedTo
            Me.AssignedBy = assignedBy
            Me.DueDate = dueDate
            Me.Priority = priority
        End Sub

        ''' <summary>
        ''' შექმნის TaskModel-ს Google Sheets-დან მოცემული მწკრივიდან
        ''' </summary>
        ''' <param name="rowData">მწკრივი Google Sheets-დან</param>
        ''' <returns>TaskModel შევსებული მონაცემებით</returns>
        Public Shared Function FromSheetRow(rowData As IList(Of Object)) As TaskModel
            ' შემოწმება მონაცემების რაოდენობის
            If rowData Is Nothing OrElse rowData.Count < 8 Then
                Throw New ArgumentException("არასაკმარისი მონაცემები TaskModel-ისთვის")
            End If

            Try
                Dim task As New TaskModel()

                ' ID და ძირითადი ინფორმაციის დაყენება
                task.Id = Integer.Parse(rowData(0).ToString())
                task.Title = rowData(1).ToString()
                task.Description = rowData(2).ToString()
                task.AssignedTo = rowData(3).ToString()
                task.AssignedBy = rowData(4).ToString()

                ' თარიღების პარსინგი
                Dim creationDateStr = rowData(5).ToString()
                Dim creationDate As DateTime
                If DateTime.TryParseExact(creationDateStr, "dd.MM.yyyy HH:mm",
                                       Globalization.CultureInfo.InvariantCulture,
                                       Globalization.DateTimeStyles.None, creationDate) Then
                    task.CreationDate = creationDate
                End If

                Dim dueDateStr = rowData(6).ToString()
                Dim dueDate As DateTime
                If DateTime.TryParseExact(dueDateStr, "dd.MM.yyyy HH:mm",
                                       Globalization.CultureInfo.InvariantCulture,
                                       Globalization.DateTimeStyles.None, dueDate) Then
                    task.DueDate = dueDate
                End If

                ' პრიორიტეტის, სტატუსის და პროცენტის დაყენება
                task.Priority = CType(Integer.Parse(rowData(7).ToString()), TaskPriority)

                If rowData.Count > 8 Then
                    task.Status = CType(Integer.Parse(rowData(8).ToString()), TaskStatus)
                End If

                If rowData.Count > 9 Then
                    task.PercentComplete = Integer.Parse(rowData(9).ToString())
                End If

                ' კატეგორიის და ტეგების დაყენება
                If rowData.Count > 10 Then
                    task.Category = rowData(10).ToString()
                End If

                If rowData.Count > 11 AndAlso Not String.IsNullOrEmpty(rowData(11).ToString()) Then
                    task.Tags = rowData(11).ToString().Split(","c).Select(Function(t) t.Trim()).ToList()
                End If

                Return task
            Catch ex As Exception
                Throw New Exception("შეცდომა TaskModel-ის შექმნისას: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' გარდაქმნის TaskModel-ს მწკრივად Google Sheets-ისთვის
        ''' </summary>
        ''' <returns>ობიექტების სია, რომელიც შეიძლება გადაეცეს Sheets API-ს</returns>
        Public Function ToSheetRow() As List(Of Object)
            Dim rowData As New List(Of Object)

            rowData.Add(Id)
            rowData.Add(Title)
            rowData.Add(Description)
            rowData.Add(AssignedTo)
            rowData.Add(AssignedBy)
            rowData.Add(CreationDate.ToString("dd.MM.yyyy HH:mm"))
            rowData.Add(DueDate.ToString("dd.MM.yyyy HH:mm"))
            rowData.Add(CInt(Priority))
            rowData.Add(CInt(Status))
            rowData.Add(PercentComplete)
            rowData.Add(Category)
            rowData.Add(TagsText)

            Return rowData
        End Function
    End Class
End Namespace