' 📄 UserControls/UC_TherapistReport.vb
' -------------------------------------------
' თერაპევტის რეპორტის UserControl - UC_BeneficiaryReport.vb-ის საფუძველზე გადაწერილი
' მარტივი და გამართული არქიტექტურა
' ძირითადი ფუნქცია: თერაპევტის არჩევა და მისი ჩატარებული სესიების ჩვენება
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_TherapistReport
    Inherits UserControl

#Region "ველები და თვისებები"

    ' სერვისები - UC_BeneficiaryReport-ის ანალოგია
    Private dataService As IDataService = Nothing
    Private dataProcessor As ScheduleDataProcessor = Nothing
    Private uiManager As ScheduleUIManager = Nothing
    Private filterManager As ScheduleFilterManager = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6

    ' 🔧 ციკლური ივენთების თავიდან აცილება
    Private isLoadingData As Boolean = False
    Private isUpdatingTherapist As Boolean = False
    Private isUpdatingFilters As Boolean = False

    ' თერაპევტის სპეციფიკური ველები
    Private currentTherapistData As List(Of SessionModel) = Nothing

#End Region

#Region "საჯარო მეთოდები"

    ''' <summary>
    ''' მონაცემთა სერვისის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    Public Sub SetDataService(service As IDataService)
        Try
            Debug.WriteLine("UC_TherapistReport: SetDataService")

            dataService = service

            ' სერვისების ინიციალიზაცია
            If dataService IsNot Nothing Then
                InitializeServices()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: SetDataService შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის დაყენება
    ''' </summary>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub SetUserInfo(email As String, role As String)
        Try
            Debug.WriteLine($"UC_TherapistReport: SetUserInfo - email='{email}', role='{role}'")

            userEmail = email

            ' როლის პარსინგი
            Dim parsedRole As Integer
            If Integer.TryParse(role, parsedRole) Then
                userRoleID = parsedRole
            Else
                userRoleID = 6 ' ნაგულისხმევი
            End If

            Debug.WriteLine($"UC_TherapistReport: საბოლოო userRoleID: {userRoleID}, userEmail: '{userEmail}'")

            ' UI-ს განახლება ახალი როლისთვის
            If uiManager IsNot Nothing Then
                LoadTherapistSpecificData()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: SetUserInfo შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების განახლება
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("UC_TherapistReport: მონაცემების განახლება")

            ' ქეშის გასუფთავება
            If TypeOf dataService Is SheetDataService Then
                DirectCast(dataService, SheetDataService).InvalidateAllCache()
            End If

            dataProcessor?.ClearCache()

            ' ComboBox-ების განახლება
            RefreshTherapistComboBoxes()

            ' მონაცემების ხელახალი ჩატვირთვა
            LoadTherapistSpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: RefreshData შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "პირადი მეთოდები"

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_TherapistReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_TherapistReport: კონტროლის ჩატვირთვა")

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვსრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadTherapistSpecificData()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სერვისების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeServices()
        Try
            Debug.WriteLine("UC_TherapistReport: სერვისების ინიციალიზაცია")

            ' მონაცემების დამუშავების სერვისი
            dataProcessor = New ScheduleDataProcessor(dataService)

            ' UI მართვის სერვისი - მხოლოდ DataGridView
            uiManager = New ScheduleUIManager(DgvSessions, Nothing, Nothing, Nothing)
            ConfigureTherapistReportUI()

            ' ფილტრების მართვის სერვისი - მხოლოდ საჭირო კონტროლები
            filterManager = New ScheduleFilterManager(
                DtpDan, DtpMde,
                Nothing, Nothing, Nothing, Nothing, Nothing, Nothing ' ComboBox-ები ხელით ვმართავთ
            )

            ' სტატუსის CheckBox-ების ინიციალიზაცია
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ივენთების მიბმა
            BindEvents()

            ' თერაპევტის ComboBox-ების შევსება
            RefreshTherapistComboBoxes()

            Debug.WriteLine("UC_TherapistReport: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის რეპორტისთვის UI-ს კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureTherapistReportUI()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის UI კონფიგურაცია")

            ' DataGridView-ის კონფიგურაცია
            uiManager.ConfigureDataGridView()

            ' თერაპევტის რეპორტისთვის სპეციალური სვეტები
            ConfigureTherapistColumns()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ConfigureTherapistReportUI შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის სვეტების კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureTherapistColumns()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის სვეტების კონფიგურაცია")

            If DgvSessions Is Nothing Then Return

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' თერაპევტის რეპორტისთვის სპეციალური სვეტები
                .Add("N", "N")                        ' ID
                .Add("DateTime", "თარიღი")            ' თარიღი
                .Add("Duration", "ხანგძლ.")           ' ხანგძლივობა
                .Add("BeneficiaryName", "ბენეფიციარი") ' ბენეფიციარის სრული სახელი
                .Add("Funding", "დაფინანსება")        ' დაფინანსების სვეტი
                .Add("Price", "თანხა")               ' თანხა
                .Add("Status", "შესრულება")           ' შესრულების სვეტი

                ' რეპორტში ჩართვის CheckBox
                Dim includeColumn As New DataGridViewCheckBoxColumn()
                includeColumn.Name = "IncludeInReport"
                includeColumn.HeaderText = "რეპორტში"
                includeColumn.Width = 70
                includeColumn.TrueValue = True
                includeColumn.FalseValue = False
                includeColumn.IndeterminateValue = False
                includeColumn.ThreeState = False
                .Add(includeColumn)

                ' რედაქტირების ღილაკი
                Dim editBtn As New DataGridViewButtonColumn()
                editBtn.Name = "Edit"
                editBtn.HeaderText = ""
                editBtn.Text = "✎"
                editBtn.UseColumnTextForButtonValue = True
                editBtn.Width = 35
                .Add(editBtn)
            End With

            ' DataGridView-ის ზოგადი პარამეტრები CheckBox-ისთვის
            DgvSessions.EditMode = DataGridViewEditMode.EditOnEnter
            DgvSessions.AllowUserToAddRows = False
            DgvSessions.AllowUserToDeleteRows = False

            ' სვეტების სიგანეების დაყენება
            SetTherapistColumnWidths()

            Debug.WriteLine("UC_TherapistReport: თერაპევტის სვეტები კონფიგურირებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ConfigureTherapistColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის სვეტების სიგანეების დაყენება
    ''' </summary>
    Private Sub SetTherapistColumnWidths()
        Try
            With DgvSessions.Columns
                .Item("N").Width = 50
                .Item("DateTime").Width = 130
                .Item("Duration").Width = 60
                .Item("BeneficiaryName").Width = 200   ' ბენეფიციარის სრული სახელი
                .Item("Funding").Width = 100           ' დაფინანსების სვეტი
                .Item("Price").Width = 80
                .Item("Status").Width = 110            ' შესრულების სვეტი
                .Item("IncludeInReport").Width = 70
                .Item("Edit").Width = 35

                ' ფორმატირება
                .Item("Price").DefaultCellStyle.Format = "N2"
                .Item("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Item("N").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("IncludeInReport").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("Status").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("Funding").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: SetTherapistColumnWidths შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის CheckBox-ების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            Dim statusCheckBoxes As New List(Of CheckBox)

            ' CheckBox1-დან CheckBox7-მდე მოძიება
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBoxRecursive(Me, $"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    statusCheckBoxes.Add(checkBox)
                End If
            Next

            If statusCheckBoxes.Count > 0 Then
                filterManager.InitializeStatusCheckBoxes(statusCheckBoxes.ToArray())

                ' ნაგულისხმევი მონიშვნა - თერაპევტის რეპორტისთვის რელევანტური სტატუსები
                filterManager.SetAllStatusCheckBoxes(False)
                filterManager.SetStatusCheckBox("შესრულებული", True)
                filterManager.SetStatusCheckBox("გაცდენა არასაპატიო", True)
                filterManager.SetStatusCheckBox("აღდგენა", True)

                Debug.WriteLine($"UC_TherapistReport: {statusCheckBoxes.Count} სტატუსის CheckBox ინიციალიზებული")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: InitializeStatusCheckBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' CheckBox-ის რეკურსიული ძებნა კონტროლების იერარქიაში
    ''' </summary>
    ''' <param name="parent">მშობელი კონტროლი</param>
    ''' <param name="name">CheckBox-ის სახელი</param>
    ''' <returns>ნაპოვნი CheckBox ან Nothing</returns>
    Private Function FindCheckBoxRecursive(parent As Control, name As String) As CheckBox
        Try
            ' პირველ რიგში უშუალო შვილები
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is CheckBox AndAlso ctrl.Name = name Then
                    Debug.WriteLine($"UC_TherapistReport: ნაპოვნია CheckBox - {name}")
                    Return DirectCast(ctrl, CheckBox)
                End If
            Next

            ' თუ არ მოიძებნა, რეკურსიულად ძებნა ყველა შვილკონტროლში
            For Each ctrl As Control In parent.Controls
                If ctrl.HasChildren Then
                    Dim found = FindCheckBoxRecursive(ctrl, name)
                    If found IsNot Nothing Then
                        Return found
                    End If
                End If
            Next

            Return Nothing

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: FindCheckBoxRecursive შეცდომა {name}-სთვის: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' ComboBox-ის რეკურსიული ძებნა კონტროლების იერარქიაში
    ''' </summary>
    ''' <param name="parent">მშობელი კონტროლი</param>
    ''' <param name="name">ComboBox-ის სახელი</param>
    ''' <returns>ნაპოვნი ComboBox ან Nothing</returns>
    Private Function FindComboBoxRecursive(parent As Control, name As String) As ComboBox
        Try
            ' პირველ რიგში უშუალო შვილები
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is ComboBox AndAlso ctrl.Name = name Then
                    Debug.WriteLine($"UC_TherapistReport: ნაპოვნია ComboBox - {name}")
                    Return DirectCast(ctrl, ComboBox)
                End If
            Next

            ' თუ არ მოიძებნა, რეკურსიულად ძებნა ყველა შვილკონტროლში
            For Each ctrl As Control In parent.Controls
                If ctrl.HasChildren Then
                    Dim found = FindComboBoxRecursive(ctrl, name)
                    If found IsNot Nothing Then
                        Return found
                    End If
                End If
            Next

            Return Nothing

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: FindComboBoxRecursive შეცდომა {name}-სთვის: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 🔧 ივენთების გაუმჯობესებული მიბმა (ორმაგი მიბმის თავიდან აცილება)
    ''' </summary>
    Private Sub BindEvents()
        Try
            Debug.WriteLine("UC_TherapistReport: ივენთების მიბმა იწყება")

            ' 🔧 ძველი ივენთების წაშლა (თუ არსებობს)
            RemoveHandler DgvSessions.CellClick, AddressOf OnDataGridViewCellClick
            RemoveHandler DgvSessions.CellValueChanged, AddressOf OnCheckBoxChanged
            RemoveHandler DgvSessions.CurrentCellDirtyStateChanged, AddressOf OnCurrentCellDirtyStateChanged

            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged

            ' DataGridView-ის ივენთები - CheckBox-ის სწორი მართვისთვის
            AddHandler DgvSessions.CellClick, AddressOf OnDataGridViewCellClick
            AddHandler DgvSessions.CellValueChanged, AddressOf OnCheckBoxChanged
            AddHandler DgvSessions.CurrentCellDirtyStateChanged, AddressOf OnCurrentCellDirtyStateChanged

            ' თერაპევტის ComboBox ივენთი
            If CBPer IsNot Nothing Then
                RemoveHandler CBPer.SelectedIndexChanged, AddressOf OnTherapistChanged
                AddHandler CBPer.SelectedIndexChanged, AddressOf OnTherapistChanged
                Debug.WriteLine("UC_TherapistReport: თერაპევტის ComboBox (CBPer) ივენთი მიბმულია")
            Else
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBPer ComboBox არ მოიძებნა!")
            End If

            ' თერაპიისა და დაფინანსების ComboBox-ების ძებნა და ივენთების მიბმა
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            Dim cbFunding As ComboBox = FindComboBoxRecursive(Me, "CBDaf")

            If cbTherapyType IsNot Nothing Then
                RemoveHandler cbTherapyType.SelectedIndexChanged, AddressOf OnTherapyTypeChanged
                AddHandler cbTherapyType.SelectedIndexChanged, AddressOf OnTherapyTypeChanged
                Debug.WriteLine("UC_TherapistReport: თერაპიის ComboBox (CBTer) ნაპოვნია და ივენთი მიბმულია")
            Else
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBTer ComboBox არ მოიძებნა!")
            End If

            If cbFunding IsNot Nothing Then
                RemoveHandler cbFunding.SelectedIndexChanged, AddressOf OnFundingChanged
                AddHandler cbFunding.SelectedIndexChanged, AddressOf OnFundingChanged
                Debug.WriteLine("UC_TherapistReport: დაფინანსების ComboBox (CBDaf) ნაპოვნია და ივენთი მიბმულია")
            Else
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBDaf ComboBox არ მოიძებნა!")
            End If

            Debug.WriteLine("UC_TherapistReport: ყველა ივენთი წარმატებით მიბმულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: BindEvents შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 DataGridView-ის კონფიგურაციის გაუმჯობესება (ორმაგი კლიკების თავიდან აცილება)
    ''' </summary>
    Private Sub ConfigureDataGridViewBehavior()
        Try
            Debug.WriteLine("UC_TherapistReport: DataGridView ქცევის კონფიგურაცია")

            If DgvSessions Is Nothing Then Return

            With DgvSessions
                ' CheckBox-ების ქცევის გაუმჯობესება
                .EditMode = DataGridViewEditMode.EditOnEnter
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .ReadOnly = False ' CheckBox-ების რედაქტირებისთვის

                ' სელექციის რეჟიმი
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = True

                ' ორმაგი კლიკის გამორთვა
                .StandardTab = True
                .TabStop = True

                ' CheckBox სვეტის სპეციალური კონფიგურაცია
                If .Columns.Contains("IncludeInReport") Then
                    .Columns("IncludeInReport").ReadOnly = False
                    .Columns("IncludeInReport").SortMode = DataGridViewColumnSortMode.NotSortable
                End If

                ' რედაქტირების ღილაკის კონფიგურაცია
                If .Columns.Contains("Edit") Then
                    .Columns("Edit").ReadOnly = True
                    .Columns("Edit").SortMode = DataGridViewColumnSortMode.NotSortable
                End If
            End With

            Debug.WriteLine("UC_TherapistReport: DataGridView ქცევა კონფიგურირებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ConfigureDataGridViewBehavior შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 გაუმჯობესებული CurrentCellDirtyStateChanged ივენთი
    ''' </summary>
    Private Sub OnCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        Static lastDirtyTime As DateTime = DateTime.MinValue

        Try
            ' 🔧 ზედმეტი გამოძახებების თავიდან აცილება
            Dim currentTime As DateTime = DateTime.Now
            If currentTime.Subtract(lastDirtyTime).TotalMilliseconds < 100 Then
                Return
            End If
            lastDirtyTime = currentTime

            ' თუ CheckBox სვეტში ვართ და უჯრა "ღია" არის
            If DgvSessions.IsCurrentCellDirty AndAlso
               DgvSessions.CurrentCell IsNot Nothing AndAlso
               DgvSessions.CurrentCell.ColumnIndex = DgvSessions.Columns("IncludeInReport").Index Then

                ' მყისიერად კომიტი - CheckBox-ის ცვლილების დაფიქსირება
                DgvSessions.CommitEdit(DataGridViewDataErrorContexts.Commit)
                Debug.WriteLine("UC_TherapistReport: CheckBox ცვლილება კომიტირებულია (DirtyState)")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnCurrentCellDirtyStateChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დაყენება გამჭვირვალე სტილით
    ''' </summary>
    Private Sub SetBackgroundColors()
        Try
            Debug.WriteLine("UC_TherapistReport: ფონის ფერების დაყენება")

            ' ღია გამჭვირვალე თეთრი ფერი
            Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

            ' ფილტრების პანელის ფონი
            If pnlFilter IsNot Nothing Then
                pnlFilter.BackColor = transparentWhite
                Debug.WriteLine("UC_TherapistReport: pnlFilter ფონი დაყენებულია")
            Else
                Debug.WriteLine("UC_TherapistReport: ⚠️ pnlFilter არ მოიძებნა!")
            End If

            ' შეიძლება დაემატოს სხვა პანელების ფონებიც
            ' მაგ: თუ გვაქვს სხვა პანელები რომლებსაც ფონი ესაჭიროება

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: SetBackgroundColors შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის ComboBox-ების განახლება პერიოდის მიხედვით
    ''' </summary>
    Private Sub RefreshTherapistComboBoxes()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის ComboBox-ების განახლება")

            If dataProcessor Is Nothing Then
                Debug.WriteLine("UC_TherapistReport: dataProcessor არ არის ინიციალიზებული")
                Return
            End If

            ' ციკლური განახლებების თავიდან აცილება
            isUpdatingTherapist = True

            Try
                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()

                ' ყველა ფილტრირებული მონაცემის მიღება
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                Debug.WriteLine($"UC_TherapistReport: მიღებულია {allSessions.Count} სესია ComboBox-ების შესავსებად")

                ' თერაპევტების ComboBox-ის განახლება
                PopulateTherapistsComboBox(allSessions)

                ' ფილტრების ComboBox-ების რესეტი
                ResetFilterComboBoxes()

            Finally
                isUpdatingTherapist = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: RefreshTherapistComboBoxes შეცდომა: {ex.Message}")
            isUpdatingTherapist = False
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების ComboBox-ის შევსება უნიკალური თერაპევტებით
    ''' </summary>
    ''' <param name="sessions">სესიების სია</param>
    Private Sub PopulateTherapistsComboBox(sessions As List(Of SessionModel))
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტების ComboBox-ის შევსება")

            If CBPer Is Nothing Then
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBPer ComboBox არ არის ხელმისაწვდომი")
                Return
            End If

            CBPer.Items.Clear()

            If sessions Is Nothing OrElse sessions.Count = 0 Then
                CBPer.Enabled = False
                Debug.WriteLine("UC_TherapistReport: სესიები არ არის - CBPer გათიშულია")
                Return
            End If

            ' უნიკალური თერაპევტების მიღება (მხოლოდ ცარიელი არ არის)
            Dim uniqueTherapists = sessions.Where(Function(s) Not String.IsNullOrWhiteSpace(s.TherapistName)) _
                                          .Select(Function(s) s.TherapistName.Trim()) _
                                          .Distinct(StringComparer.OrdinalIgnoreCase) _
                                          .OrderBy(Function(therapist) therapist) _
                                          .ToList()

            ' თერაპევტების დამატება ComboBox-ში
            For Each therapist In uniqueTherapists
                CBPer.Items.Add(therapist)
            Next

            CBPer.Enabled = (uniqueTherapists.Count > 0)

            Debug.WriteLine($"UC_TherapistReport: ჩატვირთულია {uniqueTherapists.Count} უნიკალური თერაპევტი")

            ' თუ მხოლოდ ერთი თერაპევტია, ავტომატურად არჩევა
            If uniqueTherapists.Count = 1 Then
                CBPer.SelectedIndex = 0
                Debug.WriteLine($"UC_TherapistReport: ავტომატურად არჩეულია ერთადერთი თერაპევტი: {uniqueTherapists(0)}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: PopulateTherapistsComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრების ComboBox-ების რესეტი საწყისი მდგომარეობაში
    ''' </summary>
    Private Sub ResetFilterComboBoxes()
        Try
            Debug.WriteLine("UC_TherapistReport: ფილტრების ComboBox-ების რესეტი")

            ' თერაპიის ComboBox (CBTer) რესეტი
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            If cbTherapyType IsNot Nothing Then
                cbTherapyType.Items.Clear()
                cbTherapyType.Items.Add("ყველა")
                cbTherapyType.SelectedIndex = 0
                cbTherapyType.Enabled = False
                Debug.WriteLine("UC_TherapistReport: CBTer ComboBox დარესეტდა")
            End If

            ' დაფინანსების ComboBox (CBDaf) რესეტი
            Dim cbFunding As ComboBox = FindComboBoxRecursive(Me, "CBDaf")
            If cbFunding IsNot Nothing Then
                cbFunding.Items.Clear()
                cbFunding.Items.Add("ყველა")
                cbFunding.SelectedIndex = 0
                cbFunding.Enabled = False
                Debug.WriteLine("UC_TherapistReport: CBDaf ComboBox დარესეტდა")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ResetFilterComboBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპიის ComboBox-ის შევსება კონკრეტული თერაპევტისთვის
    ''' </summary>
    ''' <param name="selectedTherapist">არჩეული თერაპევტის სახელი</param>
    ''' <param name="sessions">ყველა სესიის სია</param>
    Private Sub PopulateTherapyTypeComboBox(selectedTherapist As String, sessions As List(Of SessionModel))
        Try
            Debug.WriteLine($"UC_TherapistReport: თერაპიის ComboBox-ის შევსება თერაპევტისთვის '{selectedTherapist}'")

            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            If cbTherapyType Is Nothing Then
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBTer ComboBox არ მოიძებნა")
                Return
            End If

            cbTherapyType.Items.Clear()

            If String.IsNullOrWhiteSpace(selectedTherapist) OrElse sessions Is Nothing Then
                cbTherapyType.Items.Add("ყველა")
                cbTherapyType.SelectedIndex = 0
                cbTherapyType.Enabled = False
                Debug.WriteLine("UC_TherapistReport: CBTer გათიშული - თერაპევტი არ არის არჩეული")
                Return
            End If

            ' კონკრეტული თერაპევტის თერაპიების მიღება
            Dim therapistTherapies = sessions.Where(Function(s)
                                                        Return s.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase) AndAlso
                                                              Not String.IsNullOrWhiteSpace(s.TherapyType)
                                                    End Function) _
                                            .Select(Function(s) s.TherapyType.Trim()) _
                                            .Distinct(StringComparer.OrdinalIgnoreCase) _
                                            .OrderBy(Function(therapy) therapy) _
                                            .ToList()

            cbTherapyType.Items.Add("ყველა")
            For Each therapy In therapistTherapies
                cbTherapyType.Items.Add(therapy)
            Next

            cbTherapyType.SelectedIndex = 0
            cbTherapyType.Enabled = (therapistTherapies.Count > 0)

            Debug.WriteLine($"UC_TherapistReport: თერაპევტისთვის '{selectedTherapist}' ნაპოვნია {therapistTherapies.Count} თერაპია")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: PopulateTherapyTypeComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დაფინანსების ComboBox-ის შევსება კონკრეტული თერაპევტისთვის
    ''' </summary>
    ''' <param name="selectedTherapist">არჩეული თერაპევტის სახელი</param>
    ''' <param name="sessions">ყველა სესიის სია</param>
    Private Sub PopulateFundingComboBox(selectedTherapist As String, sessions As List(Of SessionModel))
        Try
            Debug.WriteLine($"UC_TherapistReport: დაფინანსების ComboBox-ის შევსება თერაპევტისთვის '{selectedTherapist}'")

            Dim cbFunding As ComboBox = FindComboBoxRecursive(Me, "CBDaf")
            If cbFunding Is Nothing Then
                Debug.WriteLine("UC_TherapistReport: ⚠️ CBDaf ComboBox არ მოიძებნა")
                Return
            End If

            cbFunding.Items.Clear()

            If String.IsNullOrWhiteSpace(selectedTherapist) OrElse sessions Is Nothing Then
                cbFunding.Items.Add("ყველა")
                cbFunding.SelectedIndex = 0
                cbFunding.Enabled = False
                Debug.WriteLine("UC_TherapistReport: CBDaf გათიშული - თერაპევტი არ არის არჩეული")
                Return
            End If

            ' კონკრეტული თერაპევტის დაფინანსებების მიღება
            Dim therapistFunding = sessions.Where(Function(s)
                                                      Return s.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase) AndAlso
                                                            Not String.IsNullOrWhiteSpace(s.Funding)
                                                  End Function) _
                                           .Select(Function(s) s.Funding.Trim()) _
                                           .Distinct(StringComparer.OrdinalIgnoreCase) _
                                           .OrderBy(Function(funding) funding) _
                                           .ToList()

            cbFunding.Items.Add("ყველა")
            For Each funding In therapistFunding
                cbFunding.Items.Add(funding)
            Next

            cbFunding.SelectedIndex = 0
            cbFunding.Enabled = (therapistFunding.Count > 0)

            Debug.WriteLine($"UC_TherapistReport: თერაპევტისთვის '{selectedTherapist}' ნაპოვნია {therapistFunding.Count} დაფინანსება")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: PopulateFundingComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' IList(Of IList(Of Object))-ის SessionModel-ებად გარდაქმნა
    ''' </summary>
    ''' <param name="data">Google Sheets-ის მონაცემები</param>
    ''' <returns>SessionModel-ების სია</returns>
    Private Function ConvertToSessionModels(data As List(Of IList(Of Object))) As List(Of SessionModel)
        Try
            Debug.WriteLine($"UC_TherapistReport: SessionModel-ების გარდაქმნა - {If(data?.Count, 0)} ჩანაწერი")

            Dim sessions As New List(Of SessionModel)()

            If data IsNot Nothing Then
                For i As Integer = 0 To data.Count - 1
                    Try
                        Dim row = data(i)
                        If row.Count >= 12 Then
                            Dim session = SessionModel.FromSheetRow(row)
                            sessions.Add(session)
                        Else
                            Debug.WriteLine($"UC_TherapistReport: მწკრივი {i} არასრული - {row.Count} სვეტი")
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"UC_TherapistReport: SessionModel-ის შექმნის შეცდომა მწკრივი {i}: {ex.Message}")
                        Continue For
                    End Try
                Next
            End If

            Debug.WriteLine($"UC_TherapistReport: წარმატებით გარდაიქმნა {sessions.Count} SessionModel")
            Return sessions

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ConvertToSessionModels შეცდომა: {ex.Message}")
            Return New List(Of SessionModel)()
        End Try
    End Function

    ''' <summary>
    ''' თერაპევტის სპეციფიკური მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadTherapistSpecificData()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის მონაცემების ჩატვირთვა დაიწყო")

            If dataProcessor Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_TherapistReport: სერვისები არ არის ინიციალიზებული")
                Return
            End If

            If isLoadingData Then
                Debug.WriteLine("UC_TherapistReport: მონაცემები უკვე იტვირთება - შეწყვეტა")
                Return
            End If

            isLoadingData = True

            Try
                ' არჩეული თერაპევტის შემოწმება
                Dim selectedTherapist As String = If(CBPer?.SelectedItem?.ToString(), "")

                If String.IsNullOrWhiteSpace(selectedTherapist) Then
                    ' თერაპევტი არ არის არჩეული - ცხრილის გასუფთავება
                    DgvSessions.Rows.Clear()
                    currentTherapistData = Nothing
                    Debug.WriteLine("UC_TherapistReport: თერაპევტი არ არის არჩეული - ცხრილი გასუფთავებულია")
                    Return
                End If

                Debug.WriteLine($"UC_TherapistReport: ტვირთება თერაპევტისთვის: '{selectedTherapist}'")

                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()

                ' ყველა ფილტრირებული მონაცემის მიღება
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                Debug.WriteLine($"UC_TherapistReport: მიღებულია {allSessions.Count} სესია ფილტრაციის შემდეგ")

                ' კონკრეტული თერაპევტის სესიების ფილტრაცია
                currentTherapistData = allSessions.Where(Function(s)
                                                             Return s.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase)
                                                         End Function) _
                                                  .OrderBy(Function(s) s.DateTime) _
                                                  .ToList()

                Debug.WriteLine($"UC_TherapistReport: თერაპევტი '{selectedTherapist}'-ისთვის ნაპოვნია {currentTherapistData.Count} სესია")

                ' მონაცემების ჩატვირთვა DataGridView-ში
                LoadTherapistSessionsToGrid()

            Finally
                isLoadingData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: LoadTherapistSpecificData შეცდომა: {ex.Message}")
            isLoadingData = False
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის სესიების ჩატვირთვა DataGridView-ში დამატებითი ფილტრაციით
    ''' </summary>
    Private Sub LoadTherapistSessionsToGrid()
        Try
            Debug.WriteLine($"UC_TherapistReport: თერაპევტის სესიების ჩატვირთვა DataGridView-ში - {If(currentTherapistData?.Count, 0)} სესია")

            DgvSessions.Rows.Clear()

            If currentTherapistData Is Nothing OrElse currentTherapistData.Count = 0 Then
                Debug.WriteLine("UC_TherapistReport: მონაცემები არ არის ხელმისაწვდომი")
                Return
            End If

            ' დამატებითი ფილტრების მიღება
            Dim selectedTherapyType As String = GetSelectedTherapyType()
            Dim selectedFunding As String = GetSelectedFunding()

            Debug.WriteLine($"UC_TherapistReport: დამატებითი ფილტრები - თერაპია: '{selectedTherapyType}', დაფინანსება: '{selectedFunding}'")

            Dim displayedCount As Integer = 0

            For i As Integer = 0 To currentTherapistData.Count - 1
                Dim session = currentTherapistData(i)

                ' თერაპიის ფილტრაცია
                If Not String.IsNullOrWhiteSpace(selectedTherapyType) AndAlso selectedTherapyType <> "ყველა" Then
                    If Not session.TherapyType.Trim().Equals(selectedTherapyType, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                ' დაფინანსების ფილტრაცია
                If Not String.IsNullOrWhiteSpace(selectedFunding) AndAlso selectedFunding <> "ყველა" Then
                    If Not session.Funding.Trim().Equals(selectedFunding, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                ' ბენეფიციარის სრული სახელი
                Dim beneficiaryFullName As String = $"{session.BeneficiaryName.Trim()} {session.BeneficiarySurname.Trim()}".Trim()

                ' მწკრივის მონაცემები
                Dim rowData As Object() = {
                    session.Id,                                     ' N
                    session.DateTime.ToString("dd.MM.yyyy HH:mm"), ' თარიღი
                    $"{session.Duration}წთ",                        ' ხანგძლივობა
                    beneficiaryFullName,                            ' ბენეფიციარის სრული სახელი
                    session.Funding,                                ' დაფინანსების სვეტი
                    session.Price,                                  ' თანხა
                    session.Status,                                 ' შესრულების სვეტი
                    True,                                           ' ნაგულისხმევად რეპორტში ჩართული
                    "✎"                                             ' რედაქტირების ღილაკი
                }

                ' მწკრივის დამატება
                Dim addedRowIndex = DgvSessions.Rows.Add(rowData)

                ' სესიის ID-ის შენახვა მწკრივში მომავალი გამოყენებისთვის
                DgvSessions.Rows(addedRowIndex).Tag = session.Id

                ' CheckBox-ის მნიშვნელობის ექსპლიციტური დაყენება
                DgvSessions.Rows(addedRowIndex).Cells("IncludeInReport").Value = True

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                Try
                    Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = statusColor
                Catch ex As Exception
                    Debug.WriteLine($"UC_TherapistReport: სტატუსის ფერის შეცდომა: {ex.Message}")
                    ' სტატუსის ფერის შეცდომის შემთხვევაში თეთრი ფონი
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = Color.White
                End Try

                displayedCount += 1
            Next

            ' რეპორტის ჯამების განახლება
            UpdateReportTotals()

            Debug.WriteLine($"UC_TherapistReport: წარმატებით ჩატვირთულია {displayedCount} სესია (ფილტრაციის შემდეგ)")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: LoadTherapistSessionsToGrid შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' არჩეული თერაპიის ტიპის მიღება
    ''' </summary>
    ''' <returns>არჩეული თერაპიის ტიპი ან "ყველა"</returns>
    Private Function GetSelectedTherapyType() As String
        Try
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            Return If(cbTherapyType?.SelectedItem?.ToString(), "ყველა")
        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetSelectedTherapyType შეცდომა: {ex.Message}")
            Return "ყველა"
        End Try
    End Function

    ''' <summary>
    ''' არჩეული დაფინანსების ტიპის მიღება
    ''' </summary>
    ''' <returns>არჩეული დაფინანსების ტიპი ან "ყველა"</returns>
    Private Function GetSelectedFunding() As String
        Try
            Dim cbFunding As ComboBox = FindComboBoxRecursive(Me, "CBDaf")
            Return If(cbFunding?.SelectedItem?.ToString(), "ყველა")
        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetSelectedFunding შეცდომა: {ex.Message}")
            Return "ყველა"
        End Try
    End Function

    ''' <summary>
    ''' კონტროლების სანადიმო ვალიდაცია - განვითარების დროს გამოსაყენებლად
    ''' </summary>
    Private Sub ValidateControlsExistence()
        Try
            Debug.WriteLine("UC_TherapistReport: კონტროლების ვალიდაცია")

            ' ძირითადი კონტროლების შემოწმება
            Dim missingControls As New List(Of String)()

            If DgvSessions Is Nothing Then missingControls.Add("DgvSessions")
            If CBPer Is Nothing Then missingControls.Add("CBPer")
            If DtpDan Is Nothing Then missingControls.Add("DtpDan")
            If DtpMde Is Nothing Then missingControls.Add("DtpMde")
            If pnlFilter Is Nothing Then missingControls.Add("pnlFilter")

            ' ComboBox-ების შემოწმება
            If FindComboBoxRecursive(Me, "CBTer") Is Nothing Then missingControls.Add("CBTer")
            If FindComboBoxRecursive(Me, "CBDaf") Is Nothing Then missingControls.Add("CBDaf")

            ' CheckBox-ების შემოწმება
            Dim missingCheckBoxes As New List(Of String)()
            For i As Integer = 1 To 7
                If FindCheckBoxRecursive(Me, $"CheckBox{i}") Is Nothing Then
                    missingCheckBoxes.Add($"CheckBox{i}")
                End If
            Next

            ' ღილაკების შემოწმება
            If Me.Controls.Find("BtnRef", True).Length = 0 Then missingControls.Add("BtnRef")
            If Me.Controls.Find("BtnAddSchedule", True).Length = 0 Then missingControls.Add("BtnAddSchedule")
            If Me.Controls.Find("btbPrint", True).Length = 0 Then missingControls.Add("btbPrint")
            If Me.Controls.Find("btnToPDF", True).Length = 0 Then missingControls.Add("btnToPDF")
            If Me.Controls.Find("BtnToExcel", True).Length = 0 Then missingControls.Add("BtnToExcel")

            ' შედეგის ლოგირება
            If missingControls.Count > 0 Then
                Debug.WriteLine($"UC_TherapistReport: ⚠️ ნაკლული კონტროლები: {String.Join(", ", missingControls)}")
            End If

            If missingCheckBoxes.Count > 0 Then
                Debug.WriteLine($"UC_TherapistReport: ⚠️ ნაკლული CheckBox-ები: {String.Join(", ", missingCheckBoxes)}")
            End If

            If missingControls.Count = 0 AndAlso missingCheckBoxes.Count = 0 Then
                Debug.WriteLine("UC_TherapistReport: ✅ ყველა კონტროლი ხელმისაწვდომია")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ValidateControlsExistence შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Debug რეჟიმში UI კონტროლების მდგომარეობის ჩვენება
    ''' </summary>
    Private Sub ShowControlsDebugInfo()
        Try
            Debug.WriteLine("UC_TherapistReport: === DEBUG INFO ===")
            Debug.WriteLine($"dataService: {If(dataService IsNot Nothing, "OK", "NULL")}")
            Debug.WriteLine($"dataProcessor: {If(dataProcessor IsNot Nothing, "OK", "NULL")}")
            Debug.WriteLine($"uiManager: {If(uiManager IsNot Nothing, "OK", "NULL")}")
            Debug.WriteLine($"filterManager: {If(filterManager IsNot Nothing, "OK", "NULL")}")
            Debug.WriteLine($"isLoadingData: {isLoadingData}")
            Debug.WriteLine($"isUpdatingTherapist: {isUpdatingTherapist}")
            Debug.WriteLine($"isUpdatingFilters: {isUpdatingFilters}")
            Debug.WriteLine($"currentTherapistData: {If(currentTherapistData?.Count, 0)} სესია")
            Debug.WriteLine($"DgvSessions Rows: {If(DgvSessions?.Rows.Count, 0)}")
            Debug.WriteLine($"CBPer Items: {If(CBPer?.Items.Count, 0)}")
            Debug.WriteLine($"CBPer Selected: '{If(CBPer?.SelectedItem?.ToString(), "არაფერი")}'")

            ' ფილტრების მდგომარეობა
            Debug.WriteLine($"თერაპიის ფილტრი: '{GetSelectedTherapyType()}'")
            Debug.WriteLine($"დაფინანსების ფილტრი: '{GetSelectedFunding()}'")
            Debug.WriteLine($"პერიოდი: {If(DtpDan?.Value.ToString("dd.MM.yyyy"), "N/A")} - {If(DtpMde?.Value.ToString("dd.MM.yyyy"), "N/A")}")

            Debug.WriteLine("UC_TherapistReport: === END DEBUG ===")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ShowControlsDebugInfo შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები"

    ''' <summary>
    ''' ფილტრის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine("UC_TherapistReport: OnFilterChanged")

            If isLoadingData OrElse isUpdatingTherapist Then
                Return
            End If

            ' თერაპევტის ComboBox-ების განახლება
            RefreshTherapistComboBoxes()

            ' მონაცემების განახლება
            LoadTherapistSpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnTherapistChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingTherapist Then Return

            Debug.WriteLine("UC_TherapistReport: თერაპევტი შეიცვალა")

            Dim selectedTherapist As String = CBPer.SelectedItem?.ToString()

            If Not String.IsNullOrEmpty(selectedTherapist) Then
                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                ' თერაპიისა და დაფინანსების ComboBox-ების განახლება
                PopulateTherapyTypeComboBox(selectedTherapist, allSessions)
                PopulateFundingComboBox(selectedTherapist, allSessions)
            Else
                ' თერაპევტი არ არის არჩეული - ყველა ComboBox-ის გასუფთავება
                ResetFilterComboBoxes()
            End If

            ' მონაცემების განახლება
            LoadTherapistSpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnTherapistChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპიის ტიპის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnTherapyTypeChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingTherapist OrElse isUpdatingFilters Then Return

            Debug.WriteLine("UC_TherapistReport: თერაპიის ტიპი შეიცვალა")

            ' მონაცემების განახლება ახალი თერაპიის ფილტრით
            LoadTherapistSessionsToGrid()
            UpdateReportTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnTherapyTypeChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' დაფინანსების შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFundingChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingTherapist OrElse isUpdatingFilters Then Return

            Debug.WriteLine("UC_TherapistReport: დაფინანსება შეიცვალა")

            ' მონაცემების განახლება ახალი დაფინანსების ფილტრით
            LoadTherapistSessionsToGrid()
            UpdateReportTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnFundingChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 გაუმჯობესებული CheckBox-ის შეცვლის ივენთი (მხოლოდ რეალური ცვლილებებისთვის)
    ''' </summary>
    Private Sub OnCheckBoxChanged(sender As Object, e As DataGridViewCellEventArgs)
        Static lastChangeTime As DateTime = DateTime.MinValue
        Static lastChangeRow As Integer = -1

        Try
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
              DgvSessions.Columns(e.ColumnIndex).Name = "IncludeInReport" Then

                ' 🔧 ორმაგი ცვლილების თავიდან აცილება
                Dim currentTime As DateTime = DateTime.Now
                Dim timeDifference = currentTime.Subtract(lastChangeTime).TotalMilliseconds

                If timeDifference < 200 AndAlso lastChangeRow = e.RowIndex Then
                    Debug.WriteLine($"UC_TherapistReport: ორმაგი ცვლილების იგნორირება - მწკრივი {e.RowIndex}")
                    Return
                End If

                lastChangeTime = currentTime
                lastChangeRow = e.RowIndex

                ' CheckBox-ის მნიშვნელობის მიღება
                Dim isIncluded As Boolean = False
                Dim cellValue = DgvSessions.Rows(e.RowIndex).Cells("IncludeInReport").Value

                If cellValue IsNot Nothing Then
                    If TypeOf cellValue Is Boolean Then
                        isIncluded = DirectCast(cellValue, Boolean)
                    Else
                        Boolean.TryParse(cellValue.ToString(), isIncluded)
                    End If
                End If

                Debug.WriteLine($"UC_TherapistReport: CheckBox ცვლილება (CellValueChanged) - მწკრივი {e.RowIndex}, მნიშვნელობა: {isIncluded}")

                ' ვიზუალური სტილის განახლება
                UpdateRowVisualStyle(e.RowIndex, isIncluded)

                ' რეპორტის ჯამების განახლება
                UpdateReportTotals()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnCheckBoxChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 გაუმჯობესებული DataGridView-ის უჯრაზე დაჭერის ივენთი (ორმაგი დაჭერის თავიდან აცილება)
    ''' </summary>
    Private Sub OnDataGridViewCellClick(sender As Object, e As DataGridViewCellEventArgs)
        Static lastClickTime As DateTime = DateTime.MinValue
        Static lastClickRow As Integer = -1
        Static lastClickColumn As Integer = -1

        Try
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

            ' 🔧 ორმაგი დაჭერის თავიდან აცილება
            Dim currentTime As DateTime = DateTime.Now
            Dim timeDifference = currentTime.Subtract(lastClickTime).TotalMilliseconds

            ' თუ იგივე უჯრაზე 300ms-ში დაიჭირა, იგნორირება
            If timeDifference < 300 AndAlso lastClickRow = e.RowIndex AndAlso lastClickColumn = e.ColumnIndex Then
                Debug.WriteLine($"UC_TherapistReport: ორმაგი კლიკის იგნორირება - მწკრივი {e.RowIndex}, სვეტი {e.ColumnIndex}")
                Return
            End If

            ' ახალი კლიკის პარამეტრების შენახვა
            lastClickTime = currentTime
            lastClickRow = e.RowIndex
            lastClickColumn = e.ColumnIndex

            ' CheckBox სვეტზე დაჭერა
            If DgvSessions.Columns(e.ColumnIndex).Name = "IncludeInReport" Then
                Debug.WriteLine($"UC_TherapistReport: CheckBox სვეტზე დაჭერა - მწკრივი {e.RowIndex}")

                ' CheckBox-ის მნიშვნელობის ინვერსია
                Dim currentValue As Boolean = False
                If DgvSessions.Rows(e.RowIndex).Cells("IncludeInReport").Value IsNot Nothing Then
                    Boolean.TryParse(DgvSessions.Rows(e.RowIndex).Cells("IncludeInReport").Value.ToString(), currentValue)
                End If

                ' ახალი მნიშვნელობის დაყენება
                Dim newValue As Boolean = Not currentValue
                DgvSessions.Rows(e.RowIndex).Cells("IncludeInReport").Value = newValue

                ' ვიზუალური სტილის განახლება
                UpdateRowVisualStyle(e.RowIndex, newValue)

                ' რეპორტის ჯამების განახლება
                UpdateReportTotals()

                Debug.WriteLine($"UC_TherapistReport: CheckBox შეიცვალა - {currentValue} → {newValue}")
                Return
            End If

            ' რედაქტირების ღილაკზე დაჭერა
            If DgvSessions.Columns(e.ColumnIndex).Name = "Edit" Then
                Debug.WriteLine($"UC_TherapistReport: რედაქტირების ღილაკზე დაჭერა - მწკრივი {e.RowIndex}")

                Dim sessionId As Integer = 0
                If DgvSessions.Rows(e.RowIndex).Tag IsNot Nothing Then
                    Integer.TryParse(DgvSessions.Rows(e.RowIndex).Tag.ToString(), sessionId)
                End If

                If sessionId > 0 Then
                    Debug.WriteLine($"UC_TherapistReport: რედაქტირება - სესია ID={sessionId}")

                    Try
                        Using editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_TherapistReport")
                            Dim result As DialogResult = editForm.ShowDialog()

                            If result = DialogResult.OK Then
                                RefreshData()
                                MessageBox.Show($"სესია ID={sessionId} წარმატებით განახლდა", "წარმატება",
                                             MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End Using

                    Catch formEx As Exception
                        Debug.WriteLine($"UC_TherapistReport: რედაქტირების ფორმის შეცდომა: {formEx.Message}")
                        MessageBox.Show($"რედაქტირების ფორმის გახსნის შეცდომა: {formEx.Message}", "შეცდომა",
                                      MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Else
                    Debug.WriteLine($"UC_TherapistReport: ⚠️ სესიის ID ვერ მოიძებნა მწკრივი {e.RowIndex}-სთვის")
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: OnDataGridViewCellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' განახლების ღილაკი
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            RefreshData()
        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: BtnRef_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი ჩანაწერის დამატების ღილაკი
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Try
            Debug.WriteLine("UC_TherapistReport: ახალი სესიის დამატება")

            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    frm.Focus()
                    Return
                End If
            Next

            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' NewRecordForm-ის გახსნა
            Using newRecordForm As New NewRecordForm(dataService, "სესია", userEmail, "UC_TherapistReport")
                Dim result = newRecordForm.ShowDialog()

                If result = DialogResult.OK Then
                    Debug.WriteLine("UC_TherapistReport: სესია წარმატებით დაემატა")
                    RefreshData()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: BtnAddSchedule_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ბეჭდვის ღილაკი - თერაპევტის რეპორტისთვის
    ''' </summary>
    Private Sub BtbPrint_Click(sender As Object, e As EventArgs) Handles btbPrint.Click
        Try
            Debug.WriteLine("UC_TherapistReport: ბეჭდვის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' თერაპევტის ინფორმაციის შემოწმება
            Dim therapistInfo = GetCurrentTherapistInfo()
            If String.IsNullOrEmpty(therapistInfo) Then
                MessageBox.Show("აირჩიეთ თერაპევტი ბეჭდვისთვის", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ბეჭდვის ტიპის არჩევა
            Dim printTypeResult As DialogResult = MessageBox.Show(
               "რომელი ტიპის ბეჭდვა გსურთ?" & Environment.NewLine & Environment.NewLine &
               "დიახ - თერაპევტის რეპორტი" & Environment.NewLine &
               "არა - ჩვეულებრივი ცხრილის ბეჭდვა" & Environment.NewLine &
               "გაუქმება - ოპერაციის შეწყვეტა",
               "ბეჭდვის ტიპის არჩევა",
               MessageBoxButtons.YesNoCancel,
               MessageBoxIcon.Question)

            Select Case printTypeResult
                Case DialogResult.Yes
                    ' თერაპევტის რეპორტის ბეჭდვა
                    PrintTherapistReport()

                Case DialogResult.No
                    ' ჩვეულებრივი ცხრილის ბეჭდვა
                    Using printService As New AdvancedDataGridViewPrintService(DgvSessions)
                        printService.ShowFullPrintDialog()
                    End Using

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_TherapistReport: ბეჭდვა გაუქმებულია")
            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: BtbPrint_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Excel ექსპორტის ღილაკი - თერაპევტის რეპორტი Excel ფორმატში
    ''' </summary>
    Private Sub BtnToExcel_Click(sender As Object, e As EventArgs) Handles BtnToExcel.Click
        Try
            Debug.WriteLine("UC_TherapistReport: Excel ექსპორტის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("Excel ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' თერაპევტის ინფორმაციის შემოწმება
            Dim therapistInfo = GetCurrentTherapistInfo()
            If String.IsNullOrEmpty(therapistInfo) Then
                MessageBox.Show("აირჩიეთ თერაპევტი Excel ექსპორტისთვის", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' რეპორტის ვალიდობის შემოწმება
            If Not IsReportValid() Then
                MessageBox.Show("Excel ექსპორტისთვის აირჩიეთ თერაპევტი და მინიმუმ ერთი სესია", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' CSV ფაილის შექმნა Excel-ისთვის
            CreateExcelReportFile()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: BtnToExcel_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"Excel ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' PDF ექსპორტის ღილაკი - თერაპევტის რეპორტისთვის
    ''' </summary>
    Private Sub btnToPDF_Click(sender As Object, e As EventArgs) Handles btnToPDF.Click
        Try
            Debug.WriteLine("UC_TherapistReport: PDF ექსპორტის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("PDF ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' თერაპევტის ინფორმაციის შემოწმება
            Dim therapistInfo = GetCurrentTherapistInfo()
            If String.IsNullOrEmpty(therapistInfo) Then
                MessageBox.Show("აირჩიეთ თერაპევტი PDF ექსპორტისთვის", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' PDF ექსპორტის ტიპის არჩევა
            Dim pdfTypeResult As DialogResult = MessageBox.Show(
               "რომელი ტიპის PDF ექსპორტი გსურთ?" & Environment.NewLine & Environment.NewLine &
               "დიახ - თერაპევტის რეპორტი (PDF)" & Environment.NewLine &
               "არა - ცხრილის PDF ექსპორტი" & Environment.NewLine &
               "გაუქმება - ოპერაციის შეწყვეტა",
               "PDF ექსპორტის ტიპის არჩევა",
               MessageBoxButtons.YesNoCancel,
               MessageBoxIcon.Question)

            Select Case pdfTypeResult
                Case DialogResult.Yes
                    ' თერაპევტის რეპორტის PDF ექსპორტი
                    ExportTherapistReportToPDF()

                Case DialogResult.No
                    ' ჩვეულებრივი ცხრილის PDF ექსპორტი
                    Using exportService As New SimplePDFExportService(DgvSessions)
                        exportService.ShowFullExportDialog()
                    End Using

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_TherapistReport: PDF ექსპორტი გაუქმებულია")
            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: btnToPDF_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "თერაპევტის სპეციფიკური მეთოდები"

    ''' <summary>
    ''' მიმდინარე თერაპევტის ინფორმაციის მიღება
    ''' </summary>
    ''' <returns>თერაპევტის სახელი</returns>
    Public Function GetCurrentTherapistInfo() As String
        Try
            Dim selectedTherapist As String = CBPer.SelectedItem?.ToString()
            Return If(String.IsNullOrEmpty(selectedTherapist), "", selectedTherapist)

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetCurrentTherapistInfo შეცდომა: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' რეპორტის ჯამური თანხის გამოთვლა (მხოლოდ მონიშნული სესიები)
    ''' </summary>
    ''' <returns>რეპორტის ჯამური თანხა</returns>
    Public Function GetReportTotalAmount() As Decimal
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim total As Decimal = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' შევამოწმოთ რეპორტში ჩართვის CheckBox
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    ' თუ სესია ჩართულია რეპორტში
                    If isIncluded AndAlso row.Cells("Price").Value IsNot Nothing Then
                        Dim price As Decimal
                        If Decimal.TryParse(row.Cells("Price").Value.ToString(), price) Then
                            total += price
                        End If
                    End If

                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"UC_TherapistReport: რეპორტის ჯამური თანხა: {total:N2}")
            Return total

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetReportTotalAmount შეცდომა: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' რეპორტში ჩართული სესიების რაოდენობა
    ''' </summary>
    ''' <returns>ჩართული სესიების რაოდენობა</returns>
    Public Function GetIncludedSessionsCount() As Integer
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim count As Integer = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        count += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            Return count

        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' რეპორტის ღირებულებების განახლება
    ''' </summary>
    Private Sub UpdateReportTotals()
        Try
            Debug.WriteLine("UC_TherapistReport: რეპორტის ღირებულებების განახლება")

            Dim totalAmount = GetReportTotalAmount()
            Dim includedCount = GetIncludedSessionsCount()

            Debug.WriteLine($"UC_TherapistReport: ჯამური თანხა: {totalAmount:N2}, ჩართული სესიები: {includedCount}")

            ' აქ შეიძლება დავამატოთ UI ელემენტების განახლება თანხისა და რაოდენობის ჩვენებისთვის

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: UpdateReportTotals შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სესიის რეპორტში ჩართვა
    ''' </summary>
    Public Sub IncludeAllInReport()
        Try
            Debug.WriteLine("UC_TherapistReport: ყველა სესიის რეპორტში ჩართვა")

            If DgvSessions Is Nothing Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("IncludeInReport").Value = True
                    ' ვიზუალური სტილის განახლება
                    UpdateRowVisualStyle(row.Index, True)
                Catch
                    Continue For
                End Try
            Next

            UpdateReportTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: IncludeAllInReport შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სესიის რეპორტიდან ამოღება
    ''' </summary>
    Public Sub ExcludeAllFromReport()
        Try
            Debug.WriteLine("UC_TherapistReport: ყველა სესიის რეპორტიდან ამოღება")

            If DgvSessions Is Nothing Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("IncludeInReport").Value = False
                    ' ვიზუალური სტილის განახლება
                    UpdateRowVisualStyle(row.Index, False)
                Catch
                    Continue For
                End Try
            Next

            UpdateReportTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ExcludeAllFromReport შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მწკრივის ვიზუალური სტილის განახლება CheckBox-ის მდგომარეობის მიხედვით
    ''' </summary>
    ''' <param name="rowIndex">მწკრივის ინდექსი</param>
    ''' <param name="isIncluded">რეპორტში ჩართულია თუ არა</param>
    Private Sub UpdateRowVisualStyle(rowIndex As Integer, isIncluded As Boolean)
        Try
            If rowIndex < 0 OrElse rowIndex >= DgvSessions.Rows.Count Then Return

            Dim row As DataGridViewRow = DgvSessions.Rows(rowIndex)

            If isIncluded Then
                ' ჩართული - ნორმალური სტილი (სტატუსის ფერით)
                Try
                    ' სტატუსის ფერის აღდგენა
                    Dim statusText As String = If(row.Cells("Status").Value?.ToString(), "")
                    If Not String.IsNullOrEmpty(statusText) Then
                        ' SessionModel-ის მიღება currentTherapistData-დან
                        Dim sessionId As Integer = 0
                        If row.Tag IsNot Nothing AndAlso Integer.TryParse(row.Tag.ToString(), sessionId) Then
                            Dim session = currentTherapistData?.FirstOrDefault(Function(s) s.Id = sessionId)
                            If session IsNot Nothing Then
                                Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                                row.DefaultCellStyle.BackColor = statusColor
                            End If
                        End If
                    End If
                Catch
                    row.DefaultCellStyle.BackColor = Color.White
                End Try

                ' ნორმალური ფონტი
                row.DefaultCellStyle.Font = DgvSessions.DefaultCellStyle.Font
                row.DefaultCellStyle.ForeColor = Color.Black

            Else
                ' ამოღებული - ნაცრისფერი და გადახაზული
                row.DefaultCellStyle.BackColor = Color.LightGray
                row.DefaultCellStyle.ForeColor = Color.DarkGray

                ' გადახაზული ფონტი
                Dim currentFont As System.Drawing.Font = DgvSessions.DefaultCellStyle.Font
                If currentFont IsNot Nothing Then
                    Dim strikeFont As New System.Drawing.Font(currentFont, FontStyle.Strikeout)
                    row.DefaultCellStyle.Font = strikeFont
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: UpdateRowVisualStyle შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის რეპორტის ბეჭდვა
    ''' </summary>
    Private Sub PrintTherapistReport()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის რეპორტის ბეჭდვა")

            ' რეპორტის ვალიდობის შემოწმება
            If Not IsReportValid() Then
                MessageBox.Show("რეპორტის ბეჭდვისთვის აირჩიეთ თერაპევტი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' HTML რეპორტის შექმნა და ბეჭდვა
            CreateReportHTML(True) ' True = ბეჭდვისთვის

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: PrintTherapistReport შეცდომა: {ex.Message}")
            MessageBox.Show($"რეპორტის ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის რეპორტის PDF ექსპორტი
    ''' </summary>
    Private Sub ExportTherapistReportToPDF()
        Try
            Debug.WriteLine("UC_TherapistReport: თერაპევტის რეპორტის PDF ექსპორტი")

            ' რეპორტის ვალიდობის შემოწმება
            If Not IsReportValid() Then
                MessageBox.Show("რეპორტის ექსპორტისთვის აირჩიეთ თერაპევტი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ფაილის ადგილის არჩევა
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "HTML ფაილები (*.html)|*.html"
                saveDialog.Title = "რეპორტის შენახვა"

                Dim therapistName = GetCurrentTherapistInfo().Replace(" ", "_")
                saveDialog.FileName = $"რეპორტი_{therapistName}_{DateTime.Now:yyyyMMdd}.html"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    CreateReportHTML(False, saveDialog.FileName) ' False = ფაილისთვის
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ExportTherapistReportToPDF შეცდომა: {ex.Message}")
            MessageBox.Show($"რეპორტის PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Excel რეპორტის ფაილის შექმნა (CSV ფორმატით)
    ''' </summary>
    Private Sub CreateExcelReportFile()
        Try
            Debug.WriteLine("UC_TherapistReport: Excel რეპორტის ფაილის შექმნა")

            Dim therapistName = GetCurrentTherapistInfo().Replace(" ", "_")
            Dim period = $"{DtpDan.Value:dd.MM.yyyy}-{DtpMde.Value:dd.MM.yyyy}"

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "CSV ფაილები (*.csv)|*.csv|Excel ფაილები (*.xlsx)|*.xlsx"
                saveDialog.Title = "რეპორტის Excel ექსპორტი"
                saveDialog.FileName = $"რეპორტი_{therapistName}_{period}_{DateTime.Now:yyyyMMdd}.csv"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    ' CSV ფაილის შექმნა
                    Dim csv As New System.Text.StringBuilder()
                    Dim utf8WithBom As New System.Text.UTF8Encoding(True)

                    ' რეპორტის სათაური
                    csv.AppendLine("""თერაპევტის მომსახურების რეპორტი""")
                    csv.AppendLine("""შპს """"ბავშვთა და მოზარდთა განვითარების, აბილიტაციისა და რეაბილიტაციის ცენტრი - პროსპერო""""""")
                    csv.AppendLine()
                    csv.AppendLine($"""თერაპევტი"",""{EscapeCSV(GetCurrentTherapistInfo())}""")
                    csv.AppendLine($"""პერიოდი"",""{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}""")
                    csv.AppendLine($"""შექმნილია"",""{DateTime.Now:dd.MM.yyyy HH:mm}""")
                    csv.AppendLine()

                    ' სათაურები
                    csv.AppendLine("""N"",""თარიღი"",""ხანგძლ."",""ბენეფიციარი"",""დაფინანსება"",""თანხა (₾)"",""შესრულება""")

                    ' მონაცემები (მხოლოდ რეპორტში ჩართული)
                    Dim reportNumber As Integer = 1
                    Dim totalAmount As Decimal = 0

                    For Each row As DataGridViewRow In DgvSessions.Rows
                        Try
                            ' შევამოწმოთ რეპორტში ჩართვა
                            Dim isIncluded As Boolean = True
                            If row.Cells("IncludeInReport").Value IsNot Nothing Then
                                Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                            End If

                            If isIncluded Then
                                Dim dateTime As String = If(row.Cells("DateTime").Value?.ToString(), "")
                                Dim duration As String = If(row.Cells("Duration").Value?.ToString(), "")
                                Dim beneficiary As String = If(row.Cells("BeneficiaryName").Value?.ToString(), "")
                                Dim funding As String = If(row.Cells("Funding").Value?.ToString(), "")
                                Dim status As String = If(row.Cells("Status").Value?.ToString(), "")

                                Dim price As Decimal = 0
                                If row.Cells("Price").Value IsNot Nothing Then
                                    Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                                End If

                                totalAmount += price

                                csv.AppendLine($"""{reportNumber}"",""{EscapeCSV(dateTime)}"",""{EscapeCSV(duration)}"",""{EscapeCSV(beneficiary)}"",""{EscapeCSV(funding)}"",""{price:N2}"",""{EscapeCSV(status)}""")
                                reportNumber += 1
                            End If

                        Catch
                            Continue For
                        End Try
                    Next

                    ' ჯამი
                    csv.AppendLine()
                    csv.AppendLine($"""ჯამური თანხა:"","","","","",""{totalAmount:N2}"",""""")
                    csv.AppendLine($"""სულ სესიები:"","","","","",""{GetIncludedSessionsCount()}"",""""")

                    ' ფაილის ჩაწერა
                    System.IO.File.WriteAllText(saveDialog.FileName, csv.ToString(), utf8WithBom)

                    Debug.WriteLine("UC_TherapistReport: Excel რეპორტი შეიქმნა")
                    MessageBox.Show($"Excel რეპორტი წარმატებით შეიქმნა:{Environment.NewLine}{saveDialog.FileName}", "წარმატება",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' ფაილის გახსნის შეთავაზება
                    Dim openResult As DialogResult = MessageBox.Show("გსურთ Excel ფაილის გახსნა?", "ფაილის გახსნა",
                                                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If openResult = DialogResult.Yes Then
                        System.Diagnostics.Process.Start(saveDialog.FileName)
                    End If
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: CreateExcelReportFile შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' რეპორტის HTML-ის შექმნა
    ''' </summary>
    ''' <param name="forPrinting">True თუ ბეჭდვისთვის, False თუ ფაილისთვის</param>
    ''' <param name="filePath">ფაილის მისამართი (ფაილისთვის)</param>
    Private Sub CreateReportHTML(forPrinting As Boolean, Optional filePath As String = "")
        Try
            Debug.WriteLine($"UC_TherapistReport: რეპორტის HTML შექმნა - ბეჭდვისთვის: {forPrinting}")

            Dim html As New System.Text.StringBuilder()
            Dim therapistName = GetCurrentTherapistInfo()
            Dim period = $"{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"

            ' HTML დოკუმენტის შექმნა
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html lang=""ka"">")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>თერაპევტის რეპორტი - " & therapistName & "</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        @page { size: A4; margin: 20mm; }")
            html.AppendLine("        @media print { .no-print { display: none; } }")
            html.AppendLine("        body { font-family: 'Sylfaen', Arial, sans-serif; font-size: 12px; line-height: 1.4; }")
            html.AppendLine("        .report-header { text-align: center; margin-bottom: 30px; }")
            html.AppendLine("        .report-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }")
            html.AppendLine("        .company-info { font-size: 11px; margin-bottom: 20px; }")
            html.AppendLine("        .therapist-info { border: 1px solid #333; padding: 15px; margin: 20px 0; background: #f9f9f9; }")
            html.AppendLine("        table { width: 100%; border-collapse: collapse; margin: 20px 0; }")
            html.AppendLine("        th, td { padding: 8px 4px; border: 1px solid #333; text-align: left; vertical-align: top; }")
            html.AppendLine("        th { background-color: #ddd; font-weight: bold; text-align: center; }")
            html.AppendLine("        .amount { text-align: right; }")
            html.AppendLine("        .total-section { margin-top: 20px; text-align: right; }")
            html.AppendLine("        .total-amount { font-size: 14px; font-weight: bold; }")
            html.AppendLine("        .signature-section { margin-top: 40px; text-align: center; }")
            html.AppendLine("        .print-button { padding: 15px 30px; font-size: 16px; background: #007bff; color: white; border: none; border-radius: 5px; }")
            html.AppendLine("    </style>")
            html.AppendLine("</head>")
            html.AppendLine("<body>")

            ' ბეჭდვის ღილაკი (მხოლოდ ბრაუზერისთვის)
            If Not forPrinting Then
                html.AppendLine("    <div class=""no-print"" style=""text-align: center; margin: 20px 0;"">")
                html.AppendLine("        <button class=""print-button"" onclick=""window.print(); setTimeout(() => window.close(), 1000);"">")
                html.AppendLine("            🖨️ რეპორტის ბეჭდვა PDF-ად</button>")
                html.AppendLine("        <p>ღილაკზე დაჭერის შემდეგ აირჩიეთ ""Microsoft Print to PDF""</p>")
                html.AppendLine("    </div>")
            End If

            ' რეპორტის სათაური
            html.AppendLine("    <div class=""report-header"">")
            html.AppendLine("        <div class=""report-title"">თერაპევტის მომსახურების რეპორტი</div>")
            html.AppendLine("        <div class=""company-info"">")
            html.AppendLine("            შპს ""ბავშვთა და მოზარდთა განვითარების, აბილიტაციისა და რეაბილიტაციის ცენტრი - პროსპერო""<br>")
            html.AppendLine("            მისამართი: [კომპანიის მისამართი]<br>")
            html.AppendLine("            ტელ: [ტელეფონი] | ელ-ფოსტა: [ელფოსტა]")
            html.AppendLine("        </div>")
            html.AppendLine("    </div>")

            ' თერაპევტის ინფორმაცია
            html.AppendLine("    <div class=""therapist-info"">")
            html.AppendLine($"        <strong>თერაპევტი:</strong> {EscapeHtml(therapistName)}<br>")
            html.AppendLine($"        <strong>პერიოდი:</strong> {period}<br>")
            html.AppendLine($"        <strong>რეპორტის შექმნის თარიღი:</strong> {DateTime.Now:dd.MM.yyyy HH:mm}")
            html.AppendLine("    </div>")

            ' მომსახურებების ცხრილი
            html.AppendLine("    <table>")
            html.AppendLine("        <thead>")
            html.AppendLine("            <tr>")
            html.AppendLine("                <th style=""width: 30px;"">N</th>")
            html.AppendLine("                <th style=""width: 100px;"">თარიღი</th>")
            html.AppendLine("                <th style=""width: 60px;"">ხანგძლ.</th>")
            html.AppendLine("                <th style=""width: 200px;"">ბენეფიციარი</th>")
            html.AppendLine("                <th style=""width: 100px;"">დაფინანსება</th>")
            html.AppendLine("                <th style=""width: 80px;"">თანხა (₾)</th>")
            html.AppendLine("                <th style=""width: 110px;"">შესრულება</th>")
            html.AppendLine("            </tr>")
            html.AppendLine("        </thead>")
            html.AppendLine("        <tbody>")

            ' მომსახურებების სია (მხოლოდ რეპორტში ჩართული)
            Dim reportNumber As Integer = 1
            Dim totalAmount As Decimal = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' შევამოწმოთ რეპორტში ჩართვა
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        Dim dateTime As String = If(row.Cells("DateTime").Value?.ToString(), "")
                        Dim duration As String = If(row.Cells("Duration").Value?.ToString(), "")
                        Dim beneficiary As String = If(row.Cells("BeneficiaryName").Value?.ToString(), "")
                        Dim funding As String = If(row.Cells("Funding").Value?.ToString(), "")
                        Dim status As String = If(row.Cells("Status").Value?.ToString(), "")

                        Dim price As Decimal = 0
                        If row.Cells("Price").Value IsNot Nothing Then
                            Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                        End If

                        totalAmount += price

                        html.AppendLine("            <tr>")
                        html.AppendLine($"                <td style=""text-align: center;"">{reportNumber}</td>")
                        html.AppendLine($"                <td>{EscapeHtml(dateTime)}</td>")
                        html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(duration)}</td>")
                        html.AppendLine($"                <td>{EscapeHtml(beneficiary)}</td>")
                        html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(funding)}</td>")
                        html.AppendLine($"                <td class=""amount"">{price:N2}</td>")
                        html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(status)}</td>")
                        html.AppendLine("            </tr>")

                        reportNumber += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            html.AppendLine("        </tbody>")
            html.AppendLine("    </table>")

            ' ჯამური ინფორმაცია
            html.AppendLine("    <div class=""total-section"">")
            html.AppendLine($"        <div class=""total-amount"">სულ მომსახურებები: {GetIncludedSessionsCount()}</div>")
            html.AppendLine($"        <div class=""total-amount"">ჯამური თანხა: {totalAmount:N2} ₾</div>")
            html.AppendLine("    </div>")

            ' ხელმოწერის სექცია
            html.AppendLine("    <div class=""signature-section"">")
            html.AppendLine("        <p><strong>ცენტრის დირექტორი:</strong></p>")
            html.AppendLine("        <p>თეა ჩანადირი MD PhD DBP</p>")
            html.AppendLine("        <p>მედიცინის დოქტორი, განვითარების და ქცევის პედიატრი</p>")
            html.AppendLine("        <br><br>")
            html.AppendLine("        <p>ხელმოწერა: ____________________</p>")
            html.AppendLine($"        <p>თარიღი: {DateTime.Now:dd.MM.yyyy}</p>")
            html.AppendLine("    </div>")

            html.AppendLine("</body>")
            html.AppendLine("</html>")

            ' ფაილის შენახვა ან დროებითი ფაილის შექმნა ბეჭდვისთვის
            Dim finalFilePath As String

            If forPrinting Then
                ' დროებითი ფაილი ბეჭდვისთვის
                finalFilePath = System.IO.Path.GetTempPath() & $"therapist_report_{DateTime.Now:yyyyMMddHHmmss}.html"
            Else
                ' მომხმარებლის მიერ არჩეული ფაილი
                finalFilePath = filePath
            End If

            ' HTML ფაილის ჩაწერა
            System.IO.File.WriteAllText(finalFilePath, html.ToString(), System.Text.Encoding.UTF8)

            Debug.WriteLine($"UC_TherapistReport: რეპორტის HTML შეიქმნა - {finalFilePath}")

            If forPrinting Then
                ' ბეჭდვისთვის - ფაილის გახსნა
                System.Diagnostics.Process.Start(finalFilePath)
                MessageBox.Show("რეპორტი ბრაუზერში გაიხსნა ბეჭდვისთვის", "ინფორმაცია",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' ფაილისთვის - შეტყობინება და გახსნის შეთავაზება
                MessageBox.Show($"რეპორტი წარმატებით შეიქმნა:{Environment.NewLine}{finalFilePath}" & Environment.NewLine & Environment.NewLine &
                       "დააჭირეთ ფაილში 'რეპორტის ბეჭდვა PDF-ად' ღილაკს", "წარმატება",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)

                Dim openResult As DialogResult = MessageBox.Show("გსურთ რეპორტის ფაილის გახსნა?", "ფაილის გახსნა",
                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If openResult = DialogResult.Yes Then
                    System.Diagnostics.Process.Start(finalFilePath)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: CreateReportHTML შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' რეპორტის ვალიდობის შემოწმება
    ''' </summary>
    ''' <returns>True თუ რეპორტი ვალიდურია</returns>
    Private Function IsReportValid() As Boolean
        Try
            Dim therapistName = GetCurrentTherapistInfo()
            Dim includedCount = GetIncludedSessionsCount()

            Return Not String.IsNullOrEmpty(therapistName) AndAlso includedCount > 0

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: IsReportValid შეცდომა: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' HTML ტექსტის escape
    ''' </summary>
    Private Function EscapeHtml(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""

        text = text.Replace("&", "&amp;")
        text = text.Replace("<", "&lt;")
        text = text.Replace(">", "&gt;")
        text = text.Replace("""", "&quot;")
        text = text.Replace("'", "&#39;")

        Return text
    End Function

    ''' <summary>
    ''' CSV ველის escape
    ''' </summary>
    Private Function EscapeCSV(field As String) As String
        If String.IsNullOrEmpty(field) Then Return ""

        ' ციტატების გადვოება და შემოფარება
        field = field.Replace("""", """""")
        Return field
    End Function

#End Region

#Region "რეპორტის მართვის საჯარო მეთოდები"

    ''' <summary>
    ''' მონიშნული სესიების რეპორტში ჩართვა/ამოღება
    ''' </summary>
    Public Sub ToggleSelectedSessionsReport()
        Try
            Debug.WriteLine("UC_TherapistReport: მონიშნული სესიების რეპორტში ჩართვა/ამოღება")

            If DgvSessions Is Nothing OrElse DgvSessions.SelectedRows.Count = 0 Then
                MessageBox.Show("აირჩიეთ სესიები რეპორტის კონტროლისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            For Each row As DataGridViewRow In DgvSessions.SelectedRows
                Try
                    ' მიმდინარე მდგომარეობის შებრუნება
                    Dim currentValue As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), currentValue)
                    End If

                    row.Cells("IncludeInReport").Value = Not currentValue

                    ' ვიზუალური სტილის განახლება
                    UpdateRowVisualStyle(row.Index, Not currentValue)

                Catch
                    Continue For
                End Try
            Next

            UpdateReportTotals()

            Debug.WriteLine($"UC_TherapistReport: {DgvSessions.SelectedRows.Count} სესიის მდგომარეობა შეიცვალა")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ToggleSelectedSessionsReport შეცდომა: {ex.Message}")
            MessageBox.Show($"მონიშნული სესიების კონტროლის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' რეპორტის სტატისტიკის მიღება
    ''' </summary>
    ''' <returns>სტატისტიკის ტექსტი</returns>
    Public Function GetReportStatistics() As String
        Try
            Dim therapistName = GetCurrentTherapistInfo()
            Dim totalSessions = If(DgvSessions?.Rows.Count, 0)
            Dim includedSessions = GetIncludedSessionsCount()
            Dim totalAmount = GetReportTotalAmount()
            Dim period = $"{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"

            ' განახლებული თერაპიისა და დაფინანსების ფილტრების ინფორმაცია
            Dim selectedTherapy = GetSelectedTherapyType()
            Dim selectedFunding = GetSelectedFunding()

            Dim statisticsText As String = $"თერაპევტი: {therapistName}" & Environment.NewLine &
                   $"პერიოდი: {period}" & Environment.NewLine

            ' ფილტრების ინფორმაცია
            If Not String.IsNullOrEmpty(selectedTherapy) AndAlso selectedTherapy <> "ყველა" Then
                statisticsText += $"თერაპიის ფილტრი: {selectedTherapy}" & Environment.NewLine
            End If

            If Not String.IsNullOrEmpty(selectedFunding) AndAlso selectedFunding <> "ყველა" Then
                statisticsText += $"დაფინანსების ფილტრი: {selectedFunding}" & Environment.NewLine
            End If

            statisticsText += $"სულ სესიები: {totalSessions}" & Environment.NewLine &
                             $"რეპორტში ჩართული: {includedSessions}" & Environment.NewLine &
                             $"ჯამური თანხა: {totalAmount:N2} ₾" & Environment.NewLine &
                             $"საშუალო თანხა სესიაზე: {If(includedSessions > 0, (totalAmount / includedSessions).ToString("N2"), "0")} ₾"

            Debug.WriteLine($"UC_TherapistReport: სტატისტიკა მომზადებულია - {includedSessions}/{totalSessions} სესია")

            Return statisticsText

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetReportStatistics შეცდომა: {ex.Message}")
            Return "სტატისტიკის მიღების შეცდომა"
        End Try
    End Function

    ''' <summary>
    ''' რეპორტის დეტალური ანალიზის მიღება
    ''' </summary>
    ''' <returns>დეტალური ანალიზის ტექსტი</returns>
    Public Function GetDetailedReportAnalysis() As String
        Try
            Debug.WriteLine("UC_TherapistReport: დეტალური ანალიზის მომზადება")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return "ანალიზისთვის მონაცემები არ არის ხელმისაწვდომი"
            End If

            Dim therapistName = GetCurrentTherapistInfo()
            Dim analysisText As New System.Text.StringBuilder()

            ' ძირითადი ინფორმაცია
            analysisText.AppendLine($"თერაპევტის დეტალური ანალიზი: {therapistName}")
            analysisText.AppendLine($"პერიოდი: {DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")
            analysisText.AppendLine($"ანალიზის თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}")
            analysisText.AppendLine(New String("="c, 60))

            ' სტატუსების მიხედვით დაჯგუფება
            Dim statusGroups = New Dictionary(Of String, List(Of DataGridViewRow))()
            Dim statusTotals = New Dictionary(Of String, Decimal)()

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' მხოლოდ რეპორტში ჩართული სესიები
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        Dim status As String = If(row.Cells("Status").Value?.ToString(), "უცნობი")

                        If Not statusGroups.ContainsKey(status) Then
                            statusGroups(status) = New List(Of DataGridViewRow)()
                            statusTotals(status) = 0
                        End If

                        statusGroups(status).Add(row)

                        ' თანხის დამატება
                        Dim price As Decimal = 0
                        If row.Cells("Price").Value IsNot Nothing Then
                            Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                        End If
                        statusTotals(status) += price
                    End If

                Catch
                    Continue For
                End Try
            Next

            ' სტატუსების ანალიზი
            analysisText.AppendLine("📊 სტატუსების მიხედვით ანალიზი:")
            For Each kvp In statusGroups.OrderByDescending(Function(x) x.Value.Count)
                Dim status = kvp.Key
                Dim count = kvp.Value.Count
                Dim total = statusTotals(status)
                Dim percentage = If(GetIncludedSessionsCount() > 0, (count * 100.0 / GetIncludedSessionsCount()).ToString("N1"), "0")

                analysisText.AppendLine($"   • {status}: {count} სესია ({percentage}%) - {total:N2} ₾")
            Next

            analysisText.AppendLine()

            ' დაფინანსების მიხედვით ანალიზი
            Dim fundingGroups = New Dictionary(Of String, Integer)()
            Dim fundingTotals = New Dictionary(Of String, Decimal)()

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        Dim funding As String = If(row.Cells("Funding").Value?.ToString(), "უცნობი")

                        If Not fundingGroups.ContainsKey(funding) Then
                            fundingGroups(funding) = 0
                            fundingTotals(funding) = 0
                        End If

                        fundingGroups(funding) += 1

                        Dim price As Decimal = 0
                        If row.Cells("Price").Value IsNot Nothing Then
                            Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                        End If
                        fundingTotals(funding) += price
                    End If

                Catch
                    Continue For
                End Try
            Next

            analysisText.AppendLine("💰 დაფინანსების მიხედვით ანალიზი:")
            For Each kvp In fundingGroups.OrderByDescending(Function(x) x.Value)
                Dim funding = kvp.Key
                Dim count = kvp.Value
                Dim total = fundingTotals(funding)
                Dim percentage = If(GetIncludedSessionsCount() > 0, (count * 100.0 / GetIncludedSessionsCount()).ToString("N1"), "0")

                analysisText.AppendLine($"   • {funding}: {count} სესია ({percentage}%) - {total:N2} ₾")
            Next

            analysisText.AppendLine()

            ' თვეების მიხედვით ანალიზი
            Dim monthlyData = New Dictionary(Of String, Integer)()
            Dim monthlyTotals = New Dictionary(Of String, Decimal)()

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInReport").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        Dim dateTimeStr As String = If(row.Cells("DateTime").Value?.ToString(), "")
                        Dim sessionDate As DateTime

                        If DateTime.TryParse(dateTimeStr, sessionDate) Then
                            Dim monthKey As String = sessionDate.ToString("yyyy-MM")

                            If Not monthlyData.ContainsKey(monthKey) Then
                                monthlyData(monthKey) = 0
                                monthlyTotals(monthKey) = 0
                            End If

                            monthlyData(monthKey) += 1

                            Dim price As Decimal = 0
                            If row.Cells("Price").Value IsNot Nothing Then
                                Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                            End If
                            monthlyTotals(monthKey) += price
                        End If
                    End If

                Catch
                    Continue For
                End Try
            Next

            analysisText.AppendLine("📅 თვეების მიხედვით ანალიზი:")
            For Each kvp In monthlyData.OrderBy(Function(x) x.Key)
                Dim monthKey = kvp.Key
                Dim count = kvp.Value
                Dim total = monthlyTotals(monthKey)

                Try
                    Dim monthDate = DateTime.ParseExact(monthKey, "yyyy-MM", Nothing)
                    Dim monthName = monthDate.ToString("MMMM yyyy", New System.Globalization.CultureInfo("ka-GE"))
                    analysisText.AppendLine($"   • {monthName}: {count} სესია - {total:N2} ₾")
                Catch
                    analysisText.AppendLine($"   • {monthKey}: {count} სესია - {total:N2} ₾")
                End Try
            Next

            analysisText.AppendLine()

            ' ჯამური სტატისტიკა
            analysisText.AppendLine("📈 ჯამური სტატისტიკა:")
            analysisText.AppendLine($"   • სულ სესიები: {DgvSessions.Rows.Count}")
            analysisText.AppendLine($"   • რეპორტში ჩართული: {GetIncludedSessionsCount()}")
            analysisText.AppendLine($"   • ჯამური შემოსავალი: {GetReportTotalAmount():N2} ₾")

            If GetIncludedSessionsCount() > 0 Then
                analysisText.AppendLine($"   • საშუალო შემოსავალი სესიაზე: {(GetReportTotalAmount() / GetIncludedSessionsCount()):N2} ₾")
            End If

            ' რეკომენდაციები
            analysisText.AppendLine()
            analysisText.AppendLine("💡 რეკომენდაციები:")

            If statusGroups.ContainsKey("შესრულებული") AndAlso statusGroups.ContainsKey("გაცდენა არასაპატიო") Then
                Dim completedCount = statusGroups("შესრულებული").Count
                Dim missedCount = statusGroups("გაცდენა არასაპატიო").Count
                Dim completionRate = If(GetIncludedSessionsCount() > 0, (completedCount * 100.0 / GetIncludedSessionsCount()).ToString("N1"), "0")

                analysisText.AppendLine($"   • შესრულების მაჩვენებელი: {completionRate}%")

                If completedCount < missedCount Then
                    analysisText.AppendLine("   • ⚠️  გაცდენების რაოდენობა მაღალია - კომუნიკაციის გაუმჯობესება საჭიროა")
                ElseIf completedCount > missedCount * 3 Then
                    analysisText.AppendLine("   • ✅ შესანიშნავი შესრულების მაჩვენებელი")
                End If
            End If

            Debug.WriteLine("UC_TherapistReport: დეტალური ანალიზი მომზადებულია")

            Return analysisText.ToString()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: GetDetailedReportAnalysis შეცდომა: {ex.Message}")
            Return $"დეტალური ანალიზის მიღების შეცდომა: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' ყველა სესიის მონიშვნა/განიშვნა (მასობრივი კონტროლი)
    ''' </summary>
    ''' <param name="selectAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
    Public Sub SetAllSessionsReportStatus(selectAll As Boolean)
        Try
            Debug.WriteLine($"UC_TherapistReport: ყველა სესიის რეპორტის სტატუსის დაყენება - {selectAll}")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Debug.WriteLine("UC_TherapistReport: მონაცემები არ არის ხელმისაწვდომი")
                Return
            End If

            Dim changedCount As Integer = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("IncludeInReport").Value = selectAll
                    UpdateRowVisualStyle(row.Index, selectAll)
                    changedCount += 1
                Catch
                    Continue For
                End Try
            Next

            UpdateReportTotals()

            Debug.WriteLine($"UC_TherapistReport: {changedCount} სესიის სტატუსი შეიცვალა")

            Dim actionText = If(selectAll, "ჩართულია", "ამოღებულია")
            MessageBox.Show($"{changedCount} სესია {actionText} რეპორტში", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: SetAllSessionsReportStatus შეცდომა: {ex.Message}")
            MessageBox.Show($"სესიების მასობრივი კონტროლის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის მიხედვით სესიების ფილტრაცია და კონტროლი
    ''' </summary>
    ''' <param name="statusFilter">სტატუსის ფილტრი</param>
    ''' <param name="includeInReport">რეპორტში ჩართვა/ამოღება</param>
    Public Sub FilterSessionsByStatusAndToggle(statusFilter As String, includeInReport As Boolean)
        Try
            Debug.WriteLine($"UC_TherapistReport: სტატუსის მიხედვით ფილტრაცია - '{statusFilter}', ჩართვა: {includeInReport}")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("ფილტრაციისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If String.IsNullOrWhiteSpace(statusFilter) Then
                MessageBox.Show("სტატუსის ფილტრი არ უნდა იყოს ცარიელი", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim changedCount As Integer = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim rowStatus As String = If(row.Cells("Status").Value?.ToString(), "")

                    ' თუ სტატუსი ემთხვევა ფილტრს
                    If String.Equals(rowStatus.Trim(), statusFilter.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        row.Cells("IncludeInReport").Value = includeInReport
                        UpdateRowVisualStyle(row.Index, includeInReport)
                        changedCount += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            UpdateReportTotals()

            Debug.WriteLine($"UC_TherapistReport: სტატუსით '{statusFilter}' {changedCount} სესიის სტატუსი შეიცვალა")

            Dim actionText = If(includeInReport, "ჩართულია", "ამოღებულია")
            MessageBox.Show($"სტატუსით '{statusFilter}' {changedCount} სესია {actionText} რეპორტში", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: FilterSessionsByStatusAndToggle შეცდომა: {ex.Message}")
            MessageBox.Show($"სტატუსის ფილტრაციის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' რეპორტის კონფიგურაციის ექსპორტი (JSON ფორმატით)
    ''' </summary>
    ''' <returns>რეპორტის კონფიგურაციის JSON</returns>
    Public Function ExportReportConfiguration() As String
        Try
            Debug.WriteLine("UC_TherapistReport: რეპორტის კონფიგურაციის ექსპორტი")

            Dim config As New Dictionary(Of String, Object)()

            ' ძირითადი ინფორმაცია
            config("therapist") = GetCurrentTherapistInfo()
            config("dateFrom") = DtpDan.Value.ToString("yyyy-MM-dd")
            config("dateTo") = DtpMde.Value.ToString("yyyy-MM-dd")
            config("therapyFilter") = GetSelectedTherapyType()
            config("fundingFilter") = GetSelectedFunding()
            config("exportDate") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' სტატისტიკა
            config("totalSessions") = If(DgvSessions?.Rows.Count, 0)
            config("includedSessions") = GetIncludedSessionsCount()
            config("totalAmount") = GetReportTotalAmount()

            ' ჩართული სესიების ID-ები
            Dim includedSessionIds As New List(Of Integer)()
            If DgvSessions IsNot Nothing Then
                For Each row As DataGridViewRow In DgvSessions.Rows
                    Try
                        Dim isIncluded As Boolean = True
                        If row.Cells("IncludeInReport").Value IsNot Nothing Then
                            Boolean.TryParse(row.Cells("IncludeInReport").Value.ToString(), isIncluded)
                        End If

                        If isIncluded AndAlso row.Tag IsNot Nothing Then
                            Dim sessionId As Integer
                            If Integer.TryParse(row.Tag.ToString(), sessionId) Then
                                includedSessionIds.Add(sessionId)
                            End If
                        End If
                    Catch
                        Continue For
                    End Try
                Next
            End If

            config("includedSessionIds") = includedSessionIds.ToArray()

            ' JSON-ის მარტივი სერიალიზაცია (System.Web.Script.Serialization.JavaScriptSerializer-ის გარეშე)
            Dim json As New System.Text.StringBuilder()
            json.AppendLine("{")
            json.AppendLine($"  ""therapist"": ""{EscapeJson(config("therapist").ToString())}"",")
            json.AppendLine($"  ""dateFrom"": ""{config("dateFrom")}"",")
            json.AppendLine($"  ""dateTo"": ""{config("dateTo")}"",")
            json.AppendLine($"  ""therapyFilter"": ""{EscapeJson(config("therapyFilter").ToString())}"",")
            json.AppendLine($"  ""fundingFilter"": ""{EscapeJson(config("fundingFilter").ToString())}"",")
            json.AppendLine($"  ""exportDate"": ""{config("exportDate")}"",")
            json.AppendLine($"  ""totalSessions"": {config("totalSessions")},")
            json.AppendLine($"  ""includedSessions"": {config("includedSessions")},")
            json.AppendLine($"  ""totalAmount"": {config("totalAmount")},")
            json.Append($"  ""includedSessionIds"": [{String.Join(",", includedSessionIds)}]")
            json.AppendLine()
            json.AppendLine("}")

            Debug.WriteLine("UC_TherapistReport: რეპორტის კონფიგურაცია ექსპორტირებულია")

            Return json.ToString()

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport: ExportReportConfiguration შეცდომა: {ex.Message}")
            Return $"{{""error"": ""{EscapeJson(ex.Message)}""}}"
        End Try
    End Function

    ''' <summary>
    ''' JSON სტრინგის escape
    ''' </summary>
    Private Function EscapeJson(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""

        text = text.Replace("\", "\\")
        text = text.Replace("""", "\""")
        text = text.Replace(vbCrLf, "\n")
        text = text.Replace(vbCr, "\n")
        text = text.Replace(vbLf, "\n")
        text = text.Replace(vbTab, "\t")

        Return text
    End Function

#End Region

#Region "რესურსების განთავისუფლება"

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            currentTherapistData?.Clear()
            dataProcessor?.ClearCache()

        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

End Class