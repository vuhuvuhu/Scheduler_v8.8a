' ===========================================
' 📄 UserControls/UC_BeneficiaryReport.vb
' -------------------------------------------
' ბენეფიციარის რეპორტის UserControl - UC_Schedule.vb-ის საფუძველზე გადაწერილი
' მარტივი და გამართული არქიტექტურა - არა გართულებული UI გაყოფით
' ძირითადი ფუნქცია: ბენეფიციარის არჩევა და მისი სესიების ჩვენება
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_BeneficiaryReport
    Inherits UserControl

#Region "ველები და თვისებები"

    ' სერვისები - UC_Schedule-ის ანალოგია
    Private dataService As IDataService = Nothing
    Private dataProcessor As ScheduleDataProcessor = Nothing
    Private uiManager As ScheduleUIManager = Nothing
    Private filterManager As ScheduleFilterManager = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6

    ' გვერდების მონაცემები
    Private currentPage As Integer = 1

    ' 🔧 ციკლური ივენთების თავიდან აცილება
    Private isNavigating As Boolean = False
    Private isLoadingData As Boolean = False

    ' 🆕 ბენეფიციარის სპეციფიკური ველები
    Private currentBeneficiaryData As List(Of SessionModel) = Nothing
    Private isUpdatingBeneficiary As Boolean = False

#End Region

#Region "საჯარო მეთოდები"

    ''' <summary>
    ''' მონაცემთა სერვისის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    Public Sub SetDataService(service As IDataService)
        Try
            Debug.WriteLine("UC_BeneficiaryReport: SetDataService")

            dataService = service

            ' სერვისების ინიციალიზაცია
            If dataService IsNot Nothing Then
                InitializeServices()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetDataService შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის დაყენება
    ''' </summary>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub SetUserInfo(email As String, role As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: SetUserInfo - email='{email}', role='{role}'")

            userEmail = email

            ' როლის პარსინგი
            Dim parsedRole As Integer
            If Integer.TryParse(role, parsedRole) Then
                userRoleID = parsedRole
            Else
                userRoleID = 6 ' ნაგულისხმევი
            End If

            Debug.WriteLine($"UC_BeneficiaryReport: საბოლოო userRoleID: {userRoleID}, userEmail: '{userEmail}'")

            ' UI-ს განახლება ახალი როლისთვის
            If uiManager IsNot Nothing Then
                LoadBeneficiarySpecificData()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetUserInfo შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების განახლება
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: მონაცემების განახლება")

            ' ქეშის გასუფთავება
            If TypeOf dataService Is SheetDataService Then
                DirectCast(dataService, SheetDataService).InvalidateAllCache()
            End If

            dataProcessor?.ClearCache()

            ' ComboBox-ების განახლება
            RefreshBeneficiaryComboBoxes()

            ' მონაცემების ხელახალი ჩატვირთვა
            LoadBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: RefreshData შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "პირადი მეთოდები"

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_BeneficiaryReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_BeneficiaryReport: კონტროლის ჩატვირთვა")

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვსრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadBeneficiarySpecificData()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სერვისების ინიციალიზაცია - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub InitializeServices()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: სერვისების ინიციალიზაცია")

            ' მონაცემების დამუშავების სერვისი
            dataProcessor = New ScheduleDataProcessor(dataService)

            ' UI მართვის სერვისი - ბენეფიციარის რეპორტისთვის მორგებული
            uiManager = New ScheduleUIManager(DgvSessions, LPage, BtnPrev, BtnNext)
            ConfigureBeneficiaryReportUI()

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

            ' ბენეფიციარის ComboBox-ების შევსება
            RefreshBeneficiaryComboBoxes()

            Debug.WriteLine("UC_BeneficiaryReport: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის რეპორტისთვის UI-ს კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureBeneficiaryReportUI()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის UI კონფიგურაცია")

            ' DataGridView-ის კონფიგურაცია
            uiManager.ConfigureDataGridView()
            uiManager.ConfigureNavigationButtons()

            ' ბენეფიციარის რეპორტისთვის სპეციალური სვეტები
            ConfigureBeneficiaryColumns()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConfigureBeneficiaryReportUI შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის რეპორტისთვის სვეტების კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureBeneficiaryColumns()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სვეტების კონფიგურაცია")

            If DgvSessions Is Nothing Then Return

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' ბენეფიციარის რეპორტისთვის სპეციალური სვეტები
                .Add("N", "N")                    ' ID
                .Add("DateTime", "თარიღი")        ' თარიღი
                .Add("Duration", "ხანგძლ.")       ' ხანგძლივობა
                .Add("TherapyType", "თერაპია")    ' თერაპიის ტიპი
                .Add("Therapist", "თერაპევტი")   ' თერაპევტი
                .Add("Status", "სტატუსი")        ' სტატუსი
                .Add("Price", "თანხა")           ' თანხა

                ' 🆕 ინვოისში ჩართვის CheckBox
                Dim includeColumn As New DataGridViewCheckBoxColumn()
                includeColumn.Name = "IncludeInInvoice"
                includeColumn.HeaderText = "ინვოისში"
                includeColumn.Width = 70
                includeColumn.TrueValue = True
                includeColumn.FalseValue = False
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

            ' სვეტების სიგანეების დაყენება
            SetBeneficiaryColumnWidths()

            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სვეტები კონფიგურირებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConfigureBeneficiaryColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის სვეტების სიგანეების დაყენება
    ''' </summary>
    Private Sub SetBeneficiaryColumnWidths()
        Try
            With DgvSessions.Columns
                .Item("N").Width = 50
                .Item("DateTime").Width = 130
                .Item("Duration").Width = 50
                .Item("TherapyType").Width = 200
                .Item("Therapist").Width = 180
                .Item("Status").Width = 130
                .Item("Price").Width = 80
                .Item("IncludeInInvoice").Width = 70
                .Item("Edit").Width = 35

                ' ფორმატირება
                .Item("Price").DefaultCellStyle.Format = "N2"
                .Item("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Item("N").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("IncludeInInvoice").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetBeneficiaryColumnWidths შეცდომა: {ex.Message}")
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

                ' ნაგულისხმევი მონიშვნა - ბენეფიციარის რეპორტისთვის რელევანტური სტატუსები
                filterManager.SetAllStatusCheckBoxes(False)
                filterManager.SetStatusCheckBox("შესრულებული", True)
                filterManager.SetStatusCheckBox("გაცდენა არასაპატიო", True)
                filterManager.SetStatusCheckBox("აღდგენა", True)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: InitializeStatusCheckBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' CheckBox-ის რეკურსიული ძებნა
    ''' </summary>
    Private Function FindCheckBoxRecursive(parent As Control, name As String) As CheckBox
        Try
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is CheckBox AndAlso ctrl.Name = name Then
                    Return DirectCast(ctrl, CheckBox)
                End If
            Next

            For Each ctrl As Control In parent.Controls
                Dim found = FindCheckBoxRecursive(ctrl, name)
                If found IsNot Nothing Then
                    Return found
                End If
            Next

            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' ივენთების მიბმა
    ''' </summary>
    Private Sub BindEvents()
        Try
            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged

            ' DataGridView-ის ივენთები
            AddHandler DgvSessions.CellClick, AddressOf OnDataGridViewCellClick
            AddHandler DgvSessions.CellValueChanged, AddressOf OnCheckBoxChanged

            ' ბენეფიციარის ComboBox-ების ივენთები
            AddHandler CBBeneName.SelectedIndexChanged, AddressOf OnBeneficiaryNameChanged
            AddHandler CBBeneSurname.SelectedIndexChanged, AddressOf OnBeneficiarySurnameChanged

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დაყენება
    ''' </summary>
    Private Sub SetBackgroundColors()
        Try
            Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

            If pnlFilter IsNot Nothing Then
                pnlFilter.BackColor = transparentWhite
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetBackgroundColors შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის ComboBox-ების განახლება
    ''' </summary>
    Private Sub RefreshBeneficiaryComboBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის ComboBox-ების განახლება")

            If dataProcessor Is Nothing Then Return

            isUpdatingBeneficiary = True

            Try
                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()

                ' ყველა ფილტრირებული მონაცემის მიღება
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                ' ბენეფიციარის სახელების განახლება
                PopulateBeneficiaryNamesComboBox(allSessions)

                ' გვარების ComboBox-ის გასუფთავება
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Enabled = False

            Finally
                isUpdatingBeneficiary = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: RefreshBeneficiaryComboBoxes შეცდომა: {ex.Message}")
            isUpdatingBeneficiary = False
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის სახელების ComboBox-ის შევსება
    ''' </summary>
    Private Sub PopulateBeneficiaryNamesComboBox(sessions As List(Of SessionModel))
        Try
            CBBeneName.Items.Clear()

            If sessions Is Nothing OrElse sessions.Count = 0 Then
                CBBeneName.Enabled = False
                Return
            End If

            ' უნიკალური სახელების მიღება
            Dim uniqueNames = sessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                    .Select(Function(s) s.BeneficiaryName.Trim()) _
                                    .Distinct() _
                                    .OrderBy(Function(nameItem) nameItem) _
                                    .ToList()

            For Each nameItem In uniqueNames
                CBBeneName.Items.Add(nameItem)
            Next

            CBBeneName.Enabled = (uniqueNames.Count > 0)

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {uniqueNames.Count} ბენეფიციარის სახელი")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის გვარების ComboBox-ის შევსება
    ''' </summary>
    Private Sub PopulateBeneficiarySurnamesComboBox(selectedName As String, sessions As List(Of SessionModel))
        Try
            CBBeneSurname.Items.Clear()

            If String.IsNullOrEmpty(selectedName) OrElse sessions Is Nothing Then
                CBBeneSurname.Enabled = False
                Return
            End If

            ' არჩეული სახელისთვის უნიკალური გვარების მიღება
            Dim uniqueSurnames = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                       Not String.IsNullOrEmpty(s.BeneficiarySurname)) _
                                          .Select(Function(s) s.BeneficiarySurname.Trim()) _
                                          .Distinct() _
                                          .OrderBy(Function(surname) surname) _
                                          .ToList()

            For Each surname In uniqueSurnames
                CBBeneSurname.Items.Add(surname)
            Next

            CBBeneSurname.Enabled = (uniqueSurnames.Count > 0)

            ' თუ მხოლოდ ერთი გვარია, ავტომატურად არჩევა
            If uniqueSurnames.Count = 1 Then
                CBBeneSurname.SelectedIndex = 0
            End If

            Debug.WriteLine($"UC_BeneficiaryReport: სახელისთვის '{selectedName}' ნაპოვნია {uniqueSurnames.Count} გვარი")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiarySurnamesComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' IList(Of IList(Of Object))-ის SessionModel-ებად გარდაქმნა
    ''' </summary>
    Private Function ConvertToSessionModels(data As List(Of IList(Of Object))) As List(Of SessionModel)
        Try
            Dim sessions As New List(Of SessionModel)()

            If data IsNot Nothing Then
                For Each row In data
                    Try
                        If row.Count >= 12 Then
                            Dim session = SessionModel.FromSheetRow(row)
                            sessions.Add(session)
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"UC_BeneficiaryReport: SessionModel-ის შექმნის შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next
            End If

            Return sessions

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConvertToSessionModels შეცდომა: {ex.Message}")
            Return New List(Of SessionModel)()
        End Try
    End Function

    ''' <summary>
    ''' 🆕 ბენეფიციარის სპეციფიკური მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadBeneficiarySpecificData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის მონაცემების ჩატვირთვა")

            If dataProcessor Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_BeneficiaryReport: სერვისები არ არის ინიციალიზებული")
                Return
            End If

            If isLoadingData Then
                Debug.WriteLine("UC_BeneficiaryReport: მონაცემები უკვე იტვირთება")
                Return
            End If

            isLoadingData = True

            Try
                ' არჩეული ბენეფიციარის შემოწმება
                Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
                Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

                If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) Then
                    ' ბენეფიციარი არ არის არჩეული - ცხრილის გასუფთავება
                    DgvSessions.Rows.Clear()
                    uiManager.UpdatePageLabel(1, 1)
                    uiManager.UpdateNavigationButtons(1, 1)
                    Return
                End If

                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()

                ' ყველა ფილტრირებული მონაცემის მიღება
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                ' კონკრეტული ბენეფიციარის სესიების ფილტრაცია
                currentBeneficiaryData = allSessions.Where(Function(s)
                                                               Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                      s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase)
                                                           End Function) _
                                                    .OrderBy(Function(s) s.DateTime) _
                                                    .ToList()

                Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარი '{selectedName} {selectedSurname}'-ისთვის ნაპოვნია {currentBeneficiaryData.Count} სესია")

                ' მონაცემების ჩატვირთვა DataGridView-ში
                LoadBeneficiarySessionsToGrid()

                ' ნავიგაციის განახლება (ერთი გვერდი ყველა სესიით)
                uiManager.UpdatePageLabel(1, 1)
                uiManager.UpdateNavigationButtons(1, 1)

            Finally
                isLoadingData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiarySpecificData შეცდომა: {ex.Message}")
            isLoadingData = False
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის სესიების ჩატვირთვა DataGridView-ში
    ''' </summary>
    Private Sub LoadBeneficiarySessionsToGrid()
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარის სესიების ჩატვირთვა - {currentBeneficiaryData.Count} სესია")

            DgvSessions.Rows.Clear()

            If currentBeneficiaryData Is Nothing OrElse currentBeneficiaryData.Count = 0 Then
                Return
            End If

            For i As Integer = 0 To currentBeneficiaryData.Count - 1
                Dim session = currentBeneficiaryData(i)

                ' მწკრივის მონაცემები
                Dim rowData As Object() = {
                    session.Id,                                     ' N
                    session.DateTime.ToString("dd.MM.yyyy HH:mm"), ' თარიღი
                    $"{session.Duration}წთ",                        ' ხანგძლივობა
                    session.TherapyType,                            ' თერაპია
                    session.TherapistName,                          ' თერაპევტი
                    session.Status,                                 ' სტატუსი
                    session.Price,                                  ' თანხა
                    True,                                           ' ნაგულისხმევად ინვოისში ჩართული
                    "✎"                                             ' რედაქტირების ღილაკი
                }

                Dim addedRowIndex = DgvSessions.Rows.Add(rowData)

                ' სესიის ID-ის შენახვა მწკრივში
                DgvSessions.Rows(addedRowIndex).Tag = session.Id

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                Try
                    Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = statusColor
                Catch
                    ' სტატუსის ფერის შეცდომის შემთხვევაში
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {currentBeneficiaryData.Count} სესია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiarySessionsToGrid შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები"

    ''' <summary>
    ''' ფილტრის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: OnFilterChanged")

            If isLoadingData OrElse isUpdatingBeneficiary Then
                Return
            End If

            ' ბენეფიციარის ComboBox-ების განახლება
            RefreshBeneficiaryComboBoxes()

            ' მონაცემების განახლება
            LoadBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის სახელის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnBeneficiaryNameChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary Then Return

            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სახელი შეიცვალა")

            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()

            If Not String.IsNullOrEmpty(selectedName) Then
                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                ' გვარების ComboBox-ის განახლება
                PopulateBeneficiarySurnamesComboBox(selectedName, allSessions)
            Else
                ' სახელი არ არის არჩეული - გვარების ComboBox გასუფთავება
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Enabled = False
            End If

            ' მონაცემების განახლება
            LoadBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnBeneficiaryNameChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის გვარის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnBeneficiarySurnameChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary Then Return

            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის გვარი შეიცვალა")

            ' მონაცემების განახლება
            LoadBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnBeneficiarySurnameChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის უჯრაზე დაჭერის ივენთი
    ''' </summary>
    Private Sub OnDataGridViewCellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' რედაქტირების ღილაკზე დაჭერის შემოწმება
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
               DgvSessions.Columns(e.ColumnIndex).Name = "Edit" Then

                Dim sessionId As Integer = 0
                If DgvSessions.Rows(e.RowIndex).Tag IsNot Nothing Then
                    Integer.TryParse(DgvSessions.Rows(e.RowIndex).Tag.ToString(), sessionId)
                End If

                If sessionId > 0 Then
                    Debug.WriteLine($"UC_BeneficiaryReport: რედაქტირება - სესია ID={sessionId}")

                    Try
                        Using editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_BeneficiaryReport")
                            Dim result As DialogResult = editForm.ShowDialog()

                            If result = DialogResult.OK Then
                                RefreshData()
                                MessageBox.Show($"სესია ID={sessionId} წარმატებით განახლდა", "წარმატება",
                                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End Using

                    Catch formEx As Exception
                        Debug.WriteLine($"UC_BeneficiaryReport: რედაქტირების ფორმის შეცდომა: {formEx.Message}")
                        MessageBox.Show($"რედაქტირების ფორმის გახსნის შეცდომა: {formEx.Message}", "შეცდომა",
                                       MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnDataGridViewCellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 CheckBox-ის შეცვლის ივენთი (ინვოისში ჩართვა/ამოღება)
    ''' </summary>
    Private Sub OnCheckBoxChanged(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
               DgvSessions.Columns(e.ColumnIndex).Name = "IncludeInInvoice" Then

                Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის CheckBox შეიცვალა - მწკრივი {e.RowIndex}")

                ' ინვოისის ჯამების განახლება (მომავალშ თუ საჭირო იქნება)
                UpdateInvoiceTotals()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnCheckBoxChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკი - ბენეფიციარის რეპორტში არ გამოიყენება
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs) Handles BtnPrev.Click
        ' ბენეფიციარის რეპორტში ყველა სესია ერთ გვერდზე ჩანს
        Debug.WriteLine("UC_BeneficiaryReport: წინა გვერდი - არ გამოიყენება")
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკი - ბენეფიციარის რეპორტში არ გამოიყენება
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        ' ბენეფიციარის რეპორტში ყველა სესია ერთ გვერდზე ჩანს
        Debug.WriteLine("UC_BeneficiaryReport: შემდეგი გვერდი - არ გამოიყენება")
    End Sub

    ''' <summary>
    ''' განახლების ღილაკი
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            RefreshData()
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BtnRef_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი ჩანაწერის დამატების ღილაკი
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ახალი სესიის დამატება")

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
            Using newRecordForm As New NewRecordForm(dataService, "სესია", userEmail, "UC_BeneficiaryReport")
                Dim result = newRecordForm.ShowDialog()

                If result = DialogResult.OK Then
                    Debug.WriteLine("UC_BeneficiaryReport: სესია წარმატებით დაემატა")
                    RefreshData()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BtnAddSchedule_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🖨️ ბეჭდვის ღილაკი - ბენეფიციარის რეპორტისთვის
    ''' </summary>
    Private Sub BtbPrint_Click(sender As Object, e As EventArgs) Handles btbPrint.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბეჭდვის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ბენეფიციარის ინფორმაციის შემოწმება
            Dim beneficiaryInfo = GetCurrentBeneficiaryInfo()
            If String.IsNullOrEmpty(beneficiaryInfo) Then
                MessageBox.Show("აირჩიეთ ბენეფიციარი ბეჭდვისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ბეჭდვის ტიპის არჩევა
            Dim printTypeResult As DialogResult = MessageBox.Show(
                "რომელი ტიპის ბეჭდვა გსურთ?" & Environment.NewLine & Environment.NewLine &
                "დიახ - ბენეფიციარის ინვოისი" & Environment.NewLine &
                "არა - ჩვეულებრივი ცხრილის ბეჭდვა" & Environment.NewLine &
                "გაუქმება - ოპერაციის შეწყვეტა",
                "ბეჭდვის ტიპის არჩევა",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case printTypeResult
                Case DialogResult.Yes
                    ' ბენეფიციარის ინვოისის ბეჭდვა
                    PrintBeneficiaryInvoice()

                Case DialogResult.No
                    ' ჩვეულებრივი ცხრილის ბეჭდვა
                    Using printService As New AdvancedDataGridViewPrintService(DgvSessions)
                        printService.ShowFullPrintDialog()
                    End Using

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_BeneficiaryReport: ბეჭდვა გაუქმებულია")
            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BtbPrint_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 📄 PDF ექსპორტის ღილაკი - ბენეფიციარის რეპორტისთვის
    ''' </summary>
    Private Sub btnToPDF_Click(sender As Object, e As EventArgs) Handles btnToPDF.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: PDF ექსპორტის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("PDF ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ბენეფიციარის ინფორმაციის შემოწმება
            Dim beneficiaryInfo = GetCurrentBeneficiaryInfo()
            If String.IsNullOrEmpty(beneficiaryInfo) Then
                MessageBox.Show("აირჩიეთ ბენეფიციარი PDF ექსპორტისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' PDF ექსპორტის ტიპის არჩევა
            Dim pdfTypeResult As DialogResult = MessageBox.Show(
                "რომელი ტიპის PDF ექსპორტი გსურთ?" & Environment.NewLine & Environment.NewLine &
                "დიახ - ბენეფიციარის ინვოისი (PDF)" & Environment.NewLine &
                "არა - ცხრილის PDF ექსპორტი" & Environment.NewLine &
                "გაუქმება - ოპერაციის შეწყვეტა",
                "PDF ექსპორტის ტიპის არჩევა",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case pdfTypeResult
                Case DialogResult.Yes
                    ' ბენეფიციარის ინვოისის PDF ექსპორტი
                    ExportBeneficiaryInvoiceToPDF()

                Case DialogResult.No
                    ' ჩვეულებრივი ცხრილის PDF ექსპორტი
                    Using exportService As New SimplePDFExportService(DgvSessions)
                        exportService.ShowFullExportDialog()
                    End Using

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_BeneficiaryReport: PDF ექსპორტი გაუქმებულია")
            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: btnToPDF_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "🆕 ბენეფიციარის სპეციფიკური მეთოდები"

    ''' <summary>
    ''' 🆕 მიმდინარე ბენეფიციარის ინფორმაციის მიღება
    ''' </summary>
    ''' <returns>ბენეფიციარის სრული სახელი</returns>
    Public Function GetCurrentBeneficiaryInfo() As String
        Try
            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If Not String.IsNullOrEmpty(selectedName) AndAlso Not String.IsNullOrEmpty(selectedSurname) Then
                Return $"{selectedName} {selectedSurname}"
            End If

            Return ""

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetCurrentBeneficiaryInfo შეცდომა: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 🆕 ინვოისის ჯამური თანხის გამოთვლა (მხოლოდ მონიშნული სესიები)
    ''' </summary>
    ''' <returns>ინვოისის ჯამური თანხა</returns>
    Public Function GetInvoiceTotalAmount() As Decimal
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim total As Decimal = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' შევამოწმოთ ინვოისში ჩართვის CheckBox
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInInvoice").Value.ToString(), isIncluded)
                    End If

                    ' თუ სესია ჩართულია ინვოისში
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

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის ჯამური თანხა: {total:N2}")
            Return total

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetInvoiceTotalAmount შეცდომა: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 🆕 ინვოისში ჩართული სესიების რაოდენობა
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
                    If row.Cells("IncludeInInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInInvoice").Value.ToString(), isIncluded)
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
    ''' 🆕 ინვოისის ღირებულებების განახლება
    ''' </summary>
    Private Sub UpdateInvoiceTotals()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის ღირებულებების განახლება")

            Dim totalAmount = GetInvoiceTotalAmount()
            Dim includedCount = GetIncludedSessionsCount()

            Debug.WriteLine($"UC_BeneficiaryReport: ჯამური თანხა: {totalAmount:N2}, ჩართული სესიები: {includedCount}")

            ' აქ შეიძლება დავამატოთ UI ელემენტების განახლება თანხისა და რაოდენობის ჩვენებისთვის

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateInvoiceTotals შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ყველა სესიის ინვოისში ჩართვა
    ''' </summary>
    Public Sub IncludeAllInInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სესიის ინვოისში ჩართვა")

            If DgvSessions Is Nothing Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("IncludeInInvoice").Value = True
                Catch
                    Continue For
                End Try
            Next

            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: IncludeAllInInvoice შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ყველა სესიის ინვოისიდან ამოღება
    ''' </summary>
    Public Sub ExcludeAllFromInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სესიის ინვოისიდან ამოღება")

            If DgvSessions Is Nothing Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("IncludeInInvoice").Value = False
                Catch
                    Continue For
                End Try
            Next

            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ExcludeAllFromInvoice შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის ინვოისის ბეჭდვა
    ''' </summary>
    Private Sub PrintBeneficiaryInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის ინვოისის ბეჭდვა")

            ' ინვოის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("ინვოისის ბეჭდვისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' HTML ინვოისის შექმნა და ბეჭდვა
            CreateInvoiceHTML(True) ' True = ბეჭდვისთვის

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PrintBeneficiaryInvoice შეცდომა: {ex.Message}")
            MessageBox.Show($"ინვოისის ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის ინვოისის PDF ექსპორტი
    ''' </summary>
    Private Sub ExportBeneficiaryInvoiceToPDF()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის ინვოისის PDF ექსპორტი")

            ' ინვოისის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("ინვოისის ექსპორტისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ფაილის ადგილის არჩევა
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "HTML ფაილები (*.html)|*.html"
                saveDialog.Title = "ინვოისის შენახვა"

                Dim beneficiaryName = GetCurrentBeneficiaryInfo().Replace(" ", "_")
                saveDialog.FileName = $"ინვოისი_{beneficiaryName}_{DateTime.Now:yyyyMMdd}.html"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    CreateInvoiceHTML(False, saveDialog.FileName) ' False = ფაილისთვის
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ExportBeneficiaryInvoiceToPDF შეცდომა: {ex.Message}")
            MessageBox.Show($"ინვოისის PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის HTML-ის შექმნა
    ''' </summary>
    ''' <param name="forPrinting">True თუ ბეჭდვისთვის, False თუ ფაილისთვის</param>
    ''' <param name="filePath">ფაილის მისამართი (ფაილისთვის)</param>
    Private Sub CreateInvoiceHTML(forPrinting As Boolean, Optional filePath As String = "")
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის HTML შექმნა - ბეჭდვისთვის: {forPrinting}")

            Dim html As New System.Text.StringBuilder()
            Dim beneficiaryName = GetCurrentBeneficiaryInfo()
            Dim period = $"{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"

            ' HTML დოკუმენტის შექმნა
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html lang=""ka"">")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>ინვოისი - " & beneficiaryName & "</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        @page { size: A4; margin: 20mm; }")
            html.AppendLine("        @media print { .no-print { display: none; } }")
            html.AppendLine("        body { font-family: 'Sylfaen', Arial, sans-serif; font-size: 12px; line-height: 1.4; }")
            html.AppendLine("        .invoice-header { text-align: center; margin-bottom: 30px; }")
            html.AppendLine("        .invoice-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }")
            html.AppendLine("        .company-info { font-size: 11px; margin-bottom: 20px; }")
            html.AppendLine("        .beneficiary-info { border: 1px solid #333; padding: 15px; margin: 20px 0; background: #f9f9f9; }")
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
                html.AppendLine("            🖨️ ინვოისის ბეჭდვა PDF-ად</button>")
                html.AppendLine("        <p>ღილაკზე დაჭერის შემდეგ აირჩიეთ ""Microsoft Print to PDF""</p>")
                html.AppendLine("    </div>")
            End If

            ' ინვოისის სათაური
            html.AppendLine("    <div class=""invoice-header"">")
            html.AppendLine("        <div class=""invoice-title"">ინვოისი მომსახურების გაწევაზე</div>")
            html.AppendLine("        <div class=""company-info"">")
            html.AppendLine("            შპს ""ბავშვთა და მოზარდთა განვითარების, აბილიტაციისა და რეაბილიტაციის ცენტრი - პროსპერო""<br>")
            html.AppendLine("            მისამართი: [კომპანიის მისამართი]<br>")
            html.AppendLine("            ტელ: [ტელეფონი] | ელ-ფოსტა: [ელფოსტა]")
            html.AppendLine("        </div>")
            html.AppendLine("    </div>")

            ' ბენეფიციარის ინფორმაცია
            html.AppendLine("    <div class=""beneficiary-info"">")
            html.AppendLine($"        <strong>ბენეფიციარი:</strong> {EscapeHtml(beneficiaryName)}<br>")
            html.AppendLine("        <strong>დაბადების თარიღი:</strong> _______________<br>")
            html.AppendLine("        <strong>კანონიერი წარმომადგენელი:</strong> _______________<br>")
            html.AppendLine("        <strong>წარმომადგენლის პ/ნ:</strong> _______________<br>")
            html.AppendLine($"        <strong>პერიოდი:</strong> {period}")
            html.AppendLine("    </div>")

            ' მომსახურებების ცხრილი
            html.AppendLine("    <table>")
            html.AppendLine("        <thead>")
            html.AppendLine("            <tr>")
            html.AppendLine("                <th style=""width: 30px;"">N</th>")
            html.AppendLine("                <th style=""width: 100px;"">თარიღი</th>")
            html.AppendLine("                <th style=""width: 200px;"">მომსახურების სახე</th>")
            html.AppendLine("                <th style=""width: 150px;"">თერაპევტი</th>")
            html.AppendLine("                <th style=""width: 60px;"">ხანგძლ.</th>")
            html.AppendLine("                <th style=""width: 80px;"">თანხა (₾)</th>")
            html.AppendLine("            </tr>")
            html.AppendLine("        </thead>")
            html.AppendLine("        <tbody>")

            ' მომსახურებების სია (მხოლოდ ინვოისში ჩართული)
            Dim invoiceNumber As Integer = 1
            Dim totalAmount As Decimal = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' შევამოწმოთ ინვოისში ჩართვა
                    Dim isIncluded As Boolean = True
                    If row.Cells("IncludeInInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInInvoice").Value.ToString(), isIncluded)
                    End If

                    If isIncluded Then
                        Dim dateTime As String = If(row.Cells("DateTime").Value?.ToString(), "")
                        Dim therapyType As String = If(row.Cells("TherapyType").Value?.ToString(), "")
                        Dim therapist As String = If(row.Cells("Therapist").Value?.ToString(), "")
                        Dim duration As String = If(row.Cells("Duration").Value?.ToString(), "")

                        Dim price As Decimal = 0
                        If row.Cells("Price").Value IsNot Nothing Then
                            Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                        End If

                        totalAmount += price

                        html.AppendLine("            <tr>")
                        html.AppendLine($"                <td style=""text-align: center;"">{invoiceNumber}</td>")
                        html.AppendLine($"                <td>{EscapeHtml(dateTime)}</td>")
                        html.AppendLine($"                <td>{EscapeHtml(therapyType)}</td>")
                        html.AppendLine($"                <td>{EscapeHtml(therapist)}</td>")
                        html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(duration)}</td>")
                        html.AppendLine($"                <td class=""amount"">{price:N2}</td>")
                        html.AppendLine("            </tr>")

                        invoiceNumber += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            html.AppendLine("        </tbody>")
            html.AppendLine("    </table>")

            ' ჯამური თანხა
            html.AppendLine("    <div class=""total-section"">")
            html.AppendLine($"        <div class=""total-amount"">მომსახურების საფასური სულ: {totalAmount:N2} ₾</div>")
            html.AppendLine($"        <div>თანხა სიტყვიერად: {ConvertAmountToWords(totalAmount)}</div>")
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
                finalFilePath = System.IO.Path.GetTempPath() & $"invoice_{DateTime.Now:yyyyMMddHHmmss}.html"
            Else
                ' მომხმარებლის მიერ არჩეული ფაილი
                finalFilePath = filePath
            End If

            ' HTML ფაილის ჩაწერა
            System.IO.File.WriteAllText(finalFilePath, html.ToString(), System.Text.Encoding.UTF8)

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის HTML შეიქმნა - {finalFilePath}")

            If forPrinting Then
                ' ბეჭდვისთვის - ფაილის გახსნა და ავტომატური ბეჭდვა
                System.Diagnostics.Process.Start(finalFilePath)
                MessageBox.Show("ინვოისი ბრაუზერში გაიხსნა ბეჭდვისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' ფაილისთვის - შეტყობინება და გახსნის შეთავაზება
                MessageBox.Show($"ინვოისი წარმატებით შეიქმნა:{Environment.NewLine}{finalFilePath}" & Environment.NewLine & Environment.NewLine &
                               "დააჭირეთ ფაილში 'ინვოისის ბეჭდვა PDF-ად' ღილაკს", "წარმატება",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)

                Dim openResult As DialogResult = MessageBox.Show("გსურთ ინვოისის ფაილის გახსნა?", "ფაილის გახსნა",
                                                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If openResult = DialogResult.Yes Then
                    System.Diagnostics.Process.Start(finalFilePath)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateInvoiceHTML შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის ვალიდობის შემოწმება
    ''' </summary>
    ''' <returns>True თუ ინვოისი ვალიდურია</returns>
    Private Function IsInvoiceValid() As Boolean
        Try
            Dim beneficiaryName = GetCurrentBeneficiaryInfo()
            Dim includedCount = GetIncludedSessionsCount()

            Return Not String.IsNullOrEmpty(beneficiaryName) AndAlso includedCount > 0

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: IsInvoiceValid შეცდომა: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 🆕 თანხის სიტყვიერ ფორმად გარდაქმნა (მარტივი ვერსია)
    ''' </summary>
    ''' <param name="amount">თანხა</param>
    ''' <returns>თანხა სიტყვიერად</returns>
    Private Function ConvertAmountToWords(amount As Decimal) As String
        Try
            If amount = 0 Then
                Return "ნული ლარი, ნული თეთრი"
            End If

            Dim lari As Integer = Math.Floor(amount)
            Dim tetri As Integer = Math.Round((amount - lari) * 100)

            Dim result As String = ""

            ' ლარის ნაწილი (მარტივი ვერსია)
            If lari = 0 Then
                result = "ნული ლარი"
            ElseIf lari = 1 Then
                result = "ერთი ლარი"
            Else
                result = $"{lari} ლარი"
            End If

            ' თეთრის ნაწილი
            If tetri = 0 Then
                result += ", ნული თეთრი"
            ElseIf tetri = 1 Then
                result += ", ერთი თეთრი"
            Else
                result += $", {tetri} თეთრი"
            End If

            Return result

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConvertAmountToWords შეცდომა: {ex.Message}")
            Return "თანხის გარდაქმნის შეცდომა"
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

#End Region

#Region "🆕 ინვოისის მართვის საჯარო მეთოდები"

    ''' <summary>
    ''' 🆕 მონიშნული სესიების ინვოისში ჩართვა/ამოღება
    ''' </summary>
    Public Sub ToggleSelectedSessionsInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: მონიშნული სესიების ინვოისში ჩართვა/ამოღება")

            If DgvSessions Is Nothing OrElse DgvSessions.SelectedRows.Count = 0 Then
                MessageBox.Show("აირჩიეთ სესიები ინვოისის კონტროლისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            For Each row As DataGridViewRow In DgvSessions.SelectedRows
                Try
                    ' მიმდინარე მდგომარეობის შებრუნება
                    Dim currentValue As Boolean = True
                    If row.Cells("IncludeInInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("IncludeInInvoice").Value.ToString(), currentValue)
                    End If

                    row.Cells("IncludeInInvoice").Value = Not currentValue

                Catch
                    Continue For
                End Try
            Next

            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ToggleSelectedSessionsInvoice შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის სტატისტიკის მიღება
    ''' </summary>
    ''' <returns>სტატისტიკის ტექსტი</returns>
    Public Function GetInvoiceStatistics() As String
        Try
            Dim beneficiaryName = GetCurrentBeneficiaryInfo()
            Dim totalSessions = If(DgvSessions?.Rows.Count, 0)
            Dim includedSessions = GetIncludedSessionsCount()
            Dim totalAmount = GetInvoiceTotalAmount()
            Dim period = $"{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"

            Return $"ბენეფიციარი: {beneficiaryName}" & Environment.NewLine &
                   $"პერიოდი: {period}" & Environment.NewLine &
                   $"სულ სესიები: {totalSessions}" & Environment.NewLine &
                   $"ინვოისში ჩართული: {includedSessions}" & Environment.NewLine &
                   $"ჯამური თანხა: {totalAmount:N2} ₾"

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetInvoiceStatistics შეცდომა: {ex.Message}")
            Return "სტატისტიკის მიღების შეცდომა"
        End Try
    End Function

#End Region

#Region "რესურსების განთავისუფლება"

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            currentBeneficiaryData?.Clear()
            dataProcessor?.ClearCache()

        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

End Class