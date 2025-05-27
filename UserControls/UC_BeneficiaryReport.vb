' ===========================================
' 📄 UserControls/UC_BeneficiaryReport.vb - სრულად თავიდან დაწერილი
' -------------------------------------------
' ბენეფიციარის ინვოისის ფორმის UserControl
' UC_Schedule-ის ფილტრაციის სისტემის გამოყენებით
' 
' 🎯 ხელმისაწვდომია როლებისთვის: 1 (ადმინი), 2 (მენეჯერი), 3
' ===========================================
Imports System.ComponentModel
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_BeneficiaryReport
    Inherits UserControl

#Region "ველები და თვისებები"

    ' სერვისები - UC_Schedule-ის ანალოგია
    Private dataService As IDataService = Nothing
    Private dataProcessor As ScheduleDataProcessor = Nothing
    Private filterManager As ScheduleFilterManager = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRole As String = ""

    ' ფილტრირებული მონაცემები
    Private currentFilteredSessions As List(Of SessionModel) = Nothing
    Private currentBeneficiaryData As List(Of SessionModel) = Nothing

    ' 🔧 ციკლური ივენთების თავიდან აცილება
    Private isLoadingData As Boolean = False
    Private isUpdatingComboBoxes As Boolean = False

#End Region

#Region "საჯარო მეთოდები - UC_Schedule-ის ანალოგია"

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
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
            userRole = role

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

            ' ფილტრების განახლება
            If filterManager IsNot Nothing AndAlso dataProcessor IsNot Nothing Then
                filterManager.PopulateFilterComboBoxes(dataProcessor)
            End If

            ' მონაცემების ხელახალი ჩატვირთვა
            LoadFilteredData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: RefreshData შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "პირადი მეთოდები - სერვისების ინიციალიზაცია"

    ''' <summary>
    ''' UserControl-ის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_BeneficiaryReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_BeneficiaryReport: კონტროლის ჩატვირთვა")

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვსრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadFilteredData()
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

            ' ფილტრების მართვის სერვისი - UC_Schedule-ის ანალოგია
            ' მხოლოდ საჭირო კონტროლები
            filterManager = New ScheduleFilterManager(
                DtpDan, DtpMde,
                CBBeneName, CBBeneSurname, CBPer, CBTer, Nothing, Nothing ' Space და Funding არ გვჭირდება ახლა
            )

            ' სტატუსის CheckBox-ების ინიციალიზაცია
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ივენთების მიბმა
            BindEvents()

            ' DataGridView-ის კონფიგურაცია
            ConfigureDataGridView()

            ' საწყისი ფილტრების დაყენება
            SetInitialFilters()

            Debug.WriteLine("UC_BeneficiaryReport: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის CheckBox-ების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: სტატუსის CheckBox-ების ინიციალიზაცია")

            Dim statusCheckBoxes As New List(Of CheckBox)

            ' CheckBox1-დან CheckBox7-მდე მოძიება
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBoxRecursive(Me, $"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    statusCheckBoxes.Add(checkBox)
                    Debug.WriteLine($"UC_BeneficiaryReport: ნაპოვნი CheckBox{i} - ტექსტი: '{checkBox.Text}'")
                End If
            Next

            If statusCheckBoxes.Count > 0 Then
                filterManager.InitializeStatusCheckBoxes(statusCheckBoxes.ToArray())

                ' ნაგულისხმევი მონიშვნა - მხოლოდ შესრულებული, გაცდენა არასაპატიო და აღდგენა
                filterManager.SetAllStatusCheckBoxes(False) ' ჯერ ყველა განვიშნოთ

                ' შემდეგ კონკრეტულები მონიშნული
                filterManager.SetStatusCheckBox("შესრულებული", True)
                filterManager.SetStatusCheckBox("გაცდენა არასაპატიო", True)
                filterManager.SetStatusCheckBox("აღდგენა", True)

                Debug.WriteLine("UC_BeneficiaryReport: სტატუსების ნაგულისხმევი მონიშვნა დასრულდა")
            Else
                Debug.WriteLine("UC_BeneficiaryReport: სტატუსის CheckBox-ები ვერ მოიძებნა")
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
            Debug.WriteLine("UC_BeneficiaryReport: ივენთების მიბმა")

            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის კონფიგურაცია ინვოისის ფორმისთვის
    ''' </summary>
    Private Sub ConfigureDataGridView()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: DataGridView-ის კონფიგურაცია")

            If DgvSessions Is Nothing Then
                Debug.WriteLine("UC_BeneficiaryReport: DgvSessions არ არსებობს")
                Return
            End If

            With DgvSessions
                .AutoGenerateColumns = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .ReadOnly = True
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = False

                ' თავის სტილი
                .EnableHeadersVisualStyles = False
                .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 80, 150)
                .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
                .ColumnHeadersDefaultCellStyle.Font = New Font("Sylfaen", 10, FontStyle.Bold)

                ' მწკრივების სტილი
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 250)
                .DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 200, 255)
                .DefaultCellStyle.SelectionForeColor = Color.Black
            End With

            ' სვეტების შექმნა ინვოისისთვის
            CreateInvoiceColumns()

            Debug.WriteLine("UC_BeneficiaryReport: DataGridView კონფიგურაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConfigureDataGridView შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ინვოისის სვეტების შექმნა - მხოლოდ საჭირო სვეტები
    ''' </summary>
    Private Sub CreateInvoiceColumns()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის სვეტების შექმნა")

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' ინვოისის ნომერი (არა ბაზის ID)
                .Add("InvoiceN", "N")

                ' თარიღი
                .Add("DateTime", "თარიღი")

                ' ხანგძლივობა
                .Add("Duration", "ხანგძლივობა")

                ' თერაპია
                .Add("TherapyType", "თერაპია")

                ' თერაპევტი
                .Add("Therapist", "თერაპევტი")

                ' თანხა
                .Add("Price", "თანხა")
            End With

            ' სვეტების სიგანეების დაყენება
            SetInvoiceColumnWidths()

            Debug.WriteLine($"UC_BeneficiaryReport: შეიქმნა {DgvSessions.Columns.Count} სვეტი ინვოისისთვის")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateInvoiceColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ინვოისის სვეტების სიგანეების დაყენება
    ''' </summary>
    Private Sub SetInvoiceColumnWidths()
        Try
            With DgvSessions.Columns
                .Item("InvoiceN").Width = 50          ' N
                .Item("DateTime").Width = 150         ' თარიღი
                .Item("Duration").Width = 80          ' ხანგძლივობა
                .Item("TherapyType").Width = 200      ' თერაპია
                .Item("Therapist").Width = 180        ' თერაპევტი
                .Item("Price").Width = 100            ' თანხა

                ' ფასის სვეტის ფორმატი
                .Item("Price").DefaultCellStyle.Format = "N2"
                .Item("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის სვეტების სიგანეები დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetInvoiceColumnWidths შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' საწყისი ფილტრების დაყენება
    ''' </summary>
    Private Sub SetInitialFilters()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: საწყისი ფილტრების დაყენება")

            ' თარიღების საწყისი მნიშვნელობები - მიმდინარე თვე
            DtpDan.Value = New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1) ' თვის პირველი დღე
            DtpMde.Value = DateTime.Today ' დღეს

            ' ComboBox-ების საწყისი მდგომარეობა
            CBBeneName.Enabled = False
            CBBeneSurname.Enabled = False
            CBPer.Enabled = False
            CBTer.Enabled = False

            CBBeneName.Items.Clear()
            CBBeneSurname.Items.Clear()
            CBPer.Items.Clear()
            CBTer.Items.Clear()

            ' DataGridView-ის გასუფთავება
            DgvSessions.Rows.Clear()

            Debug.WriteLine("UC_BeneficiaryReport: საწყისი ფილტრები დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetInitialFilters შეცდომა: {ex.Message}")
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

#End Region

#Region "ფილტრაციის ლოგიკა"

    ''' <summary>
    ''' ფილტრაციის შედეგების ჩატვირთვა - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub LoadFilteredData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: LoadFilteredData")

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
                ' ფილტრის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()

                ' ყველა ფილტრირებული მონაცემის მიღება (არა გვერდების მიხედვით)
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)

                ' SessionModel-ების შექმნა
                currentFilteredSessions = ConvertToSessionModels(result.Data)

                Debug.WriteLine($"UC_BeneficiaryReport: ფილტრირებულია {currentFilteredSessions.Count} სესია")

                ' 🔧 UC_Schedule-ის ანალოგია: ComboBox-ების განახლება
                UpdateComboBoxesIfNeeded()

                ' ბენეფიციარის სპეციფიკური მონაცემების ფილტრაცია
                FilterBeneficiarySpecificData()

            Finally
                isLoadingData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadFilteredData შეცდომა: {ex.Message}")
            isLoadingData = False
        End Try
    End Sub

    ''' <summary>
    ''' IList(Of IList(Of Object))-ის SessionModel-ებად გარდაქმნა
    ''' </summary>
    Private Function ConvertToSessionModels(data As List(Of IList(Of Object))) As List(Of SessionModel)
        Try
            Dim sessions As New List(Of SessionModel)()

            If data IsNot Nothing Then
                For i As Integer = 0 To data.Count - 1
                    Try
                        Dim row = data(i)
                        If row.Count >= 12 Then ' მინიმუმ საჭირო სვეტები
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
    ''' ComboBox-ების განახლება თუ საჭიროა - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub UpdateComboBoxesIfNeeded()
        Try
            If isUpdatingComboBoxes Then Return

            ' ბენეფიციარების სახელების ჩატვირთვა - UC_Schedule-ის მანერით
            PopulateBeneficiaryNamesComboBox()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateComboBoxesIfNeeded შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სახელების ComboBox-ის შევსება - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub PopulateBeneficiaryNamesComboBox()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox")

            If CBBeneName Is Nothing OrElse currentFilteredSessions Is Nothing Then Return

            ' 🔧 UC_Schedule-ის მანერით: შევინახოთ მიმდინარე არჩეული მნიშვნელობა
            Dim selectedValue As String = GetComboBoxSelectedValue(CBBeneName)

            CBBeneName.Items.Clear()

            If currentFilteredSessions.Count = 0 Then
                CBBeneName.Enabled = False
                Debug.WriteLine("UC_BeneficiaryReport: ფილტრირებული სესიები ცარიელია")
                Return
            End If

            ' უნიკალური სახელების მიღება
            Dim uniqueNames = currentFilteredSessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                                   .Select(Function(s) s.BeneficiaryName.Trim()) _
                                                   .Distinct() _
                                                   .OrderBy(Function(nameItem) nameItem) _
                                                   .ToList()

            ' სახელების დამატება
            For i As Integer = 0 To uniqueNames.Count - 1
                CBBeneName.Items.Add(uniqueNames(i))
            Next

            CBBeneName.Enabled = (uniqueNames.Count > 0)

            ' 🔧 UC_Schedule-ის მანერით: ვაბრუნებთ არჩეულ მნიშვნელობას
            SetComboBoxSelectedValue(CBBeneName, selectedValue)

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {uniqueNames.Count} სახელი, აღდგენილია: '{selectedValue}'")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ComboBox-ის მიმდინარე არჩეული მნიშვნელობის მიღება - UC_Schedule-დან
    ''' </summary>
    Private Function GetComboBoxSelectedValue(comboBox As ComboBox) As String
        Try
            If comboBox IsNot Nothing AndAlso comboBox.SelectedItem IsNot Nothing Then
                Return comboBox.SelectedItem.ToString()
            End If
            Return ""
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' ComboBox-ის არჩეული მნიშვნელობის დაყენება - UC_Schedule-დან
    ''' </summary>
    Private Sub SetComboBoxSelectedValue(comboBox As ComboBox, value As String)
        Try
            If comboBox Is Nothing OrElse String.IsNullOrEmpty(value) Then Return

            For i As Integer = 0 To comboBox.Items.Count - 1
                If String.Equals(comboBox.Items(i).ToString(), value, StringComparison.OrdinalIgnoreCase) Then
                    comboBox.SelectedIndex = i
                    Return
                End If
            Next

            ' თუ არ მოიძებნა, პირველი ელემენტი არ დავაყენოთ

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetComboBoxSelectedValue შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სახელების ჩატვირთვა
    ''' </summary>
    Private Sub LoadBeneficiaryNames()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: LoadBeneficiaryNames")

            isUpdatingComboBoxes = True

            Try
                CBBeneName.Items.Clear()
                CBBeneSurname.Items.Clear()
                CBPer.Items.Clear()
                CBTer.Items.Clear()

                CBBeneName.Text = ""
                CBBeneSurname.Text = ""
                CBPer.Text = ""
                CBTer.Text = ""

                If currentFilteredSessions Is Nothing OrElse currentFilteredSessions.Count = 0 Then
                    CBBeneName.Enabled = False
                    CBBeneSurname.Enabled = False
                    CBPer.Enabled = False
                    CBTer.Enabled = False
                    Debug.WriteLine("UC_BeneficiaryReport: ფილტრირებული სესიები ცარიელია")
                    Return
                End If

                ' უნიკალური სახელების მიღება
                Dim uniqueNames = currentFilteredSessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                                       .Select(Function(s) s.BeneficiaryName.Trim()) _
                                                       .Distinct() _
                                                       .OrderBy(Function(nameItem) nameItem) _
                                                       .ToList()

                ' სახელების დამატება
                For Each nameItem In uniqueNames
                    CBBeneName.Items.Add(nameItem)
                Next

                CBBeneName.Enabled = (uniqueNames.Count > 0)
                CBBeneSurname.Enabled = False
                CBPer.Enabled = False
                CBTer.Enabled = False

                Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {uniqueNames.Count} უნიკალური სახელი")

            Finally
                isUpdatingComboBoxes = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiaryNames შეცდომა: {ex.Message}")
            isUpdatingComboBoxes = False
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის სპეციფიკური მონაცემების ფილტრაცია
    ''' </summary>
    Private Sub FilterBeneficiarySpecificData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: FilterBeneficiarySpecificData")

            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) Then
                ' ბენეფიციარი არ არის სრულად არჩეული
                currentBeneficiaryData = New List(Of SessionModel)()
                LoadSessionsToGrid(currentBeneficiaryData)
                Return
            End If

            ' კონკრეტული ბენეფიციარის სესიების ფილტრაცია
            currentBeneficiaryData = currentFilteredSessions.Where(Function(s)
                                                                       Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                              s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase)
                                                                   End Function) _
                                                            .OrderBy(Function(s) s.DateTime) _
                                                            .ToList()

            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარი '{selectedName} {selectedSurname}'-ისთვის ნაპოვნია {currentBeneficiaryData.Count} სესია")

            ' სესიების ჩატვირთვა DataGridView-ში
            LoadSessionsToGrid(currentBeneficiaryData)

            ' თერაპევტების და თერაპიების ComboBox-ების განახლება
            UpdateTherapistAndTherapyComboBoxes()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: FilterBeneficiarySpecificData შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ჩატვირთვა DataGridView-ში ინვოისის ფორმატით
    ''' </summary>
    Private Sub LoadSessionsToGrid(sessions As List(Of SessionModel))
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: LoadSessionsToGrid - {sessions.Count} სესია")

            If DgvSessions Is Nothing Then
                Debug.WriteLine("UC_BeneficiaryReport: DgvSessions არ არსებობს")
                Return
            End If

            DgvSessions.Rows.Clear()

            For i As Integer = 0 To sessions.Count - 1
                Dim session = sessions(i)

                ' ინვოისის ნომერი (არა ბაზის ID)
                Dim invoiceNumber As Integer = i + 1

                ' მწკრივის მონაცემები
                Dim rowData As Object() = {
                    invoiceNumber,                                  ' ინვოისის N
                    session.DateTime.ToString("dd.MM.yyyy HH:mm"), ' თარიღი
                    $"{session.Duration} წთ",                       ' ხანგძლივობა
                    session.TherapyType,                            ' თერაპია
                    session.TherapistName,                          ' თერაპევტი
                    session.Price                                   ' თანხა
                }

                Dim addedRowIndex = DgvSessions.Rows.Add(rowData)

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                Try
                    Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = statusColor
                Catch
                    ' სტატუსის ფერის დაყენების შეცდომის შემთხვევაში ნაგულისხმევი
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {sessions.Count} სესია DataGridView-ში")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadSessionsToGrid შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების და თერაპიების ComboBox-ების განახლება
    ''' </summary>
    Private Sub UpdateTherapistAndTherapyComboBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: UpdateTherapistAndTherapyComboBoxes")

            If currentBeneficiaryData Is Nothing OrElse currentBeneficiaryData.Count = 0 Then
                CBPer.Items.Clear()
                CBTer.Items.Clear()
                CBPer.Enabled = False
                CBTer.Enabled = False
                Return
            End If

            isUpdatingComboBoxes = True

            Try
                ' თერაპევტები
                CBPer.Items.Clear()
                CBPer.Items.Add("ყველა")

                Dim uniqueTherapists = currentBeneficiaryData.Where(Function(s) Not String.IsNullOrEmpty(s.TherapistName)) _
                                                           .Select(Function(s) s.TherapistName.Trim()) _
                                                           .Distinct() _
                                                           .OrderBy(Function(therapistItem) therapistItem) _
                                                           .ToList()

                For i As Integer = 0 To uniqueTherapists.Count - 1
                    CBPer.Items.Add(uniqueTherapists(i))
                Next

                CBPer.SelectedIndex = 0
                CBPer.Enabled = True

                ' თერაპიები
                CBTer.Items.Clear()
                CBTer.Items.Add("ყველა")

                Dim uniqueTherapies = currentBeneficiaryData.Where(Function(s) Not String.IsNullOrEmpty(s.TherapyType)) _
                                                             .Select(Function(s) s.TherapyType.Trim()) _
                                                             .Distinct() _
                                                             .OrderBy(Function(therapyItem) therapyItem) _
                                                             .ToList()

                For i As Integer = 0 To uniqueTherapies.Count - 1
                    CBTer.Items.Add(uniqueTherapies(i))
                Next

                CBTer.SelectedIndex = 0
                CBTer.Enabled = True

                Debug.WriteLine($"UC_BeneficiaryReport: თერაპევტები: {uniqueTherapists.Count}, თერაპიები: {uniqueTherapies.Count}")

            Finally
                isUpdatingComboBoxes = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateTherapistAndTherapyComboBoxes შეცდომა: {ex.Message}")
            isUpdatingComboBoxes = False
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები"

    ''' <summary>
    ''' ფილტრის შეცვლის ივენთი - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: OnFilterChanged - isLoadingData: {isLoadingData}")

            ' 🔧 UC_Schedule-ის ანალოგია: მონაცემების დატვირთვის დროს არ ვუშვებთ ციკლს
            If isLoadingData Then
                Debug.WriteLine("UC_BeneficiaryReport: მონაცემების დატვირთვის დროს FilterChanged ივენთი იგნორირებულია")
                Return
            End If

            LoadFilteredData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის სახელის ComboBox-ის შეცვლა - UC_Schedule-ის ანალოგია
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            ' 🔧 UC_Schedule-ის ანალოგია: ციკლური განახლებების თავიდან აცილება
            If isUpdatingComboBoxes Then
                Debug.WriteLine("UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged იგნორირებულია - ComboBox-ები განახლების პროცესშია")
                Return
            End If

            Debug.WriteLine("UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged")

            Dim actualSelectedName As String = If(CBBeneName?.SelectedItem?.ToString(), "")
            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარის სახელი შეიცვალა: '{actualSelectedName}'")

            ' ბენეფიციარის გვარების ComboBox-ის განახლება - UC_Schedule-ის ანალოგია
            UpdateBeneficiarySurnames(actualSelectedName)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის გვარების ComboBox-ის განახლება - UC_Schedule-დან კოპირებული ლოგიკა
    ''' </summary>
    Public Sub UpdateBeneficiarySurnames(selectedName As String)
        Try
            If CBBeneSurname Is Nothing OrElse currentFilteredSessions Is Nothing Then Return

            CBBeneSurname.Items.Clear()

            If String.IsNullOrWhiteSpace(selectedName) Then
                CBBeneSurname.Enabled = False
                CBPer.Enabled = False
                CBTer.Enabled = False

                ' ცარიელი მონაცემები
                currentBeneficiaryData = New List(Of SessionModel)()
                LoadSessionsToGrid(currentBeneficiaryData)
                Return
            End If

            ' უნიკალური გვარების მიღება არჩეული სახელისთვის
            Dim uniqueSurnames = currentFilteredSessions.Where(Function(s)
                                                                   Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                          Not String.IsNullOrEmpty(s.BeneficiarySurname)
                                                               End Function) _
                                                       .Select(Function(s) s.BeneficiarySurname.Trim()) _
                                                       .Distinct() _
                                                       .OrderBy(Function(surnameItem) surnameItem) _
                                                       .ToList()

            For i As Integer = 0 To uniqueSurnames.Count - 1
                CBBeneSurname.Items.Add(uniqueSurnames(i))
            Next

            CBBeneSurname.Enabled = (uniqueSurnames.Count > 0)

            ' თუ მხოლოდ ერთი გვარია, ავტომატურად არჩევა
            If uniqueSurnames.Count = 1 Then
                CBBeneSurname.SelectedIndex = 0
            End If

            Debug.WriteLine($"UC_BeneficiaryReport: სახელისთვის '{selectedName}' ნაპოვნია {uniqueSurnames.Count} გვარი")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateBeneficiarySurnames შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვარების ComboBox-ის განახლება არჩეული სახელისთვის
    ''' </summary>
    Private Sub UpdateSurnamesComboBox(selectedName As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateSurnamesComboBox - სახელი: '{selectedName}'")

            isUpdatingComboBoxes = True

            Try
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Text = ""
                CBPer.Items.Clear()
                CBTer.Items.Clear()

                If String.IsNullOrEmpty(selectedName) Then
                    CBBeneSurname.Enabled = False
                    CBPer.Enabled = False
                    CBTer.Enabled = False
                    currentBeneficiaryData = New List(Of SessionModel)()
                    LoadSessionsToGrid(currentBeneficiaryData)
                    Return
                End If

                ' გვარების ComboBox-ის განახლება არჩეული სახელისთვის
                Dim uniqueSurnames = currentFilteredSessions.Where(Function(s)
                                                                       Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                              Not String.IsNullOrEmpty(s.BeneficiarySurname)
                                                                   End Function) _
                                                            .Select(Function(s) s.BeneficiarySurname.Trim()) _
                                                            .Distinct() _
                                                            .OrderBy(Function(surnameItem) surnameItem) _
                                                            .ToList()

                For i As Integer = 0 To uniqueSurnames.Count - 1
                    CBBeneSurname.Items.Add(uniqueSurnames(i))
                Next

                CBBeneSurname.Enabled = (uniqueSurnames.Count > 0)

                ' თუ მხოლოდ ერთი გვარია, ავტომატურად არჩევა
                If uniqueSurnames.Count = 1 Then
                    CBBeneSurname.SelectedIndex = 0
                End If

                Debug.WriteLine($"UC_BeneficiaryReport: სახელისთვის '{selectedName}' ნაპოვნია {uniqueSurnames.Count} გვარი")

            Finally
                isUpdatingComboBoxes = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateSurnamesComboBox შეცდომა: {ex.Message}")
            isUpdatingComboBoxes = False
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის გვარის ComboBox-ის შეცვლა
    ''' </summary>
    Private Sub CBBeneSurname_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneSurname.SelectedIndexChanged
        Try
            If isUpdatingComboBoxes Then Return

            Debug.WriteLine("UC_BeneficiaryReport: CBBeneSurname_SelectedIndexChanged")

            ' ბენეფიციარის სპეციფიკური მონაცემების განახლება
            FilterBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneSurname_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის ComboBox-ის შეცვლა
    ''' </summary>
    Private Sub CBPer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPer.SelectedIndexChanged
        Try
            If isUpdatingComboBoxes Then Return

            Debug.WriteLine("UC_BeneficiaryReport: CBPer_SelectedIndexChanged")

            ' ბენეფიციარის მონაცემების ხელახალი ფილტრაცია თერაპევტის მიხედვით
            ApplyTherapistAndTherapyFilter()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBPer_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპიის ComboBox-ის შეცვლა
    ''' </summary>
    Private Sub CBTer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTer.SelectedIndexChanged
        Try
            If isUpdatingComboBoxes Then Return

            Debug.WriteLine("UC_BeneficiaryReport: CBTer_SelectedIndexChanged")

            ' ბენეფიციარის მონაცემების ხელახალი ფილტრაცია თერაპიის მიხედვით
            ApplyTherapistAndTherapyFilter()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBTer_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტისა და თერაპიის ფილტრის გამოყენება
    ''' </summary>
    Private Sub ApplyTherapistAndTherapyFilter()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ApplyTherapistAndTherapyFilter")

            If currentBeneficiaryData Is Nothing OrElse currentBeneficiaryData.Count = 0 Then
                Debug.WriteLine("UC_BeneficiaryReport: currentBeneficiaryData ცარიელია")
                Return
            End If

            Dim selectedTherapist As String = CBPer.SelectedItem?.ToString()
            Dim selectedTherapy As String = CBTer.SelectedItem?.ToString()

            ' ფილტრაცია
            Dim filteredData = currentBeneficiaryData.AsEnumerable()

            ' თერაპევტის ფილტრი
            If Not String.IsNullOrEmpty(selectedTherapist) AndAlso selectedTherapist <> "ყველა" Then
                filteredData = filteredData.Where(Function(s) s.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase))
            End If

            ' თერაპიის ფილტრი
            If Not String.IsNullOrEmpty(selectedTherapy) AndAlso selectedTherapy <> "ყველა" Then
                filteredData = filteredData.Where(Function(s) s.TherapyType.Trim().Equals(selectedTherapy, StringComparison.OrdinalIgnoreCase))
            End If

            Dim finalData = filteredData.OrderBy(Function(s) s.DateTime).ToList()

            Debug.WriteLine($"UC_BeneficiaryReport: თერაპევტი '{selectedTherapist}', თერაპია '{selectedTherapy}' - შედეგი: {finalData.Count} სესია")

            ' DataGridView-ის განახლება
            LoadSessionsToGrid(finalData)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ApplyTherapistAndTherapyFilter შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "დამხმარე მეთოდები"

    ''' <summary>
    ''' ინვოისის ჯამური თანხის გამოთვლა
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
                    If row.Cells("Price").Value IsNot Nothing Then
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
    ''' ინვოისის სესიების რაოდენობის მიღება
    ''' </summary>
    ''' <returns>სესიების რაოდენობა</returns>
    Public Function GetInvoiceSessionCount() As Integer
        Try
            Return If(DgvSessions?.Rows.Count, 0)
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' მიმდინარე ბენეფიციარის ინფორმაციის მიღება
    ''' </summary>
    ''' <returns>ბენეფიციარის სრული სახელი</returns>
    Public Function GetCurrentBeneficiaryName() As String
        Try
            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If Not String.IsNullOrEmpty(selectedName) AndAlso Not String.IsNullOrEmpty(selectedSurname) Then
                Return $"{selectedName} {selectedSurname}"
            End If

            Return ""

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetCurrentBeneficiaryName შეცდომა: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' ინვოისის პერიოდის მიღება
    ''' </summary>
    ''' <returns>პერიოდის ტექსტი</returns>
    Public Function GetInvoicePeriod() As String
        Try
            Return $"{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' ინვოისის ვალიდობის შემოწმება
    ''' </summary>
    ''' <returns>True თუ ინვოისი ვალიდურია</returns>
    Public Function IsInvoiceValid() As Boolean
        Try
            Dim beneficiaryName = GetCurrentBeneficiaryName()
            Dim sessionCount = GetInvoiceSessionCount()

            Return Not String.IsNullOrEmpty(beneficiaryName) AndAlso sessionCount > 0

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: IsInvoiceValid შეცდომა: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "რესურსების განთავისუფლება"

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            currentFilteredSessions?.Clear()
            currentBeneficiaryData?.Clear()
            dataProcessor?.ClearCache()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

End Class