' 📄 UserControls/UC_BeneficiaryReport.vb
' -------------------------------------------
' ბენეფიციარის რეპორტის UserControl - UC_Schedule.vb-ის საფუძველზე გადაწერილი
' მარტივი და გამართული არქიტექტურა - არა გართულებული UI გაყოფით
' ძირითადი ფუნქცია: ბენეფიციარის არჩევა და მისი სესიების ჩვენება
' 🆕 თერაპევტისა და თერაპიის ფილტრები, ინვოისში შესრულების სვეტი
' 🆕 რადიობუტონები: RBInvoice (ინვოისის რეჟიმი) და RBRaport (რეპორტის რეჟიმი)
' 🆕 დაფინანსების ფილტრი CBDaf - მუშაობს მხოლოდ ინვოისის რეჟიმში
' ===========================================
Imports System.Security.Policy
Imports System.Windows.Forms
Imports iTextSharp.text
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

    ' 🆕 დაფინანსების ფილტრაციის ველები
    Private isUpdatingFunding As Boolean = False

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

                ' 🆕 რადიობუტონების საწყისი მდგომარეობის დაყენება
                InitializeReportModes()
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

            ' 🆕 რადიობუტონების საწყისი მდგომარეობის დაყენება
            InitializeReportModes()

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
    ''' 🆕 ბენეფიციარის სვეტების კონფიგურაცია (შესრულების სვეტით)
    ''' </summary>
    Private Sub ConfigureBeneficiaryColumns()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სვეტების კონფიგურაცია")

            If DgvSessions Is Nothing Then Return

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' ბენეფიციარის რეპორტისთვის სპეციალური სვეტები
                .Add("N", "N")                        ' ID
                .Add("DateTime", "თარიღი")            ' თარიღი
                .Add("Duration", "ხანგძლ.")           ' ხანგძლივობა
                .Add("Status", "შესრულება")           ' 🆕 შესრულების სვეტი
                .Add("TherapyType", "თერაპია")        ' თერაპიის ტიპი
                .Add("Therapist", "თერაპევტი")       ' თერაპევტი
                .Add("Price", "თანხა")               ' თანხა
                .Add("Funding", "დაფინანსება")        ' 🆕 დაფინანსების სვეტი

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
                .Item("Duration").Width = 60
                .Item("Status").Width = 110           ' 🆕 შესრულების სვეტი
                .Item("TherapyType").Width = 180
                .Item("Therapist").Width = 160
                .Item("Price").Width = 80
                .Item("Funding").Width = 100          ' 🆕 დაფინანსების სვეტი
                .Item("IncludeInInvoice").Width = 70
                .Item("Edit").Width = 35

                ' ფორმატირება
                .Item("Price").DefaultCellStyle.Format = "N2"
                .Item("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Item("N").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("IncludeInInvoice").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("Status").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("Funding").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetBeneficiaryColumnWidths შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 რეპორტის რეჟიმების საწყისი ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeReportModes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: რეპორტის რეჟიმების ინიციალიზაცია")

            ' თავიდან RBInvoice უნდა იყოს მონიშნული
            If RBInvoice IsNot Nothing Then
                RBInvoice.Checked = True
            End If

            ' CBDaf-ის საწყისი ინიციალიზაცია
            If CBDaf IsNot Nothing Then
                CBDaf.Items.Clear()
                CBDaf.Items.Add("კერძო")
                CBDaf.SelectedIndex = 0
                CBDaf.Enabled = True
            End If

            Debug.WriteLine("UC_BeneficiaryReport: რეპორტის რეჟიმები ინიციალიზებულია - RBInvoice მონიშნული, CBDaf-ში 'კერძო' არჩეული")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: InitializeReportModes შეცდომა: {ex.Message}")
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
    ''' 🆕 ComboBox-ის რეკურსიული ძებნა
    ''' </summary>
    Private Function FindComboBoxRecursive(parent As Control, name As String) As ComboBox
        Try
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is ComboBox AndAlso ctrl.Name = name Then
                    Return DirectCast(ctrl, ComboBox)
                End If
            Next

            For Each ctrl As Control In parent.Controls
                Dim found = FindComboBoxRecursive(ctrl, name)
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
    ''' 🆕 RadioButton-ის რეკურსიული ძებნა
    ''' </summary>
    Private Function FindRadioButtonRecursive(parent As Control, name As String) As RadioButton
        Try
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is RadioButton AndAlso ctrl.Name = name Then
                    Return DirectCast(ctrl, RadioButton)
                End If
            Next

            For Each ctrl As Control In parent.Controls
                Dim found = FindRadioButtonRecursive(ctrl, name)
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
            ' 🔧 CheckBox სვეტისთვის EditingControlShowing ივენთი
            AddHandler DgvSessions.EditingControlShowing, AddressOf OnEditingControlShowing

            ' ბენეფიციარის ComboBox-ების ივენთები
            AddHandler CBBeneName.SelectedIndexChanged, AddressOf OnBeneficiaryNameChanged
            AddHandler CBBeneSurname.SelectedIndexChanged, AddressOf OnBeneficiarySurnameChanged

            ' 🆕 თერაპევტისა და თერაპიის ComboBox-ების ივენთები
            ' შევამოწმოთ ფორმაზე არსებობს თუ არა ეს კონტროლები
            Dim cbTherapist As ComboBox = FindComboBoxRecursive(Me, "CBPer")     ' 🔧 CBPer - არსებული კონტროლი
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")   ' 🔧 CBTer - არსებული კონტროლი

            If cbTherapist IsNot Nothing Then
                AddHandler cbTherapist.SelectedIndexChanged, AddressOf OnTherapistChanged
                Debug.WriteLine("UC_BeneficiaryReport: თერაპევტის ComboBox (CBPer) ნაპოვნია და ივენთი მიბმულია")
            End If

            If cbTherapyType IsNot Nothing Then
                AddHandler cbTherapyType.SelectedIndexChanged, AddressOf OnTherapyTypeChanged
                Debug.WriteLine("UC_BeneficiaryReport: თერაპიის ComboBox (CBTer) ნაპოვნია და ივენთი მიბმულია")
            End If

            ' 🆕 რადიობუტონების და დაფინანსების ComboBox-ის ივენთები
            ' რადიობუტონები და CBDaf უკვე არსებობს დიზაინში
            AddHandler RBInvoice.CheckedChanged, AddressOf OnInvoiceModeChanged
            AddHandler RBRaport.CheckedChanged, AddressOf OnReportModeChanged
            AddHandler CBDaf.SelectedIndexChanged, AddressOf OnFundingChanged

            Debug.WriteLine("UC_BeneficiaryReport: რადიობუტონები და დაფინანსების ComboBox ივენთები მიბმულია")

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

                ' 🆕 თერაპევტისა და თერაპიის ComboBox-ების გასუფთავება
                ResetTherapistAndTherapyComboBoxes()

                ' 🆕 დაფინანსების ComboBox-ის რესეტი
                ResetFundingComboBox()

            Finally
                isUpdatingBeneficiary = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: RefreshBeneficiaryComboBoxes შეცდომა: {ex.Message}")
            isUpdatingBeneficiary = False
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 თერაპევტისა და თერაპიის ComboBox-ების რესეტი
    ''' </summary>
    Private Sub ResetTherapistAndTherapyComboBoxes()
        Try
            ' თერაპევტის ComboBox (CBPer)
            Dim cbTherapist As ComboBox = FindComboBoxRecursive(Me, "CBPer")
            If cbTherapist IsNot Nothing Then
                cbTherapist.Items.Clear()
                cbTherapist.Items.Add("ყველა")
                cbTherapist.SelectedIndex = 0
                cbTherapist.Enabled = False
            End If

            ' თერაპიის ComboBox (CBTer)
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            If cbTherapyType IsNot Nothing Then
                cbTherapyType.Items.Clear()
                cbTherapyType.Items.Add("ყველა")
                cbTherapyType.SelectedIndex = 0
                cbTherapyType.Enabled = False
            End If

            Debug.WriteLine("UC_BeneficiaryReport: CBPer და CBTer ComboBox-ები დარესეტდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResetTherapistAndTherapyComboBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 დაფინანსების ComboBox-ის რესეტი
    ''' </summary>
    Private Sub ResetFundingComboBox()
        Try
            If CBDaf Is Nothing Then Return

            CBDaf.Items.Clear()
            CBDaf.Items.Add("კერძო")
            CBDaf.SelectedIndex = 0
            CBDaf.Enabled = RBInvoice.Checked ' მხოლოდ ინვოისის რეჟიმში

            Debug.WriteLine("UC_BeneficiaryReport: CBDaf ComboBox დარესეტდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResetFundingComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის სახელების ComboBox-ის შევსება
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
    ''' ბენეფიციარის გვარების ComboBox-ის შევსება
    ''' </summary>
    Private Sub PopulateBeneficiarySurnamesComboBox(selectedName As String, sessions As List(Of SessionModel))
        Try
            CBBeneSurname.Items.Clear()

            If String.IsNullOrEmpty(selectedName) OrElse sessions Is Nothing Then
                CBBeneSurname.Enabled = False
                Return
            End If

            ' 🔧 არჩეული სახელისთვის უნიკალური გვარების მიღება (მხოლოდ პერიოდის სესიებიდან)
            Dim uniqueSurnames = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                       Not String.IsNullOrEmpty(s.BeneficiarySurname) AndAlso
                                                       s.DateTime.Date >= filterManager.GetFilterCriteria().DateFrom.Date AndAlso
                                                       s.DateTime.Date <= filterManager.GetFilterCriteria().DateTo.Date) _
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
    ''' 🆕 თერაპევტის ComboBox-ის შევსება ბენეფიციარისთვის
    ''' </summary>
    Private Sub PopulateTherapistComboBox(selectedName As String, selectedSurname As String, sessions As List(Of SessionModel))
        Try
            Dim cbTherapist As ComboBox = FindComboBoxRecursive(Me, "CBPer")
            If cbTherapist Is Nothing Then Return

            cbTherapist.Items.Clear()

            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) OrElse sessions Is Nothing Then
                cbTherapist.Items.Add("ყველა")
                cbTherapist.SelectedIndex = 0
                cbTherapist.Enabled = False
                Return
            End If

            ' კონკრეტული ბენეფიციარის თერაპევტების მიღება
            Dim beneficiaryTherapists = sessions.Where(Function(s)
                                                           Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                 s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                 Not String.IsNullOrEmpty(s.TherapistName)
                                                       End Function) _
                                               .Select(Function(s) s.TherapistName.Trim()) _
                                               .Distinct() _
                                               .OrderBy(Function(therapist) therapist) _
                                               .ToList()

            cbTherapist.Items.Add("ყველა")
            For Each therapist In beneficiaryTherapists
                cbTherapist.Items.Add(therapist)
            Next

            cbTherapist.SelectedIndex = 0
            cbTherapist.Enabled = (beneficiaryTherapists.Count > 0)

            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარისთვის '{selectedName} {selectedSurname}' ნაპოვნია {beneficiaryTherapists.Count} თერაპევტი (CBPer)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateTherapistComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 თერაპიის ComboBox-ის შევსება ბენეფიციარისთვის
    ''' </summary>
    Private Sub PopulateTherapyTypeComboBox(selectedName As String, selectedSurname As String, sessions As List(Of SessionModel))
        Try
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            If cbTherapyType Is Nothing Then Return

            cbTherapyType.Items.Clear()

            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) OrElse sessions Is Nothing Then
                cbTherapyType.Items.Add("ყველა")
                cbTherapyType.SelectedIndex = 0
                cbTherapyType.Enabled = False
                Return
            End If

            ' კონკრეტული ბენეფიციარის თერაპიების მიღება
            Dim beneficiaryTherapies = sessions.Where(Function(s)
                                                          Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                Not String.IsNullOrEmpty(s.TherapyType)
                                                      End Function) _
                                              .Select(Function(s) s.TherapyType.Trim()) _
                                              .Distinct() _
                                              .OrderBy(Function(therapy) therapy) _
                                              .ToList()

            cbTherapyType.Items.Add("ყველა")
            For Each therapy In beneficiaryTherapies
                cbTherapyType.Items.Add(therapy)
            Next

            cbTherapyType.SelectedIndex = 0
            cbTherapyType.Enabled = (beneficiaryTherapies.Count > 0)

            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარისთვის '{selectedName} {selectedSurname}' ნაპოვნია {beneficiaryTherapies.Count} თერაპია (CBTer)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateTherapyTypeComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 დაფინანსების ComboBox-ის შევსება ბენეფიციარისთვის
    ''' </summary>
    Private Sub PopulateFundingComboBox(selectedName As String, selectedSurname As String, sessions As List(Of SessionModel))
        Try
            If CBDaf Is Nothing Then Return

            CBDaf.Items.Clear()

            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) OrElse sessions Is Nothing Then
                CBDaf.Items.Add("კერძო")
                CBDaf.SelectedIndex = 0
                CBDaf.Enabled = False
                Return
            End If

            ' კონკრეტული ბენეფიციარის დაფინანსებების მიღება
            Dim beneficiaryFunding = sessions.Where(Function(s)
                                                        Return s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                              s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                              Not String.IsNullOrEmpty(s.Funding)
                                                    End Function) _
                                             .Select(Function(s) s.Funding.Trim()) _
                                             .Distinct() _
                                             .OrderBy(Function(funding) funding) _
                                             .ToList()

            ' ყოველთვის ვამატებთ "კერძო"-ს
            CBDaf.Items.Add("კერძო")
            For Each funding In beneficiaryFunding
                If Not funding.Equals("კერძო", StringComparison.OrdinalIgnoreCase) Then
                    CBDaf.Items.Add(funding)
                End If
            Next

            ' "კერძო"-ს არჩევა ნაგულისხმევად
            CBDaf.SelectedIndex = 0
            CBDaf.Enabled = True

            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარისთვის '{selectedName} {selectedSurname}' ნაპოვნია {beneficiaryFunding.Count} დაფინანსება (CBDaf)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateFundingComboBox შეცდომა: {ex.Message}")
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
    ''' 🆕 ბენეფიციარის სესიების ჩატვირთვა DataGridView-ში (თერაპ. ფილტრებით)
    ''' </summary>
    Private Sub LoadBeneficiarySessionsToGrid()
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარის სესიების ჩატვირთვა - {If(currentBeneficiaryData?.Count, 0)} სესია")

            DgvSessions.Rows.Clear()

            If currentBeneficiaryData Is Nothing OrElse currentBeneficiaryData.Count = 0 Then
                Return
            End If

            ' 🆕 თერაპევტისა და თერაპიის ფილტრების მიღება
            Dim selectedTherapist As String = GetSelectedTherapist()
            Dim selectedTherapyType As String = GetSelectedTherapyType()
            Dim selectedFunding As String = GetSelectedFunding()

            For i As Integer = 0 To currentBeneficiaryData.Count - 1
                Dim session = currentBeneficiaryData(i)

                ' 🆕 თერაპევტის ფილტრაცია
                If Not String.IsNullOrEmpty(selectedTherapist) AndAlso selectedTherapist <> "ყველა" Then
                    If Not session.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                ' 🆕 თერაპიის ფილტრაცია
                If Not String.IsNullOrEmpty(selectedTherapyType) AndAlso selectedTherapyType <> "ყველა" Then
                    If Not session.TherapyType.Trim().Equals(selectedTherapyType, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                ' 🆕 დაფინანსების ფილტრაცია (მხოლოდ ინვოისის რეჟიმში)
                If RBInvoice.Checked AndAlso Not String.IsNullOrEmpty(selectedFunding) Then
                    If Not session.Funding.Trim().Equals(selectedFunding, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                ' მწკრივის მონაცემები (შესრულების სვეტით + დაფინანსება)
                Dim rowData As Object() = {
                    session.Id,                                     ' N
                    session.DateTime.ToString("dd.MM.yyyy HH:mm"), ' თარიღი
                    $"{session.Duration}წთ",                        ' ხანგძლივობა
                    session.Status,                                 ' 🆕 შესრულების სვეტი
                    session.TherapyType,                            ' თერაპია
                    session.TherapistName,                          ' თერაპევტი
                    session.Price,                                  ' თანხა
                    session.Funding,                                ' 🆕 დაფინანსების სვეტი
                    True,                                           ' ნაგულისხმევად ინვოისში ჩართული
                    "✎"                                             ' რედაქტირების ღილაკი
                }

                ' მწკრივის მონაცემების ჩატვირთვის შემდეგ ვიზუალური სტილის დაყენება
                Dim addedRowIndex = DgvSessions.Rows.Add(rowData)

                ' სესიის ID-ის შენახვა მწკრივში
                DgvSessions.Rows(addedRowIndex).Tag = session.Id

                ' 🔧 CheckBox-ის საწყისი მნიშვნელობის განახლება (ნაგულისხმევად ჩართული)
                DgvSessions.Rows(addedRowIndex).Cells("IncludeInInvoice").Value = True

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                Try
                    Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = statusColor
                Catch
                    ' სტატუსის ფერის შეცდომის შემთხვევაში
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {DgvSessions.Rows.Count} სესია (ფილტრაციის შემდეგ)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiarySessionsToGrid შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 არჩეული თერაპევტის მიღება
    ''' </summary>
    Private Function GetSelectedTherapist() As String
        Try
            Dim cbTherapist As ComboBox = FindComboBoxRecursive(Me, "CBPer")
            Return If(cbTherapist?.SelectedItem?.ToString(), "ყველა")
        Catch
            Return "ყველა"
        End Try
    End Function

    ''' <summary>
    ''' 🆕 არჩეული თერაპიის მიღება
    ''' </summary>
    Private Function GetSelectedTherapyType() As String
        Try
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")
            Return If(cbTherapyType?.SelectedItem?.ToString(), "ყველა")
        Catch
            Return "ყველა"
        End Try
    End Function

    ''' <summary>
    ''' 🆕 არჩეული დაფინანსების მიღება
    ''' </summary>
    Private Function GetSelectedFunding() As String
        Try
            Return If(CBDaf?.SelectedItem?.ToString(), "კერძო")
        Catch
            Return "კერძო"
        End Try
    End Function

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

                ' 🆕 თერაპევტისა და თერაპიის ComboBox-ების რესეტი (სახელი მაინც შეიცვალა)
                ResetTherapistAndTherapyComboBoxes()
                ResetFundingComboBox()
            Else
                ' სახელი არ არის არჩეული - ყველა ComboBox-ის გასუფთავება
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Enabled = False
                ResetTherapistAndTherapyComboBoxes()
                ResetFundingComboBox()
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

            ' 🆕 თერაპევტისა და თერაპიის ComboBox-ების განახლება
            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If Not String.IsNullOrEmpty(selectedName) AndAlso Not String.IsNullOrEmpty(selectedSurname) Then
                ' ფილტრაციის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)

                ' თერაპევტისა და თერაპიის ComboBox-ების განახლება
                PopulateTherapistComboBox(selectedName, selectedSurname, allSessions)
                PopulateTherapyTypeComboBox(selectedName, selectedSurname, allSessions)

                ' 🆕 დაფინანსების ComboBox-ის განახლება (მხოლოდ ინვოისის რეჟიმში)
                If RBInvoice.Checked Then
                    PopulateFundingComboBox(selectedName, selectedSurname, allSessions)
                End If
            Else
                ' ბენეფიციარი არ არის სრულად არჩეული - ComboBox-ების რესეტი
                ResetTherapistAndTherapyComboBoxes()
                ResetFundingComboBox()
            End If

            ' მონაცემების განახლება
            LoadBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnBeneficiarySurnameChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 თერაპევტის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnTherapistChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary Then Return

            Debug.WriteLine("UC_BeneficiaryReport: თერაპევტი შეიცვალა")

            ' მონაცემების განახლება ახალი თერაპევტის ფილტრით
            LoadBeneficiarySessionsToGrid()
            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnTherapistChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 თერაპიის ტიპის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnTherapyTypeChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary Then Return

            Debug.WriteLine("UC_BeneficiaryReport: თერაპიის ტიპი შეიცვალა")

            ' მონაცემების განახლება ახალი თერაპიის ფილტრით
            LoadBeneficiarySessionsToGrid()
            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnTherapyTypeChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 დაფინანსების შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFundingChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary OrElse isUpdatingFunding Then Return

            Debug.WriteLine("UC_BeneficiaryReport: დაფინანსება შეიცვალა")

            ' მონაცემების განახლება ახალი დაფინანსების ფილტრით
            LoadBeneficiarySessionsToGrid()
            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnFundingChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის რეჟიმის ივენთი (RBInvoice)
    ''' </summary>
    Private Sub OnInvoiceModeChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary OrElse isUpdatingFunding Then Return

            Dim rbInvoice As RadioButton = TryCast(sender, RadioButton)
            If rbInvoice Is Nothing OrElse Not rbInvoice.Checked Then Return

            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის რეჟიმი ჩართულია")

            isUpdatingFunding = True

            Try
                ' CBDaf-ის ჩართვა და განახლება
                If CBDaf IsNot Nothing Then
                    CBDaf.Enabled = True

                    ' ბენეფიციარის დაფინანსებების განახლება
                    Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
                    Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

                    If Not String.IsNullOrEmpty(selectedName) AndAlso Not String.IsNullOrEmpty(selectedSurname) Then
                        ' ფილტრაციის კრიტერიუმების მიღება
                        Dim criteria = filterManager.GetFilterCriteria()
                        Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                        Dim allSessions = ConvertToSessionModels(result.Data)

                        PopulateFundingComboBox(selectedName, selectedSurname, allSessions)
                    Else
                        ResetFundingComboBox()
                    End If
                End If

            Finally
                isUpdatingFunding = False
            End Try

            ' მონაცემების განახლება
            LoadBeneficiarySessionsToGrid()
            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnInvoiceModeChanged შეცდომა: {ex.Message}")
            isUpdatingFunding = False
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 რეპორტის რეჟიმის ივენთი (RBRaport)
    ''' </summary>
    Private Sub OnReportModeChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingBeneficiary OrElse isUpdatingFunding Then Return

            Dim rbReport As RadioButton = TryCast(sender, RadioButton)
            If rbReport Is Nothing OrElse Not rbReport.Checked Then Return

            Debug.WriteLine("UC_BeneficiaryReport: რეპორტის რეჟიმი ჩართულია")

            isUpdatingFunding = True

            Try
                ' CBDaf-ის გამორთვა რეპორტის რეჟიმში
                If CBDaf IsNot Nothing Then
                    CBDaf.Enabled = False
                End If

            Finally
                isUpdatingFunding = False
            End Try

            ' მონაცემების განახლება (ყველა დაფინანსების სესია)
            LoadBeneficiarySessionsToGrid()
            UpdateInvoiceTotals()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnReportModeChanged შეცდომა: {ex.Message}")
            isUpdatingFunding = False
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 EditingControlShowing ივენთი CheckBox-ების სწორი მუშაობისთვის
    ''' </summary>
    Private Sub OnEditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs)
        Try
            If DgvSessions.CurrentCell.ColumnIndex = DgvSessions.Columns("IncludeInInvoice").Index Then
                If TypeOf e.Control Is CheckBox Then
                    Dim chk As CheckBox = DirectCast(e.Control, CheckBox)
                    ' ძველი ივენთების წაშლა
                    RemoveHandler chk.CheckedChanged, AddressOf OnInvoiceCheckBoxChanged
                    ' ახალი ივენთის დამატება
                    AddHandler chk.CheckedChanged, AddressOf OnInvoiceCheckBoxChanged
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnEditingControlShowing შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის CheckBox-ის ცვლილება (ღია ივენთი)
    ''' </summary>
    Private Sub OnInvoiceCheckBoxChanged(sender As Object, e As EventArgs)
        Try
            If DgvSessions.CurrentRow IsNot Nothing Then
                Dim rowIndex As Integer = DgvSessions.CurrentRow.Index
                Dim isChecked As Boolean = DirectCast(sender, CheckBox).Checked

                Debug.WriteLine($"UC_BeneficiaryReport: CheckBox შეიცვალა - მწკრივი {rowIndex}, მონიშნული: {isChecked}")

                ' მწკრივის ვიზუალური სტილის განახლება
                UpdateRowVisualStyle(rowIndex, isChecked)

                ' ინვოისის ჯამების განახლება
                UpdateInvoiceTotals()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnInvoiceCheckBoxChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 CheckBox-ის შეცვლის ივენთი (ინვოისში ჩართვა/ამოღება) - CellValueChanged
    ''' </summary>
    Private Sub OnCheckBoxChanged(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
              DgvSessions.Columns(e.ColumnIndex).Name = "IncludeInInvoice" Then

                Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის CheckBox (CellValueChanged) შეიცვალა - მწკრივი {e.RowIndex}")

                ' CheckBox-ის მნიშვნელობის მიღება
                Dim isIncluded As Boolean = True
                If DgvSessions.Rows(e.RowIndex).Cells("IncludeInInvoice").Value IsNot Nothing Then
                    Boolean.TryParse(DgvSessions.Rows(e.RowIndex).Cells("IncludeInInvoice").Value.ToString(), isIncluded)
                End If

                ' მწკრივის ვიზუალური სტილის განახლება
                UpdateRowVisualStyle(e.RowIndex, isIncluded)

                ' ინვოისის ჯამების განახლება
                UpdateInvoiceTotals()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnCheckBoxChanged შეცდომა: {ex.Message}")
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
    ''' 📊 Excel ექსპორტის ღილაკი - ბენეფიციარის ინვოისი Excel ფორმატში
    ''' </summary>
    Private Sub BtnToExcel_Click(sender As Object, e As EventArgs) Handles BtnToExcel.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: Excel ექსპორტის ღილაკზე დაჭერა")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                MessageBox.Show("Excel ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ბენეფიციარის ინფორმაციის შემოწმება
            Dim beneficiaryInfo = GetCurrentBeneficiaryInfo()
            If String.IsNullOrEmpty(beneficiaryInfo) Then
                MessageBox.Show("აირჩიეთ ბენეფიციარი Excel ექსპორტისთვის", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ინვოისის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("Excel ექსპორტისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' CSV ფაილის შექმნა Excel-ისთვის
            CreateExcelInvoiceFile()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BtnToExcel_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"Excel ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
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
                    ' ვიზუალური სტილის განახლება
                    UpdateRowVisualStyle(row.Index, True)
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
                    ' ვიზუალური სტილის განახლება
                    UpdateRowVisualStyle(row.Index, False)
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
    ''' 🆕 მწკრივის ვიზუალური სტილის განახლება CheckBox-ის მდგომარეობის მიხედვით
    ''' </summary>
    ''' <param name="rowIndex">მწკრივის ინდექსი</param>
    ''' <param name="isIncluded">ინვოისში ჩართულია თუ არა</param>
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
                        ' SessionModel-ის მიღება currentBeneficiaryData-დან
                        Dim sessionId As Integer = 0
                        If row.Tag IsNot Nothing AndAlso Integer.TryParse(row.Tag.ToString(), sessionId) Then
                            Dim session = currentBeneficiaryData?.FirstOrDefault(Function(s) s.Id = sessionId)
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
                Dim strikeFont As Font = New Font(DgvSessions.DefaultCellStyle.Font, FontStyle.Strikeout)
                row.DefaultCellStyle.Font = strikeFont
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateRowVisualStyle შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბენეფიციარის ინვოისის ბეჭდვა
    ''' </summary>
    Private Sub PrintBeneficiaryInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის ინვოისის ბეჭდვა")

            ' ინვოისის ვალიდობის შემოწმება
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
    ''' 🆕 Excel ინვოისის ფაილის შექმნა (CSV ფორმატით)
    ''' </summary>
    Private Sub CreateExcelInvoiceFile()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: Excel ინვოისის ფაილის შექმნა")

            Dim beneficiaryName = GetCurrentBeneficiaryInfo().Replace(" ", "_")
            Dim period = $"{DtpDan.Value:dd.MM.yyyy}-{DtpMde.Value:dd.MM.yyyy}"

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "CSV ფაილები (*.csv)|*.csv|Excel ფაილები (*.xlsx)|*.xlsx"
                saveDialog.Title = "ინვოისის Excel ექსპორტი"
                saveDialog.FileName = $"ინვოისი_{beneficiaryName}_{period}_{DateTime.Now:yyyyMMdd}.csv"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    ' CSV ფაილის შექმნა
                    Dim csv As New System.Text.StringBuilder()
                    Dim utf8WithBom As New System.Text.UTF8Encoding(True)

                    ' ინვოისის სათაური
                    csv.AppendLine("""ინვოისი მომსახურების გაწევაზე""")
                    csv.AppendLine("""შპს """"ბავშვთა და მოზარდთა განვითარების, აბილიტაციისა და რეაბილიტაციის ცენტრი - პროსპერო""""""")
                    csv.AppendLine()
                    csv.AppendLine($"""ბენეფიციარი"",""{EscapeCSV(GetCurrentBeneficiaryInfo())}""")
                    csv.AppendLine($"""პერიოდი"",""{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}""")
                    csv.AppendLine($"""შექმნილია"",""{DateTime.Now:dd.MM.yyyy HH:mm}""")
                    csv.AppendLine()

                    ' სათაურები
                    csv.AppendLine("""N"",""თარიღი"",""ხანგძლ."",""შესრულება"",""მომსახურების სახე"",""თერაპევტი"",""დაფინანსება"",""თანხა (₾)""")

                    ' მონაცემები (მხოლოდ ინვოისში ჩართული)
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
                                Dim duration As String = If(row.Cells("Duration").Value?.ToString(), "")
                                Dim status As String = If(row.Cells("Status").Value?.ToString(), "")
                                Dim therapyType As String = If(row.Cells("TherapyType").Value?.ToString(), "")
                                Dim therapist As String = If(row.Cells("Therapist").Value?.ToString(), "")
                                Dim funding As String = If(row.Cells("Funding").Value?.ToString(), "")

                                Dim price As Decimal = 0
                                If row.Cells("Price").Value IsNot Nothing Then
                                    Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                                End If

                                totalAmount += price

                                csv.AppendLine($"""{invoiceNumber}"",""{EscapeCSV(dateTime)}"",""{EscapeCSV(duration)}"",""{EscapeCSV(status)}"",""{EscapeCSV(therapyType)}"",""{EscapeCSV(therapist)}"",""{EscapeCSV(funding)}"",""{price:N2}""")
                                invoiceNumber += 1
                            End If

                        Catch
                            Continue For
                        End Try
                    Next

                    ' ჯამი
                    csv.AppendLine()
                    csv.AppendLine($"""ჯამური თანხა:"","","","","","","","{totalAmount: N2}""")
                    csv.AppendLine($"""თანხა სიტყვიერად"",""{EscapeCSV(ConvertAmountToWords(totalAmount))}""")

                    ' ფაილის ჩაწერა
                    System.IO.File.WriteAllText(saveDialog.FileName, csv.ToString(), utf8WithBom)

                    Debug.WriteLine("UC_BeneficiaryReport Excel ინვოისი შეიქმნა")
                    MessageBox.Show($"Excel ინვოისი წარმატებით შეიქმნა: {Environment.NewLine}{saveDialog.FileName}", "წარმატება",
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
            Debug.WriteLine($"UC_BeneficiaryReport CreateExcelInvoiceFile შეცდომა:  {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ინვოისის HTML-ის შექმნა (შესრულების სვეტით + დაფინანსება)
    ''' </summary>
    ''' <param name="forPrinting">True თუ ბეჭდვისთვის, False თუ ფაილისთვის</param>
    ''' <param name="filePath">ფაილის მისამართი (ფაილისთვის)</param>
    Private Sub CreateInvoiceHTML(forPrinting As Boolean, Optional filePath As String = "")
        Try
            Debug.WriteLine($"UC_BeneficiaryReport ინვოისის html შექმნა - ბეჭდვისთვის:  {forPrinting}")

            Dim html As New System.Text.StringBuilder()
            Dim beneficiaryName = GetCurrentBeneficiaryInfo()
            Dim period = $"{DtpDan.Valuedd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}"

            ' HTML დოკუმენტის შექმნა
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html lang=""ka"">")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>ინვოისი - " & beneficiaryName & "</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        @page { size A4; margin: 20mm; }")
            html.AppendLine("        @media print { .no-print { display none; } }")
            html.AppendLine("        body { font-family 'Sylfaen', Arial, sans-serif; font-size: 12px; line-height: 1.4; }")
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

                    ' მომსახურებების ცხრილი (შესრულების სვეტით + დაფინანსება)
                    html.AppendLine("    <table>")
                    html.AppendLine("        <thead>")
                    html.AppendLine("            <tr>")
                    html.AppendLine("                <th style=""width: 30px;"">N</th>")
                    html.AppendLine("                <th style=""width: 100px;"">თარიღი</th>")
                    html.AppendLine("                <th style=""width: 60px;"">ხანგძლ.</th>")
                    html.AppendLine("                <th style=""width: 110px;"">შესრულება</th>")
                    html.AppendLine("                <th style=""width: 180px;"">მომსახურების სახე</th>")
                    html.AppendLine("                <th style=""width: 140px;"">თერაპევტი</th>")
                    html.AppendLine("                <th style=""width: 100px;"">დაფინანსება</th>")
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
                                Dim duration As String = If(row.Cells("Duration").Value?.ToString(), "")
                                Dim status As String = If(row.Cells("Status").Value?.ToString(), "")
                                Dim therapyType As String = If(row.Cells("TherapyType").Value?.ToString(), "")
                                Dim therapist As String = If(row.Cells("Therapist").Value?.ToString(), "")
                                Dim funding As String = If(row.Cells("Funding").Value?.ToString(), "")

                                Dim price As Decimal = 0
                                If row.Cells("Price").Value IsNot Nothing Then
                                    Decimal.TryParse(row.Cells("Price").Value.ToString(), price)
                                End If

                                totalAmount += price

                                html.AppendLine("            <tr>")
                                html.AppendLine($"                <td style=""text-align: center;"">{invoiceNumber}</td>")
                                html.AppendLine($"                <td>{EscapeHtml(dateTime)}</td>")
                                html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(duration)}</td>")
                                html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(status)}</td>")
                                html.AppendLine($"                <td>{EscapeHtml(therapyType)}</td>")
                                html.AppendLine($"                <td>{EscapeHtml(therapist)}</td>")
                                html.AppendLine($"                <td style=""text-align: center;"">{EscapeHtml(funding)}</td>")
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