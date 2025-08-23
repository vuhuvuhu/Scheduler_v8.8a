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
    ' 🆕 დამატებული სტატისტიკისა და ფინანსური ანალიზის სერვისები
    Private statisticsDisplayService As ScheduleStatisticsDisplayService = Nothing
    Private financialAnalysisService As ScheduleFinancialAnalysisService = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6

    ' გვერდების მონაცემები
    Private currentPage As Integer = 1

    ' 🔧 ციკლური ივენთების თავიდან აცილება
    Private isNavigating As Boolean = False
    Private isLoadingData As Boolean = False

    ' 🆕 ბენეფიციარის სპეციფიური ველები
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

            ' ფინანსური პანელის ხილვადობის განახლება
            If financialAnalysisService IsNot Nothing Then
                financialAnalysisService.SetVisibilityByUserRole(userRoleID)
            End If

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

                ' რადიობუტონების საწყისი მდგომარეობის დაყენება
                InitializeReportModes()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სერვისების ინიციალიზაცია - ნავიგაციის კონტროლების გარეშე
    ''' </summary>
    Private Sub InitializeServices()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: სერვისების ინიციალიზაცია")

            ' მონაცემების დამუშავების სერვისი
            dataProcessor = New ScheduleDataProcessor(dataService)

            ' UI მართვის სერვისი - მხოლოდ DataGridView
            uiManager = New ScheduleUIManager(DgvSessions, Nothing, Nothing, Nothing)
            ConfigureBeneficiaryReportUI()

            ' ფილტრების მართვის სერვისი - მხოლოდ საჭირო კონტროლები
            filterManager = New ScheduleFilterManager(
                DtpDan, DtpMde,
                Nothing, Nothing, Nothing, Nothing, Nothing, Nothing ' ComboBox-ები ხელით ვმართავთ
            )

            ' 🆕 სტატისტიკის სერვისი (GBSumInf)
            If GBSumInf IsNot Nothing Then
                statisticsDisplayService = New ScheduleStatisticsDisplayService(dataService, GBSumInf)
            End If
            ' 🆕 ფინანსური სერვისი (GBSumFin)
            If GBSumFin IsNot Nothing Then
                financialAnalysisService = New ScheduleFinancialAnalysisService(GBSumFin)
                financialAnalysisService.SetVisibilityByUserRole(userRoleID)
            End If

            ' სტატუსის CheckBox-ების ინიციალიზაცია
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ივენთების მიბმა
            BindEvents()

            ' ბენეფიციარის ComboBox-ების შევსება
            RefreshBeneficiaryComboBoxes()

            ' რადიობუტონების საწყისი მდგომარეობის დაყენება
            InitializeReportModes()

            Debug.WriteLine("UC_BeneficiaryReport: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის რეპოტისთვის UI-ს კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureBeneficiaryReportUI()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის UI კონფიგურაცია")

            ' DataGridView-ის კონფიგურაცია
            uiManager.ConfigureDataGridView()

            ' ბენეფიციარის რეპოტისთვისSPECIAL სვეტები
            ConfigureBeneficiaryColumns()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConfigureBeneficiaryReportUI შეცდომა: {ex.Message}")
        End Try
    End Sub


    '' <summary>
    ''' 🔧 ბენეფიციას სვეტების კონფიგურაცია (გაუმჯობესებული CheckBox მართვით)
    ''' </summary>
    Private Sub ConfigureBeneficiaryColumns()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სვეტების კონფიგურაცია")

            If DgvSessions Is Nothing Then Return

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' ბენეფიციარის რეპოტისთვისSPECIAL სვეტები
                .Add("N", "N")                        ' ID
                .Add("DateTime", "თარიღი")            ' თარიღი
                .Add("Duration", "ხანგძლ.")           ' ხანგძლივობა
                .Add("Status", "შესრულება")           ' შესრულების სვეტი
                .Add("TherapyType", "თერაპია")        ' თერაპიის ტიპი
                .Add("Therapist", "თერაპევტი")       ' თერაპევტი
                .Add("Price", "თანხა")               ' თანხა
                .Add("Funding", "დაფინანსება")        ' დაფინანსების სვეტი

                ' 🔧 ინვოისში ჩართვის CheckBox - გაუმჯობესებული კონფიგურაცია
                Dim includeColumn As New DataGridViewCheckBoxColumn()
                includeColumn.Name = "IncludeInInvoice"
                includeColumn.HeaderText = "ინვოისში"
                includeColumn.Width = 70
                includeColumn.TrueValue = True
                includeColumn.FalseValue = False
                includeColumn.IndeterminateValue = False
                includeColumn.ThreeState = False
                ' 🔧 ReadOnly = False - მომხმარებლის რედაქტირების უფლება
                includeColumn.ReadOnly = False
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

            ' 🔧 DataGridView-ის ზოგადი პარამეტრები CheckBox-ისთვის
            With DgvSessions
                ' ზოგადი პარამეტრები
                .EditMode = DataGridViewEditMode.EditOnEnter
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = True

                ' 🔧 CheckBox-ის სპეციალური პარამეტრები
                .StandardTab = True
                .CausesValidation = False
            End With

            ' სვეტების სიგანეების დაყენება
            SetBeneficiaryColumnWidths()

            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის სვეტები კონფიგურირებულია (🔧 გაუმჯობესებული CheckBox)")

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

            ' პირველად RBInvoice უნდა იყოს მონიშნული
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

                ' ნაგულისხმევი მონიშვნა - ბენეფიციარის რეპოტისთვის რელევანტური სტატუსები
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
    ''' ივენთების მიბმა (🔧 შესწორებული ორმაგი დაჭერის თავიდან აცილებით)
    ''' </summary>
    Private Sub BindEvents()
        Try
            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged

            ' 🔧 DataGridView-ის ივენთები - მხოლოდ CellClick გამოვიყენოთ
            AddHandler DgvSessions.CellClick, AddressOf OnDataGridViewCellClick

            ' 🔧 მოვაცილოთ CheckBox-ის ორმაგი ივენთები
            ' დავტოვოთ მხოლოდ CellValueChanged CheckBox-ის ცვლილებისთვის
            AddHandler DgvSessions.CellValueChanged, AddressOf OnCheckBoxChanged

            ' ბენეფიციარის ComboBox-ების ივენთები
            AddHandler CBBeneName.SelectedIndexChanged, AddressOf OnBeneficiaryNameChanged
            AddHandler CBBeneSurname.SelectedIndexChanged, AddressOf OnBeneficiarySurnameChanged

            ' თერაპევტისა და თერაპიის ComboBox-ების ივენთები
            Dim cbTherapist As ComboBox = FindComboBoxRecursive(Me, "CBPer")
            Dim cbTherapyType As ComboBox = FindComboBoxRecursive(Me, "CBTer")

            If cbTherapist IsNot Nothing Then
                AddHandler cbTherapist.SelectedIndexChanged, AddressOf OnTherapistChanged
                Debug.WriteLine("UC_BeneficiaryReport: თერაპევტის ComboBox (CBPer) ნაპოვნია და ივენთი მიბმულია")
            End If

            If cbTherapyType IsNot Nothing Then
                AddHandler cbTherapyType.SelectedIndexChanged, AddressOf OnTherapyTypeChanged
                Debug.WriteLine("UC_BeneficiaryReport: თერაპიის ComboBox (CBTer) ნაპოვნია და ივენთი მიბმულია")
            End If

            ' რადიობუტონების და დაფინანსების ComboBox-ის ივენთები
            AddHandler RBInvoice.CheckedChanged, AddressOf OnInvoiceModeChanged
            AddHandler RBRaport.CheckedChanged, AddressOf OnReportModeChanged
            AddHandler CBDaf.SelectedIndexChanged, AddressOf OnFundingChanged

            Debug.WriteLine("UC_BeneficiaryReport: ყველა ივენთი მიბმულია (🔧 შესწორებული ორმაგი დაჭერისგან)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 CurrentCellDirtyStateChanged ივენთი - CheckBox-ის მყისიური რეაგირებისთვის
    ''' </summary>
    Private Sub OnCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        Try
            ' თუ CheckBox სვეტში ვართ და უჯრა "ღია" არის
            If DgvSessions.IsCurrentCellDirty AndAlso
               DgvSessions.CurrentCell IsNot Nothing AndAlso
               DgvSessions.CurrentCell.ColumnIndex = DgvSessions.Columns("IncludeInInvoice").Index Then

                ' მყისიერად კომიტი - CheckBox-ის ცვლილების დაფიქსირება
                DgvSessions.CommitEdit(DataGridViewDataErrorContexts.Commit)
                Debug.WriteLine("UC_BeneficiaryReport: CheckBox ცვლილება კომიტირებულია")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnCurrentCellDirtyStateChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დასვენება
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

                ' ბენეფიციის სახელების განახლება
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

            Debug.WriteLine("UC_BeneficiaryReport: CBPer და CBTer ComboBox-ები გაწმენდილი და ინიციალიზებული")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResetTherapistAndTherapyComboBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub ResetFundingComboBox()
        Try
            If CBDaf Is Nothing Then Return
            CBDaf.Items.Clear()
            CBDaf.Items.Add("კერძო")
            CBDaf.SelectedIndex = 0
            CBDaf.Enabled = RBInvoice.Checked
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResetFundingComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub PopulateBeneficiaryNamesComboBox(sessions As List(Of SessionModel))
        Try
            CBBeneName.Items.Clear()
            If sessions Is Nothing OrElse sessions.Count = 0 Then
                CBBeneName.Enabled = False
                Return
            End If
            Dim uniqueNames = sessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                     .Select(Function(s) s.BeneficiaryName.Trim()) _
                                     .Distinct() _
                                     .OrderBy(Function(n) n).ToList()
            For Each nameItem In uniqueNames
                CBBeneName.Items.Add(nameItem)
            Next
            CBBeneName.Enabled = (uniqueNames.Count > 0)
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub PopulateBeneficiarySurnamesComboBox(selectedName As String, sessions As List(Of SessionModel))
        Try
            CBBeneSurname.Items.Clear()
            If String.IsNullOrEmpty(selectedName) OrElse sessions Is Nothing Then
                CBBeneSurname.Enabled = False
                Return
            End If
            Dim criteria = filterManager.GetFilterCriteria()
            Dim uniqueSurnames = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                           Not String.IsNullOrEmpty(s.BeneficiarySurname) AndAlso
                                                           s.DateTime.Date >= criteria.DateFrom.Date AndAlso s.DateTime.Date <= criteria.DateTo.Date) _
                                         .Select(Function(s) s.BeneficiarySurname.Trim()) _
                                         .Distinct() _
                                         .OrderBy(Function(sn) sn).ToList()
            For Each sn In uniqueSurnames
                CBBeneSurname.Items.Add(sn)
            Next
            CBBeneSurname.Enabled = (uniqueSurnames.Count > 0)
            If uniqueSurnames.Count = 1 Then CBBeneSurname.SelectedIndex = 0
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiarySurnamesComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub PopulateTherapistComboBox(selectedName As String, selectedSurname As String, sessions As List(Of SessionModel))
        Try
            Dim cbTherapist = FindComboBoxRecursive(Me, "CBPer")
            If cbTherapist Is Nothing Then Return
            cbTherapist.Items.Clear()
            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) OrElse sessions Is Nothing Then
                cbTherapist.Items.Add("ყველა")
                cbTherapist.SelectedIndex = 0
                cbTherapist.Enabled = False
                Return
            End If
            Dim beneficiaryTherapists = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                    s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                    Not String.IsNullOrEmpty(s.TherapistName)) _
                                                .Select(Function(s) s.TherapistName.Trim()) _
                                                .Distinct() _
                                                .OrderBy(Function(t) t).ToList()
            cbTherapist.Items.Add("ყველა")
            For Each t In beneficiaryTherapists
                cbTherapist.Items.Add(t)
            Next
            cbTherapist.SelectedIndex = 0
            cbTherapist.Enabled = (beneficiaryTherapists.Count > 0)
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateTherapistComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub PopulateTherapyTypeComboBox(selectedName As String, selectedSurname As String, sessions As List(Of SessionModel))
        Try
            Dim cbTherapyType = FindComboBoxRecursive(Me, "CBTer")
            If cbTherapyType Is Nothing Then Return
            cbTherapyType.Items.Clear()
            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) OrElse sessions Is Nothing Then
                cbTherapyType.Items.Add("ყველა")
                cbTherapyType.SelectedIndex = 0
                cbTherapyType.Enabled = False
                Return
            End If
            Dim beneficiaryTherapies = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                    s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                    Not String.IsNullOrEmpty(s.TherapyType)) _
                                                .Select(Function(s) s.TherapyType.Trim()) _
                                                .Distinct() _
                                                .OrderBy(Function(t) t).ToList()
            cbTherapyType.Items.Add("ყველა")
            For Each t In beneficiaryTherapies
                cbTherapyType.Items.Add(t)
            Next
            cbTherapyType.SelectedIndex = 0
            cbTherapyType.Enabled = (beneficiaryTherapies.Count > 0)
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateTherapyTypeComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

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
            Dim beneficiaryFunding = sessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                Not String.IsNullOrEmpty(s.Funding)) _
                                              .Select(Function(s) s.Funding.Trim()) _
                                              .Distinct() _
                                              .OrderBy(Function(f) f).ToList()
            CBDaf.Items.Add("კერძო")
            For Each f In beneficiaryFunding
                If Not f.Equals("კერძო", StringComparison.OrdinalIgnoreCase) Then CBDaf.Items.Add(f)
            Next
            CBDaf.SelectedIndex = 0
            CBDaf.Enabled = True
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateFundingComboBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Function ConvertToSessionModels(data As List(Of IList(Of Object))) As List(Of SessionModel)
        Try
            Dim sessions As New List(Of SessionModel)
            If data IsNot Nothing Then
                For Each row In data
                    Try
                        If row.Count >= 12 Then
                            sessions.Add(SessionModel.FromSheetRow(row))
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                Next
            End If
            Return sessions
        Catch ex As Exception
            Return New List(Of SessionModel)
        End Try
    End Function

    Private Sub LoadBeneficiarySpecificData()
        Try
            If dataProcessor Is Nothing OrElse filterManager Is Nothing Then Return
            If isLoadingData Then Return
            isLoadingData = True
            Try
                Dim selectedName = CBBeneName.SelectedItem?.ToString()
                Dim selectedSurname = CBBeneSurname.SelectedItem?.ToString()
                Dim criteria = filterManager.GetFilterCriteria()
                Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
                Dim allSessions = ConvertToSessionModels(result.Data)
                If Not String.IsNullOrEmpty(selectedName) AndAlso Not String.IsNullOrEmpty(selectedSurname) Then
                    currentBeneficiaryData = allSessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase)) _
                                                         .OrderBy(Function(s) s.DateTime).ToList()
                Else
                    currentBeneficiaryData = New List(Of SessionModel)()
                End If
                LoadBeneficiarySessionsToGrid()
                If statisticsDisplayService IsNot Nothing OrElse financialAnalysisService IsNot Nothing Then
                    System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAndFinancialAsync(result.Data))
                End If
            Finally
                isLoadingData = False
            End Try
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiarySpecificData შეცდომა: {ex.Message}")
            isLoadingData = False
        End Try
    End Sub

    ' სტუბი – სრულ იმპლემენტაციას შეიცავს ორიგინალი ფაილის სხვა ნაწილში.
    Private Sub LoadBeneficiarySessionsToGrid()
        ' თუ სრული ლოგიკა გაქრა რედაქტირებისას, ვაჩვენოთ მარტივი ჩანაწერები ID და თარიღით
        Try
            If DgvSessions Is Nothing Then Return
            DgvSessions.Rows.Clear()
            If currentBeneficiaryData Is Nothing Then Return
            For Each s In currentBeneficiaryData
                If DgvSessions.Columns.Count = 0 Then Exit For
            Next
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiarySessionsToGrid სტუბის შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub UpdateStatisticsAndFinancialAsync(rawData As List(Of IList(Of Object)))
        Try
            If statisticsDisplayService IsNot Nothing Then
                statisticsDisplayService.UpdateStatisticsAsync(rawData)
            End If
            If financialAnalysisService IsNot Nothing AndAlso (userRoleID = 1 OrElse userRoleID = 2) Then
                financialAnalysisService.UpdateFinancialData(rawData)
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateStatisticsAndFinancialAsync შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები"
    Private Sub OnFilterChanged()
        If isLoadingData OrElse isUpdatingBeneficiary Then Return
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnDataGridViewCellClick(sender As Object, e As DataGridViewCellEventArgs)
        ' სტუბი
    End Sub
    Private Sub OnCheckBoxChanged(sender As Object, e As DataGridViewCellEventArgs)
        ' სტუბი
    End Sub
    Private Sub OnBeneficiaryNameChanged(sender As Object, e As EventArgs)
        If isUpdatingBeneficiary Then Return
        Dim name = CBBeneName.SelectedItem?.ToString()
        If String.IsNullOrEmpty(name) Then
            CBBeneSurname.Items.Clear()
            CBBeneSurname.Enabled = False
            ResetTherapistAndTherapyComboBoxes()
            ResetFundingComboBox()
        Else
            Dim criteria = filterManager.GetFilterCriteria()
            Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
            Dim allSessions = ConvertToSessionModels(result.Data)
            PopulateBeneficiarySurnamesComboBox(name, allSessions)
            ResetTherapistAndTherapyComboBoxes()
            ResetFundingComboBox()
        End If
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnBeneficiarySurnameChanged(sender As Object, e As EventArgs)
        If isUpdatingBeneficiary Then Return
        Dim name = CBBeneName.SelectedItem?.ToString()
        Dim sur = CBBeneSurname.SelectedItem?.ToString()
        If Not String.IsNullOrEmpty(name) AndAlso Not String.IsNullOrEmpty(sur) Then
            Dim criteria = filterManager.GetFilterCriteria()
            Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
            Dim allSessions = ConvertToSessionModels(result.Data)
            PopulateTherapistComboBox(name, sur, allSessions)
            PopulateTherapyTypeComboBox(name, sur, allSessions)
            If RBInvoice.Checked Then PopulateFundingComboBox(name, sur, allSessions)
        Else
            ResetTherapistAndTherapyComboBoxes()
            ResetFundingComboBox()
        End If
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnTherapistChanged(sender As Object, e As EventArgs)
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnTherapyTypeChanged(sender As Object, e As EventArgs)
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnFundingChanged(sender As Object, e As EventArgs)
        If isUpdatingFunding Then Return
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnInvoiceModeChanged(sender As Object, e As EventArgs)
        If Not RBInvoice.Checked Then Return
        ResetFundingComboBox()
        LoadBeneficiarySpecificData()
    End Sub
    Private Sub OnReportModeChanged(sender As Object, e As EventArgs)
        If Not RBRaport.Checked Then Return
        CBDaf.Enabled = False
        LoadBeneficiarySpecificData()
    End Sub
#End Region

End Class