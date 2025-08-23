' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განახლებული განრიგის UserControl - გამარტივებული და ოპტიმიზირებული
' მომხმარებლის ელფოსტისა და როლის სწორი მიღება Form1-დან
' დამატებულია რედაქტირების ფორმის ორმაგი გახსნის პრევენცია
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Schedule
    Inherits UserControl

#Region "ველები და თვისებები"

    ' სერვისები
    Private dataService As IDataService = Nothing
    Private dataProcessor As ScheduleDataProcessor = Nothing
    Private uiManager As ScheduleUIManager = Nothing
    Private filterManager As ScheduleFilterManager = Nothing
    Private statisticsDisplayService As ScheduleStatisticsDisplayService = Nothing
    Private financialAnalysisService As ScheduleFinancialAnalysisService = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6

    ' გვერდების მონაცემები
    Private currentPage As Integer = 1

    ' 🔧 ციკლური ივენთების თავიდან აცილება - ეს ველები დაამატეთ
    Private isNavigating As Boolean = False
    Private isLoadingData As Boolean = False

    ' 🔧 რედაქტირების ფორმის დუბლიური გახსნის პრევენცია
    Private _editDialogOpen As Boolean = False
    Private _lastEditDialogClosed As DateTime = DateTime.MinValue
#End Region

#Region "საჯარო მეთოდები"

    ''' <summary>
    ''' 🔧 გამარტივებული SetDataService მეთოდი
    ''' მხოლოდ dataService-ის დაყენება, მომხმარებლის ინფორმაცია განცალკევებით
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    Public Sub SetDataService(service As IDataService)
        Try
            Debug.WriteLine("UC_Schedule: SetDataService")

            dataService = service

            ' სერვისების ინიციალიზაცია
            If dataService IsNot Nothing Then
                InitializeServices()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetDataService შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 მომხმარებლის ინფორმაციის დაყენება Form1-დან
    ''' </summary>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub SetUserInfo(email As String, role As String)
        Try
            Debug.WriteLine($"UC_Schedule: SetUserInfo - email='{email}', role='{role}'")

            userEmail = email

            ' როლის პარსინგი
            Dim parsedRole As Integer
            If Integer.TryParse(role, parsedRole) Then
                userRoleID = parsedRole
            Else
                userRoleID = 6 ' ნაგულისხმევი
            End If

            Debug.WriteLine($"UC_Schedule: საბოლოო userRoleID: {userRoleID}, userEmail: '{userEmail}'")

            ' ფინანსური პანელის ხილვობის განახლება
            If financialAnalysisService IsNot Nothing Then
                financialAnalysisService.SetVisibilityByUserRole(userRoleID)
            End If

            ' UI-ს განახლება ახალი როლისთვის
            If uiManager IsNot Nothing Then
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetUserInfo შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების განახლება
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("UC_Schedule: მონაცემების განახლება")

            ' ქეშის გასუფთავება
            If TypeOf dataService Is SheetDataService Then
                DirectCast(dataService, SheetDataService).InvalidateAllCache()
            End If

            dataProcessor?.ClearCache()

            ' ComboBox-ების განახლება
            If filterManager IsNot Nothing AndAlso dataProcessor IsNot Nothing Then
                filterManager.PopulateFilterComboBoxes(dataProcessor)
            End If

            ' მონაცემების ხელახალი ჩატვირთვა
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: RefreshData შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' კონკრეტულ გვერდზე გადასვლა - შესწორებული ვერსია
    ''' </summary>
    Public Sub GoToPage(pageNumber As Integer)
        Try
            Debug.WriteLine($"UC_Schedule: GoToPage - მოთხოვნილი გვერდი: {pageNumber}, მიმდინარე: {currentPage}")

            If pageNumber >= 1 AndAlso pageNumber <> currentPage Then
                currentPage = pageNumber ' 🔧 ჯერ currentPage-ს ვაყენებთ
                Debug.WriteLine($"UC_Schedule: გვერდის შეცვლა: {currentPage}")
                LoadFilteredSchedule() ' 🔧 შემდეგ ვტვირთავთ
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: GoToPage შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "პირადი მეთოდები"

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_Schedule: კონტროლის ჩატვირთვა")

            ' ფონის ფერების დასაყენებლად
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დანიშნულია, ვასრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სერვისების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeServices()
        Try
            Debug.WriteLine("UC_Schedule: სერვისების ინიციალიზაცია")

            ' მონაცემების დამუშავების სერვისი
            dataProcessor = New ScheduleDataProcessor(dataService)

            ' UI მართვის სერვისი
            uiManager = New ScheduleUIManager(DgvSchedule, LPage, BtnPrev, BtnNext)
            uiManager.ConfigureDataGridView()
            uiManager.ConfigureNavigationButtons()

            ' ფილტრების მართვის სერვისი
            filterManager = New ScheduleFilterManager(
                DtpDan, DtpMde,
                CBBeneName, CBBeneSurname, CBPer, CBTer, CBSpace, CBDaf,
                RB20, RB50, RB100
            )

            ' სტატისტიკის ჩვენების სერვისი
            If GBSumInf IsNot Nothing Then
                statisticsDisplayService = New ScheduleStatisticsDisplayService(dataService, GBSumInf)
            End If

            ' ფინანსური ანალიზის სერვისი
            If GBSumFin IsNot Nothing Then
                financialAnalysisService = New ScheduleFinancialAnalysisService(GBSumFin)
                financialAnalysisService.SetVisibilityByUserRole(userRoleID)
            End If

            ' სტატუსის CheckBox-ების ინციალიზაცია
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ივენთების მიბმა
            BindEvents()

            ' ComboBox-ების შევსება
            System.Threading.Tasks.Task.Run(Sub() PopulateComboBoxesAsync())

            Debug.WriteLine("UC_Schedule: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ComboBox-ების ასინქრონული შევსება
    ''' </summary>
    Private Sub PopulateComboBoxesAsync()
        Try
            Me.Invoke(Sub()
                          If filterManager IsNot Nothing AndAlso dataProcessor IsNot Nothing Then
                              filterManager.PopulateFilterComboBoxes(dataProcessor)
                          End If
                      End Sub)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: PopulateComboBoxesAsync შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის CheckBox-ების ინიციისალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            Dim statusCheckBoxes As New List(Of CheckBox)

            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBoxRecursive(Me, $"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    statusCheckBoxes.Add(checkBox)
                End If
            Next

            filterManager.InitializeStatusCheckBoxes(statusCheckBoxes.ToArray())

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: InitializeStatusCheckBoxes შეცდომა: {ex.Message}")
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
    ''' ივენთების მიბმა - ნავიგაციის ღილაკების ივენთები ამოღებულია
    ''' </summary>
    Private Sub BindEvents()
        Try
            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged
            AddHandler filterManager.PageSizeChanged, AddressOf OnPageSizeChanged

            ' DataGridView-ის ივენთები
            AddHandler DgvSchedule.CellClick, AddressOf OnDataGridViewCellClick

            ' 🔧 ნავიგაციის ღილაკების ივენთები ამოღებულია - იყენებენ Handles-ს

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დასაყენებლად
    ''' </summary>
    Private Sub SetBackgroundColors()
        Try
            Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

            If pnlFilter IsNot Nothing Then pnlFilter.BackColor = transparentWhite
            If GBSumInf IsNot Nothing Then GBSumInf.BackColor = transparentWhite
            If GBSumFin IsNot Nothing Then GBSumFin.BackColor = transparentWhite

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetBackgroundColors შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' განრიგის მონაცემების ჩატვირთვა - შესწორებული ვერსია
    ''' currentPage-ის ორმაგი განახლების თავიდან ასაცილებლად
    ''' </summary>
    Private Sub LoadFilteredSchedule()
        Try
            ' 🔧 შემოწმებები ციკლური გადაკვრის თავიდან ასაცილებლად
            If dataProcessor Is Nothing OrElse uiManager Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_Schedule: სერვისები არ არის ინციალიზებული")
                Return
            End If

            If isLoadingData AndAlso Not isNavigating Then
                Debug.WriteLine("UC_Schedule: მონაცემები უკვე იტვირთება, LoadFilteredSchedule გადავახტეთ")
                Return
            End If

            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule - გვერდი {currentPage}, isNavigating: {isNavigating}")

            ' 🔧 მხოლოდ ნავიგაციის გარეშე ვაყენოთ isLoadingData
            If Not isNavigating Then
                isLoadingData = True
            End If

            Try
                ' ფილტრის კრიტერიუმების მიღება
                Dim criteria = filterManager.GetFilterCriteria()
                Dim pageSize = filterManager.GetPageSize()

                ' მონაცემების მიღება
                Dim result = dataProcessor.GetFilteredSchedule(criteria, currentPage, pageSize)

                ' UI-ს განახლება
                uiManager.LoadDataToGrid(result.Data, userRoleID)
                uiManager.UpdatePageLabel(result.CurrentPage, result.TotalPages)
                uiManager.UpdateNavigationButtons(result.CurrentPage, result.TotalPages)

                currentPage = result.CurrentPage

                ' 🔧 სტატისტიკა მხოლოდ ნავიგაციის გარეშე
                If Not isNavigating Then
                    System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAsync(criteria))
                End If

                Debug.WriteLine($"UC_Schedule: ჩატვირთულია {result.Data.Count} ჩანაწერი გვერდზე {currentPage}")

            Finally
                ' 🔧 isLoadingData მხოლოდ ნავიგაციის გარეშე ვანულებთ
                If Not isNavigating Then
                    isLoadingData = False
                End If
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule შეცდომა: {ex.Message}")
            isLoadingData = False
            MessageBox.Show($"მონაწილეობის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სტატისტიკის ასინქრონული განახლება
    ''' </summary>
    Private Sub UpdateStatisticsAsync(criteria As ScheduleDataProcessor.FilterCriteria)
        Try
            ' ყველა გაფილტრული მონაცემის მიღება
            Dim allFilteredResult = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
            Dim allFilteredData = allFilteredResult.Data

            ' UI Thread-ზე განახლება
            Me.Invoke(Sub()
                          Try
                              ' სტატისტიკის განახლება
                              If statisticsDisplayService IsNot Nothing Then
                                  statisticsDisplayService.UpdateStatisticsAsync(allFilteredData)
                              End If

                              ' ფინანსური ანალიზის განახლებადა (მხოლოდ ადმინისა და მენეჯერისთვის)
                              If financialAnalysisService IsNot Nothing AndAlso (userRoleID = 1 OrElse userRoleID = 2) Then
                                  financialAnalysisService.UpdateFinancialData(allFilteredData)
                              End If

                          Catch ex As Exception
                              Debug.WriteLine($"UC_Schedule: სტატისტიკის განახლების შეცდომა: {ex.Message}")
                          End Try
                      End Sub)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateStatisticsAsync შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები"

    ''' <summary>
    ''' 🔧 გამარტივებული ფილტრის შეცვლის ივენტი
    ''' ნავიგაციის დროს გვერდი არ რესეტდება
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine($"UC_Schedule: OnFilterChanged - isNavigating: {isNavigating}, isLoadingData: {isLoadingData}")

            ' 🔧 ნავიგაციის დროს ფილტრის ივენტი არ უნდა გავარდეს
            If isNavigating Then
                Debug.WriteLine("UC_Schedule: ნავიგაციის დროს FilterChanged ივენთი იგნორირებულია")
                Return
            End If

            ' 🔧 მონაცემების დატვირთვის დროს არ ვუშვებთ ციკლს
            If isLoadingData Then
                Debug.WriteLine("UC_Schedule: მონაცემების დატვირთვის დროს FilterChanged ივენთი იგნორირებულია")
                Return
            End If

            ' ფილტრის შეცვლისას გვერდი ресetდება 1-ზე
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვერდის ზომის შეცვლის ივენტი - შესწორებული ვერსია
    ''' </summary>
    Private Sub OnPageSizeChanged()
        Try
            Debug.WriteLine("UC_Schedule: OnPageSizeChanged - გვერდის ზომა შეიცვალა, პირველ გვერდზე გადასვლა")
            currentPage = 1 ' 🔧 გვერდის ზომის ცვლილებისას ყოველთვის პირველ გვერდზე
            LoadFilteredSchedule()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPageSizeChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის უჯრაზე დაჭერის ივენთი - დამატებულია ბაგის პრევენცია ორმაგ გახსნაზე
    ''' </summary>
    Private Sub OnDataGridViewCellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' არ ვმუშავებთ Header ან არასწორ ინდექსებს
            If e Is Nothing OrElse e.RowIndex < 0 Then Return

            ' რედაქტირების ღილაკზე დაჭერის შემოწმება
            If uiManager Is Nothing OrElse Not uiManager.IsEditButtonClicked(e) Then Return

            ' დუბლიური გახსნის პრევენცია
            If _editDialogOpen OrElse (DateTime.Now - _lastEditDialogClosed).TotalMilliseconds < 300 Then
                Debug.WriteLine("UC_Schedule: რედაქტირების ფორმა უკვე იხსნება ან ახლახან დაიხურა")
                Return
            End If

            Dim sessionIdValue = uiManager.GetCellValue("N", e.RowIndex)
            If sessionIdValue Is Nothing OrElse Not IsNumeric(sessionIdValue) Then Return

            Dim sessionId As Integer = CInt(sessionIdValue)
            Debug.WriteLine($"UC_Schedule: რედაქტირება - სესია ID={sessionId}")

            If dataService Is Nothing Then
                MessageBox.Show("მონაცემათა სერვისი არ არის ინციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _editDialogOpen = True
            Try
                Using editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_Schedule")
                    Dim result As DialogResult = editForm.ShowDialog()
                    If result = DialogResult.OK Then
                        RefreshData()
                        MessageBox.Show($"სესია ID={sessionId} წარმატებით განახლდა", "წარმატება",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            Catch formEx As Exception
                Debug.WriteLine($"UC_Schedule: რედაქტირების ფორმის შეცდომა: {formEx.Message}")
                MessageBox.Show($"რედაქტირების ფორმის გახსნის შეცდომა: {formEx.Message}", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                _editDialogOpen = False
                _lastEditDialogClosed = DateTime.Now
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnDataGridViewCellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ' ========================================
    ' 🔧 UC_Schedule.vb-ში შეცვალეთ ეს მიდგომები:
    ' ========================================

    ''' <summary>
    ''' წინა გვერდის ღილაკი - შესწორებული ვერსია
    ''' </summary>
    Private Sub OnPreviousPageClick(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine($"UC_Schedule: OnPreviousPageClick - მიმდინარე გვერდი: {currentPage}")

            If currentPage > 1 AndAlso Not isLoadingData Then
                ' 🔧 ნავიგაციის Flag-ი
                isNavigating = True
                Try
                    currentPage -= 1
                    LoadFilteredSchedule()
                    Debug.WriteLine($"UC_Schedule: წარმატებით გადავდი გვერდზე {currentPage}")
                Finally
                    isNavigating = False
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPreviousPageClick შეცდომა: {ex.Message}")
            isNavigating = False
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკი - შესწორებული ვერსია
    ''' </summary>
    Private Sub OnNextPageClick(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine($"UC_Schedule: OnNextPageClick - მიმდინარე გვერდი: {currentPage}")

            If BtnNext.Enabled AndAlso Not isLoadingData Then
                ' 🔧 ნავიგაციის Flag-ი
                isNavigating = True
                Try
                    currentPage += 1
                    LoadFilteredSchedule()
                    Debug.WriteLine($"UC_Schedule: წარმატებით გადავდი გვერდზე {currentPage}")
                Finally
                    isNavigating = False
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnNextPageClick შეცდომა: {ex.Message}")
            isNavigating = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 წინა გვერდის ღილაკის პირდაპირ ივენთი (Designer-დან)
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs) Handles BtnPrev.Click
        OnPreviousPageClick(sender, e)
    End Sub

    ''' <summary>
    ''' 🔧 შემდეგი გვერდის ღილაკის პირდაპირ ივენთი (Designer-დან)
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        OnNextPageClick(sender, e)
    End Sub



    ''' <summary>
    ''' განახლების ღილაკი
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            RefreshData()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtnRef_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ახალი ჩანაწერის დამატების ღილაკი
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Debug.WriteLine("BtnAddSchedule_Click: ახალი ჩანაწერის დამატების მოთხოვნა")
        Try
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    Debug.WriteLine("BtnAddSchedule_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციაალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("BtnAddSchedule_Click: dataService არ არის ინტიალიზებული")
                Return
            End If
            Dim recordType As String = "სესია"
            Dim newRecordForm As New NewRecordForm(dataService, recordType, userEmail, "UC_Calendar")
            Dim result = newRecordForm.ShowDialog()
            If result = DialogResult.OK Then
                Debug.WriteLine("BtnAddSchedule_Click: სესია წარმატებით ემატება")
                RefreshData()
            End If
        Catch ex As Exception
            Debug.WriteLine($"BtnAddSchedule_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ბეჭდვის ღილაკი
    ''' </summary>
    Private Sub BtbPrint_Click(sender As Object, e As EventArgs) Handles btbPrint.Click
        Try
            Debug.WriteLine("UC_Schedule: გაუმჯობესებული ბეჭდვის ღილაკზე დაჭერა")
            If DgvSchedule Is Nothing OrElse DgvSchedule.Rows.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            Dim printTypeResult As DialogResult = MessageBox.Show(
                "რომელი ტიპის ბეჭდვა გსურთ?" & Environment.NewLine & Environment.NewLine &
                "დიახ - გაუმჯობესებული ბეჭდვა (სვეტების მონიშვნა + ლანდშაფტი)" & Environment.NewLine &
                "არა - ჩვეულებრივი ბეჭდვა" & Environment.NewLine &
                "გაუქმება - ოპერაციის შეწყვეტა",
                "ბეჭდვის ტიპის არჩევა", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case printTypeResult
                Case DialogResult.Yes
                    Using advancedPrintService As New AdvancedDataGridViewPrintService(DgvSchedule)
                        advancedPrintService.ShowFullPrintDialog()
                    End Using
                Case DialogResult.No
                    Using simplePrintService As New DataGridViewPrintService(DgvSchedule)
                        Dim result As DialogResult = MessageBox.Show("გსურთ ჯერ ნახოთ ბეჭდვის პრევიუ?", "ბეჭდვის პრევიუ", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        Select Case result
                            Case DialogResult.Yes : simplePrintService.ShowPrintPreview()
                            Case DialogResult.No : simplePrintService.Print()
                        End Select
                    End Using
                Case DialogResult.Cancel
                    Debug.WriteLine("UC_Schedule: ბეჭდვა გაუქმებულია")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtbPrint_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' PDF ექსპორტის ღილაკი
    ''' </summary>
    Private Sub btnToPDF_Click(sender As Object, e As EventArgs) Handles btnToPDF.Click
        Try
            Debug.WriteLine("UC_Schedule: PDF ექსპორტი (გამარტივებული) ჯერ არ არის ოპტიმიზირებული ამ ცვლილებაში")
            MessageBox.Show("PDF ექსპორტის სრული ოპტიმიზაცია მოგვიანებით", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: btnToPDF_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

End Class