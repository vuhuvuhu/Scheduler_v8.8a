' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' ოპტიმიზირებული განრიგის UserControl (განახლებული ვერსია)
' დინამიური ფილტრაცია: ComboBox-ები შეიცავს მხოლოდ შერჩეულ პერიოდში არსებულ მნიშვნელობებს
' ბენეფიციარის გვარი ხსნის მხოლოდ სახელის არჩევის შემდეგ
' ===========================================
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Schedule
    Inherits UserControl

#Region "ველები და თვისებები"

    ' სერვისები - ოპტიმიზირებული არქიტექტურა
    Private dataService As IDataService = Nothing
    Private dataProcessor As ScheduleDataProcessor = Nothing
    Private uiManager As ScheduleUIManager = Nothing
    Private filterManager As ScheduleFilterManager = Nothing
    Private statisticsService As ScheduleStatisticsService = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6 ' ნაგულისხმევი როლი

    ' გვერდების მონაცემები
    Private currentPage As Integer = 1

    ''' <summary>
    ''' ნავიგაციის ღილაკების დაცვა ორმაგი კლიკისგან
    ''' </summary>
    Private isNavigating As Boolean = False
    Private lastNavigationTime As DateTime = DateTime.MinValue
    Private Const NAVIGATION_DELAY_MS As Integer = 500 ' 500 მილიწამი დაცვა

    ' სტატისტიკის თვისებები - თავდაპირველი API-ის შენარჩუნებისთვის
    ''' <summary>მიმდინარე გვერდის ნომერი</summary>
    Public ReadOnly Property CurrentPageNumber As Integer
        Get
            Return currentPage
        End Get
    End Property

    ''' <summary>სულ გვერდების რაოდენობა</summary>
    Public ReadOnly Property TotalPagesCount As Integer
        Get
            Try
                ' მიმდინარე ფილტრის კრიტერიუმებით გვერდების რაოდენობის მიღება
                If dataProcessor IsNot Nothing AndAlso filterManager IsNot Nothing Then
                    Dim criteria = filterManager.GetFilterCriteria()
                    Dim pageSize = filterManager.GetPageSize()
                    Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, pageSize)
                    Return result.TotalPages
                End If
                Return 1
            Catch
                Return 1
            End Try
        End Get
    End Property

    ''' <summary>მიმდინარე მომხმარებლის როლი</summary>
    Public ReadOnly Property CurrentUserRole As Integer
        Get
            Return userRoleID
        End Get
    End Property

    ''' <summary>მონიშნული სტატუსების რაოდენობა</summary>
    Public ReadOnly Property SelectedStatusCount As Integer
        Get
            If filterManager IsNot Nothing Then
                Return filterManager.SelectedStatusCount
            Else
                Return 0
            End If
        End Get
    End Property

#End Region
#Region "საჯარო მეთოდები - თავდაპირველი API"

    ''' <summary>
    ''' მონაცემთა სერვისისა და მომხმარებლის ინფორმაციის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub SetDataService(service As IDataService, Optional email As String = "", Optional role As Integer = 6)
        Try
            Debug.WriteLine($"UC_Schedule: მონაცემთა სერვისის დაყენება - email={email}, role={role}")

            dataService = service
            userEmail = email
            userRoleID = role

            ' როლის ავტომატური განსაზღვრა dataService-დან თუ საჭიროა
            If String.IsNullOrEmpty(email) OrElse role = 6 Then
                If dataService IsNot Nothing AndAlso Not String.IsNullOrEmpty(email) Then
                    Try
                        Dim userRole As String = dataService.GetUserRole(email)
                        If Not String.IsNullOrEmpty(userRole) Then
                            If Integer.TryParse(userRole, userRoleID) Then
                                Debug.WriteLine($"UC_Schedule: როლი განისაზღვრა dataService-დან: {userRoleID}")
                            End If
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"UC_Schedule: როლის განსაზღვრის შეცდომა: {ex.Message}")
                    End Try
                End If
            End If

            ' სერვისების ინიციალიზაცია
            If dataService IsNot Nothing Then
                InitializeServices()

                ' თუ კონტროლი უკვე ჩატვირთულია, ვტვირთავთ მონაცემებს
                If Me.IsHandleCreated Then
                    LoadFilteredSchedule()
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetDataService შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების ხელახალი ჩატვირთვა - თავდაპირველი API
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("UC_Schedule: მონაცემების განახლება")

            ' ქეშის გასუფთავება - SheetDataService-ისთვის
            If TypeOf dataService Is SheetDataService Then
                DirectCast(dataService, SheetDataService).InvalidateAllCache()
            End If

            ' ოპტიმიზირებული ქეშის გასუფთავება
            dataProcessor?.ClearCache()

            ' ComboBox-ების განახლება ახალი მონაცემებით
            If filterManager IsNot Nothing AndAlso dataProcessor IsNot Nothing Then
                filterManager.PopulateFilterComboBoxes(dataProcessor)
            End If

            ' ახლიდან ჩავტვირთოთ მონაცემები
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: RefreshData შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სტატუსის CheckBox-ის მონიშვნა/განიშვნა - თავდაპირველი API
    ''' </summary>
    ''' <param name="checkAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
    Public Sub SetAllStatusCheckBoxes(checkAll As Boolean)
        Try
            Debug.WriteLine($"UC_Schedule: ყველა CheckBox-ის დაყენება: {checkAll}")
            filterManager?.SetAllStatusCheckBoxes(checkAll)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetAllStatusCheckBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონკრეტული სტატუსის CheckBox-ის მონიშვნა/განიშვნა - თავდაპირველი API
    ''' </summary>
    ''' <param name="statusText">სტატუსის ტექსტი</param>
    ''' <param name="isChecked">მონიშნული თუ არა</param>
    Public Sub SetStatusCheckBox(statusText As String, isChecked As Boolean)
        Try
            filterManager?.SetStatusCheckBox(statusText, isChecked)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetStatusCheckBox შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონკრეტულ გვერდზე გადასვლა - თავდაპირველი API
    ''' </summary>
    ''' <param name="pageNumber">გვერდის ნომერი</param>
    Public Sub GoToPage(pageNumber As Integer)
        Try
            If pageNumber >= 1 Then
                currentPage = pageNumber
                Debug.WriteLine($"UC_Schedule: გადასვლა გვერდზე {currentPage}")
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: GoToPage შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრების განახლება - თავდაპირველი API
    ''' </summary>
    Public Sub UpdateFilters()
        Try
            Debug.WriteLine("UC_Schedule: ფილტრების განახლება")
            currentPage = 1
            LoadFilteredSchedule()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateFilters შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მომხმარებლის როლის ძაღლით დაყენება - ტესტირებისთვის
    ''' </summary>
    ''' <param name="roleId">როლის ID</param>
    Public Sub ForceSetUserRole(roleId As Integer)
        Try
            Debug.WriteLine($"UC_Schedule: ძაღლით დაყენება userRoleID = {roleId}")
            userRoleID = roleId

            ' UI-ს ხელახალი კონფიგურაცია ახალი როლისთვის
            If uiManager IsNot Nothing Then
                LoadFilteredSchedule() ' ეს ახლიდან შექმნის სვეტებს
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ForceSetUserRole შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region
#Region "პირადი მეთოდები - ოპტიმიზირებული იმპლემენტაცია"

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_Schedule: კონტროლის ჩატვირთვა")

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვსრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: Load შეცდომა: {ex.Message}")
            MessageBox.Show($"კონტროლის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სერვისების ინიციალიზაცია - ოპტიმიზირებული არქიტექტურა
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

            ' სტატისტიკის სერვისი
            statisticsService = New ScheduleStatisticsService(dataService)

            ' სტატუსის CheckBox-ების ძებნა და დამატება
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ისენთების მიბმა
            BindEvents()

            ' ✅ მრგვალი ღილაკები:
            ConfigureRoundButtons()

            ' ✅ ფინანსური სტატისტიკის განახლება:
            UpdateFinancialSummary()

            ' ComboBox-ების შევსება ფონზე (პერფორმანსისთვის) - განახლებული ლოგიკა
            System.Threading.Tasks.Task.Run(Sub() PopulateComboBoxesAsync())

            Debug.WriteLine("UC_Schedule: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: InitializeServices შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ღილაკების მრგვალი ფორმის დაყენება - არსებული ღილაკების მოდიფიკაცია
    ''' </summary>
    Private Sub ConfigureRoundButtons()
        Try
            Debug.WriteLine("UC_Schedule: ღილაკების მრგვალი ფორმის კონფიგურაცია")

            ' ყველა ღილაკის სია
            Dim buttons As Button() = {BtnRef, BtnAddSchedule, btbPrint, btnToPDF, BtnNext, BtnPrev}

            ' ყოველი ღილაკისთვის მრგვალი ფორმის დაყენება
            For Each btn As Button In buttons
                If btn IsNot Nothing Then
                    MakeButtonRound(btn)
                    Debug.WriteLine($"UC_Schedule: ღილაკი '{btn.Name}' გახდა მრგვალი")
                End If
            Next

            Debug.WriteLine("UC_Schedule: ყველა ღილაკი მრგვალი ფორმისაა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ConfigureRoundButtons შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ღილაკის მრგვალი ფორმის შექმნა - ზომის შეცვლის გარეშე
    ''' </summary>
    ''' <param name="button">ღილაკი რომელიც უნდა გახდეს მრგვალი</param>
    Private Sub MakeButtonRound(button As Button)
        Try
            If button Is Nothing Then Return

            ' GraphicsPath-ის შექმნა წრიული ფორმისთვის (არსებული ზომებით)
            Dim path As New Drawing2D.GraphicsPath()
            path.AddEllipse(0, 0, button.Width, button.Height)

            ' Region-ის დაყენება ღილაკისთვის - ეს ხდის მრგვალს
            button.Region = New Region(path)

            Debug.WriteLine($"UC_Schedule: ღილაკი '{button.Name}' - ზომა: {button.Width}x{button.Height}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: MakeButtonRound შეცდომა ღილაკისთვის '{button?.Name}': {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ComboBox-ების ასინქრონული შევსება - განახლებული ლოგიკა
    ''' </summary>
    Private Sub PopulateComboBoxesAsync()
        Try
            ' UI Thread-ზე გადასვლა
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
    ''' სტატუსის CheckBox-ების ძებნა და ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            ' CheckBox-ების ძებნა კონტროლში
            Dim statusCheckBoxes As New List(Of CheckBox)

            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBoxRecursive(Me, $"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    statusCheckBoxes.Add(checkBox)
                    Debug.WriteLine($"UC_Schedule: ნაპოვნია CheckBox{i}: '{checkBox.Text}'")
                End If
            Next

            ' CheckBox-ების ინიციალიზაცია ფილტრის მენეჯერში
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
            ' ჯერ პირდაპირ შვილებში ვეძებთ
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is CheckBox AndAlso ctrl.Name = name Then
                    Return DirectCast(ctrl, CheckBox)
                End If
            Next

            ' მერე რეკურსიულად შვილების შვილებში
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
    ''' ივენთების მიბმა - განახლებული ლოგიკა
    ''' </summary>
    Private Sub BindEvents()
        Try
            ' ფილტრების ივენთები
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged
            AddHandler filterManager.PageSizeChanged, AddressOf OnPageSizeChanged

            ' DataGridView-ის ივენთები
            AddHandler DgvSchedule.CellClick, AddressOf OnDataGridViewCellClick

            ' ნავიგაციის ღილაკების ივენთები
            AddHandler BtnPrev.Click, AddressOf OnPreviousPageClick
            AddHandler BtnNext.Click, AddressOf OnNextPageClick

            Debug.WriteLine("UC_Schedule: ივენთები მიბმულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დაყენება
    ''' </summary>
    Private Sub SetBackgroundColors()
        Try
            ' ღია გამჭვირვალე თეთრი ფერი
            Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

            ' პანელებისა და გრუპბოქსების ფერები
            If pnlFilter IsNot Nothing Then pnlFilter.BackColor = transparentWhite
            If GBSumInf IsNot Nothing Then GBSumInf.BackColor = transparentWhite
            If GBSumFin IsNot Nothing Then GBSumFin.BackColor = transparentWhite

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetBackgroundColors შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' განრიგის მონაცემების ჩატვირთვა - ოპტიმიზირებული ვერსია
    ''' </summary>
    Public Sub LoadFilteredSchedule()
        Try
            If dataProcessor Is Nothing OrElse uiManager Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_Schedule: სერვისები არ არის ინიციალიზებული")
                Return
            End If

            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule - გვერდი {currentPage}")

            ' ფილტრის კრიტერიუმების მიღება
            Dim criteria = filterManager.GetFilterCriteria()
            Dim pageSize = filterManager.GetPageSize()

            ' მონაცემების მიღება ოპტიმიზირებული პროცესორით
            Dim result = dataProcessor.GetFilteredSchedule(criteria, currentPage, pageSize)

            ' UI-ს განახლება
            uiManager.LoadDataToGrid(result.Data, userRoleID)
            uiManager.UpdatePageLabel(result.CurrentPage, result.TotalPages)
            uiManager.UpdateNavigationButtons(result.CurrentPage, result.TotalPages)

            ' მიმდინარე გვერდის განახლება
            currentPage = result.CurrentPage
            'ინფოპანელის განახლება
            UpdateGBSumInfAfterLoad()
            ' სტატისტიკის განახლება (ასინქრონულად)
            System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAsync(criteria))

            Debug.WriteLine($"UC_Schedule: ჩატვირთულია {result.Data.Count} ჩანაწერი, გვერდი {result.CurrentPage}/{result.TotalPages}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სტატისტიკის ასინქრონული განახლება
    ''' </summary>
    Private Sub UpdateStatisticsAsync(criteria As ScheduleDataProcessor.FilterCriteria)
        Try
            ' სწრაფი სტატისტიკის მიღება
            Dim quickStats = statisticsService.GetQuickStatistics(criteria)

            ' UI Thread-ზე სტატისტიკის განახლება
            Me.Invoke(Sub()
                          Try
                              ' სტატისტიკის ჩვენება UI-ში (GBSumInf და GBSumFin პანელებში)
                              UpdateSummaryPanels(quickStats.TotalSessions, quickStats.CompletedSessions, quickStats.TotalRevenue)
                          Catch ex As Exception
                              Debug.WriteLine($"UC_Schedule: სტატისტიკის UI განახლების შეცდომა: {ex.Message}")
                          End Try
                      End Sub)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateStatisticsAsync შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region
#Region "ივენთ ჰენდლერები"

    ''' <summary>
    ''' ფილტრის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine("UC_Schedule: ფილტრი შეიცვალა")
            currentPage = 1
            LoadFilteredSchedule()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვერდის ზომის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnPageSizeChanged()
        Try
            Debug.WriteLine("UC_Schedule: გვერდის ზომა შეიცვალა")
            currentPage = 1
            LoadFilteredSchedule()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPageSizeChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის უჯრაზე დაჭერის ივენთი
    ''' </summary>
    Private Sub OnDataGridViewCellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' შევამოწმოთ რედაქტირების ღილაკზე დაჭერა
            If uiManager.IsEditButtonClicked(e) Then
                ' მივიღოთ სესიის ID
                Dim sessionIdValue = uiManager.GetCellValue("N", e.RowIndex)

                If sessionIdValue IsNot Nothing AndAlso IsNumeric(sessionIdValue) Then
                    Dim sessionId As Integer = CInt(sessionIdValue)
                    Debug.WriteLine($"UC_Schedule: რედაქტირება - სესია ID={sessionId}")

                    ' NewRecordForm-ის გახსნა რედაქტირების რეჟიმში - თავდაპირველი ლოგიკა
                    Try
                        Dim currentUserEmail As String = If(String.IsNullOrEmpty(userEmail), "user@example.com", userEmail)

                        Using editForm As New NewRecordForm(dataService, "სესია", sessionId, currentUserEmail, "UC_Schedule")
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
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnDataGridViewCellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკზე დაჭერის ივენთი - დაცული ვერსია
    ''' </summary>
    Private Sub OnPreviousPageClick(sender As Object, e As EventArgs)
        Try
            ' დაცვა ორმაგი კლიკისგან
            If Not CanNavigate() Then
                Debug.WriteLine("UC_Schedule: წინა გვერდი - ნავიგაცია დაიცავა (ძალიან სწრაფი კლიკი)")
                Return
            End If

            If currentPage > 1 Then
                ' ნავიგაციის ფლაგის დაყენება
                SetNavigating(True)

                currentPage -= 1
                Debug.WriteLine($"UC_Schedule: წინა გვერდი - {currentPage}")

                ' მონაცემების ჩატვირთვა
                LoadFilteredSchedule()

                ' ნავიგაციის ფლაგის გათიშვა 200 მილიწამის შემდეგ
                System.Threading.Tasks.Task.Delay(200).ContinueWith(Sub() SetNavigating(False))
            Else
                Debug.WriteLine("UC_Schedule: წინა გვერდი - უკვე პირველ გვერდზე ვართ")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPreviousPageClick შეცდომა: {ex.Message}")
            SetNavigating(False) ' შეცდომის შემთხვევაში ფლაგის გათიშვა
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკზე დაჭერის ივენთი - დაცული ვერსია
    ''' </summary>
    Private Sub OnNextPageClick(sender As Object, e As EventArgs)
        Try
            ' დაცვა ორმაგი კლიკისგან
            If Not CanNavigate() Then
                Debug.WriteLine("UC_Schedule: შემდეგი გვერდი - ნავიგაცია დაიცავა (ძალიან სწრაფი კლიკი)")
                Return
            End If

            If BtnNext.Enabled Then
                ' ნავიგაციის ფლაგის დაყენება
                SetNavigating(True)

                currentPage += 1
                Debug.WriteLine($"UC_Schedule: შემდეგი გვერდი - {currentPage}")

                ' მონაცემების ჩატვირთვა
                LoadFilteredSchedule()

                ' ნავიგაციის ფლაგის გათიშვა 200 მილიწამის შემდეგ
                System.Threading.Tasks.Task.Delay(200).ContinueWith(Sub() SetNavigating(False))
            Else
                Debug.WriteLine("UC_Schedule: შემდეგი გვერდი - ღილაკი არ არის აქტიური")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnNextPageClick შეცდომა: {ex.Message}")
            SetNavigating(False) ' შეცდომის შემთხვევაში ფლაგის გათიშვა
        End Try
    End Sub

    ''' <summary>
    ''' შემოწმება შეიძლება თუ არა ნავიგაცია ამ მომენტში
    ''' </summary>
    Private Function CanNavigate() As Boolean
        Try
            ' თუ ნავიგაცია მიმდინარეობს
            If isNavigating Then
                Return False
            End If

            ' დროის შემოწმება - ბოლო ნავიგაციიდან გასული დრო
            Dim timeSinceLastNavigation = DateTime.Now.Subtract(lastNavigationTime).TotalMilliseconds
            If timeSinceLastNavigation < NAVIGATION_DELAY_MS Then
                Return False
            End If

            Return True

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CanNavigate შეცდომა: {ex.Message}")
            Return True ' შეცდომის შემთხვევაში ნავიგაციის ნებართვა
        End Try
    End Function

    ''' <summary>
    ''' ნავიგაციის ფლაგის დაყენება thread-safe რეჟიმში
    ''' </summary>
    Private Sub SetNavigating(value As Boolean)
        Try
            isNavigating = value

            If value Then
                ' ნავიგაციის დაწყების დროის შენახვა
                lastNavigationTime = DateTime.Now

                ' ღილაკების დროებით გათიშვა
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() DisableNavigationButtons())
                Else
                    DisableNavigationButtons()
                End If
            Else
                ' ღილაკების ხელახალი გააქტიურება
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() EnableNavigationButtons())
                Else
                    EnableNavigationButtons()
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetNavigating შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების დროებითი გათიშვა
    ''' </summary>
    Private Sub DisableNavigationButtons()
        Try
            If BtnPrev IsNot Nothing Then BtnPrev.Enabled = False
            If BtnNext IsNot Nothing Then BtnNext.Enabled = False
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: DisableNavigationButtons შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების ხელახალი გააქტიურება
    ''' </summary>
    Private Sub EnableNavigationButtons()
        Try
            ' UI Manager-ს საშუალებით სწორი მდგომარეობის აღდგენა
            If uiManager IsNot Nothing Then
                Dim totalPages = TotalPagesCount
                uiManager.UpdateNavigationButtons(currentPage, totalPages)
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: EnableNavigationButtons შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region
#Region "ინფო გრუპბოქსი"
    ''' <summary>
    ''' GBSumInf გრუპბოქსის ყველა ლეიბლის განახლება გაფილტრული მონაცემებით
    ''' ეს არის ერთადერთი UpdateSummaryPanels მეთოდი რომელიც უნდა იყოს ფაილში
    ''' </summary>
    ''' <param name="totalSessions">სულ სესიების რაოდენობა</param>
    ''' <param name="completedSessions">შესრულებული სესიების რაოდენობა</param>
    ''' <param name="totalRevenue">სულ შემოსავალი</param>
    Private Sub UpdateSummaryPanels(totalSessions As Integer, completedSessions As Integer, totalRevenue As Decimal)
        Try
            Debug.WriteLine($"UC_Schedule: GBSumInf-ის სრული განახლება - სესიები: {totalSessions}, შემოსავალი: {totalRevenue:F2}")

            ' გაფილტრული მონაცემების მიღება დეტალური სტატისტიკისთვის
            If filterManager IsNot Nothing AndAlso dataProcessor IsNot Nothing Then
                Dim criteria = filterManager.GetFilterCriteria()
                Dim filteredData = GetFilteredSessionsForStatistics(criteria)

                ' ყველა სტატისტიკის გამოთვლა სტატუსების მიხედვით
                Dim stats = CalculateDetailedStatistics(filteredData)

                ' ყველა ლეიბლის განახლება
                UpdateAllStatisticsLabels(stats)

                Debug.WriteLine($"UC_Schedule: სრული სტატისტიკა გამოთვლილია - {filteredData.Count} სესია")
                If userRoleID = 1 OrElse userRoleID = 2 Then
                    CalculateAndUpdateFinancialStatistics()
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateSummaryPanels შეცდომა: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' დეტალური სტატისტიკის სტრუქტურა
    ''' </summary>
    Private Structure DetailedStatistics
        ' სულ სესიები
        Public TotalSessions As Integer
        Public TotalMinutes As Integer
        Public Total30Min As Double
        Public Total60Min As Double

        ' შესრულებული სესიები
        Public CompletedSessions As Integer
        Public CompletedMinutes As Integer
        Public Completed30Min As Double
        Public Completed60Min As Double

        ' აღდგენილი სესიები
        Public RestoredSessions As Integer
        Public RestoredMinutes As Integer
        Public Restored30Min As Double
        Public Restored60Min As Double

        ' გაცდენა არასაპატიო
        Public MissedUnexcusedSessions As Integer
        Public MissedUnexcusedMinutes As Integer
        Public MissedUnexcused30Min As Double
        Public MissedUnexcused60Min As Double

        ' გაცდენა საპატიო
        Public MissedExcusedSessions As Integer
        Public MissedExcusedMinutes As Integer
        Public MissedExcused30Min As Double
        Public MissedExcused60Min As Double

        ' პროგრამით გატარება
        Public AutoProcessedSessions As Integer
        Public AutoProcessedMinutes As Integer
        Public AutoProcessed30Min As Double
        Public AutoProcessed60Min As Double

        ' გაუქმებული
        Public CancelledSessions As Integer
        Public CancelledMinutes As Integer
        Public Cancelled30Min As Double
        Public Cancelled60Min As Double

        ' დაგეგმილი (ვადაში)
        Public PlannedInTimeSessions As Integer
        Public PlannedInTimeMinutes As Integer
        Public PlannedInTime30Min As Double
        Public PlannedInTime60Min As Double

        ' დაგეგმილი (ვადაგადაცილებული)
        Public PlannedOverdueSessions As Integer
        Public PlannedOverdueMinutes As Integer
        Public PlannedOverdue30Min As Double
        Public PlannedOverdue60Min As Double

        ' უნიკალური მონაწილეები
        Public UniqueBeneficiaries As Integer
        Public UniqueTherapists As Integer
    End Structure

    ''' <summary>
    ''' დეტალური სტატისტიკის გამოთვლა ყველა სტატუსისთვის
    ''' </summary>
    ''' <param name="sessions">გაფილტრული სესიების სია</param>
    ''' <returns>დეტალური სტატისტიკა</returns>
    Private Function CalculateDetailedStatistics(sessions As List(Of IList(Of Object))) As DetailedStatistics
        Try
            Dim stats As New DetailedStatistics()
            Dim currentTime As DateTime = DateTime.Now

            ' უნიკალური მონაწილეების სეტები
            Dim uniqueBeneficiaries As New HashSet(Of String)()
            Dim uniqueTherapists As New HashSet(Of String)()

            Debug.WriteLine($"CalculateDetailedStatistics: დაიწყო {sessions.Count} სესიის ანალიზი")

            For Each session In sessions
                Try
                    ' ძირითადი ვალიდაცია
                    If session.Count < 13 Then Continue For

                    ' სტატუსის მიღება (M სვეტი - ინდექსი 12)
                    Dim status As String = If(session(12) IsNot Nothing, session(12).ToString().Trim().ToLower(), "")

                    ' ხანგძლივობის მიღება (G სვეტი - ინდექსი 6) - ინლაინ გამოთვლა
                    Dim duration As Integer = 60 ' ნაგულისხმევი ხანგძლივობა
                    If session.Count > 6 AndAlso session(6) IsNot Nothing Then
                        Dim durationStr = session(6).ToString().Trim()
                        Dim parsedDuration As Integer = 0
                        If Integer.TryParse(durationStr, parsedDuration) AndAlso parsedDuration > 0 Then
                            duration = parsedDuration
                        End If
                    End If

                    ' სესიის თარიღის მიღება (F სვეტი - ინდექსი 5) - ინლაინ გამოთვლა
                    Dim sessionDate As DateTime = DateTime.MinValue ' ნაგულისხმევი თარიღი
                    If session.Count > 5 AndAlso session(5) IsNot Nothing Then
                        Dim dateStr = session(5).ToString().Trim()
                        Dim parsedDate As DateTime
                        Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                        If DateTime.TryParseExact(dateStr, formats, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) OrElse
                       DateTime.TryParse(dateStr, parsedDate) Then
                            sessionDate = parsedDate
                        End If
                    End If

                    ' უნიკალური ბენეფიციარის დამატება (D და E სვეტები - ინდექსები 3,4)
                    If session.Count > 4 Then
                        Dim beneficiaryName = $"{session(3)?.ToString().Trim()} {session(4)?.ToString().Trim()}"
                        If Not String.IsNullOrWhiteSpace(beneficiaryName.Trim()) Then
                            uniqueBeneficiaries.Add(beneficiaryName.Trim())
                        End If
                    End If

                    ' უნიკალური თერაპევტის დამატება (I სვეტი - ინდექსი 8)
                    If session.Count > 8 AndAlso session(8) IsNot Nothing Then
                        Dim therapist = session(8).ToString().Trim()
                        If Not String.IsNullOrWhiteSpace(therapist) Then
                            uniqueTherapists.Add(therapist)
                        End If
                    End If

                    ' სულ სესიები
                    stats.TotalSessions += 1
                    stats.TotalMinutes += duration

                    ' სტატუსების მიხედვით კლასიფიკაცია
                    Select Case status
                        Case "შესრულებული"
                            stats.CompletedSessions += 1
                            stats.CompletedMinutes += duration

                        Case "აღდგენა"
                            stats.RestoredSessions += 1
                            stats.RestoredMinutes += duration

                        Case "გაცდენა არასაპატიო"
                            stats.MissedUnexcusedSessions += 1
                            stats.MissedUnexcusedMinutes += duration

                        Case "გაცდენა საპატიო"
                            stats.MissedExcusedSessions += 1
                            stats.MissedExcusedMinutes += duration

                        Case "პროგრამით გატარება"
                            stats.AutoProcessedSessions += 1
                            stats.AutoProcessedMinutes += duration

                        Case "გაუქმება", "გაუქმებული"
                            stats.CancelledSessions += 1
                            stats.CancelledMinutes += duration

                        Case "დაგეგმილი"
                            ' დაგეგმილი სესიების დაყოფა ვადაში/ვადაგადაცილებულად
                            If sessionDate <= currentTime Then
                                ' ვადაგადაცილებული
                                stats.PlannedOverdueSessions += 1
                                stats.PlannedOverdueMinutes += duration
                            Else
                                ' ვადაში
                                stats.PlannedInTimeSessions += 1
                                stats.PlannedInTimeMinutes += duration
                            End If
                    End Select

                Catch sessionEx As Exception
                    Debug.WriteLine($"CalculateDetailedStatistics: სესიის დამუშავების შეცდომა: {sessionEx.Message}")
                    Continue For
                End Try
            Next

            ' 30 და 60 წუთიანი ექვივალენტების გამოთვლა
            stats.Total30Min = Math.Round(stats.TotalMinutes / 30.0, 1)
            stats.Total60Min = Math.Round(stats.TotalMinutes / 60.0, 1)

            stats.Completed30Min = Math.Round(stats.CompletedMinutes / 30.0, 1)
            stats.Completed60Min = Math.Round(stats.CompletedMinutes / 60.0, 1)

            stats.Restored30Min = Math.Round(stats.RestoredMinutes / 30.0, 1)
            stats.Restored60Min = Math.Round(stats.RestoredMinutes / 60.0, 1)

            stats.MissedUnexcused30Min = Math.Round(stats.MissedUnexcusedMinutes / 30.0, 1)
            stats.MissedUnexcused60Min = Math.Round(stats.MissedUnexcusedMinutes / 60.0, 1)

            stats.MissedExcused30Min = Math.Round(stats.MissedExcusedMinutes / 30.0, 1)
            stats.MissedExcused60Min = Math.Round(stats.MissedExcusedMinutes / 60.0, 1)

            stats.AutoProcessed30Min = Math.Round(stats.AutoProcessedMinutes / 30.0, 1)
            stats.AutoProcessed60Min = Math.Round(stats.AutoProcessedMinutes / 60.0, 1)

            stats.Cancelled30Min = Math.Round(stats.CancelledMinutes / 30.0, 1)
            stats.Cancelled60Min = Math.Round(stats.CancelledMinutes / 60.0, 1)

            stats.PlannedInTime30Min = Math.Round(stats.PlannedInTimeMinutes / 30.0, 1)
            stats.PlannedInTime60Min = Math.Round(stats.PlannedInTimeMinutes / 60.0, 1)

            stats.PlannedOverdue30Min = Math.Round(stats.PlannedOverdueMinutes / 30.0, 1)
            stats.PlannedOverdue60Min = Math.Round(stats.PlannedOverdueMinutes / 60.0, 1)

            ' უნიკალური მონაწილეების რაოდენობა
            stats.UniqueBeneficiaries = uniqueBeneficiaries.Count
            stats.UniqueTherapists = uniqueTherapists.Count

            Debug.WriteLine($"CalculateDetailedStatistics: სტატისტიკა გამოთვლილია - ბენეფიციარები: {stats.UniqueBeneficiaries}, თერაპევტები: {stats.UniqueTherapists}")

            Return stats

        Catch ex As Exception
            Debug.WriteLine($"CalculateDetailedStatistics: შეცდომა: {ex.Message}")
            Return New DetailedStatistics()
        End Try
    End Function

    ''' <summary>
    ''' სესიის თარიღის მიღება
    ''' </summary>
    Private Function GetSessionDate(session As IList(Of Object)) As DateTime
        Try
            If session.Count > 5 AndAlso session(5) IsNot Nothing Then
                Dim dateStr = session(5).ToString().Trim()
                Dim sessionDate As DateTime

                Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                If DateTime.TryParseExact(dateStr, formats, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, sessionDate) OrElse
               DateTime.TryParse(dateStr, sessionDate) Then
                    Return sessionDate
                End If
            End If

            ' ნაგულისხმევი თარიღი
            Return DateTime.MinValue

        Catch
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>
    ''' ყველა სტატისტიკის ლეიბლის განახლება
    ''' </summary>
    Private Sub UpdateAllStatisticsLabels(stats As DetailedStatistics)
        Try
            Debug.WriteLine("UpdateAllStatisticsLabels: ყველა ლეიბლის განახლება დაიწყო")

            ' სულ სესიები (lsr, lsm, ls30, ls60)
            UpdateLabelSafely(lsr, stats.TotalSessions.ToString())
            UpdateLabelSafely(lsm, stats.TotalMinutes.ToString())
            UpdateLabelSafely(ls30, stats.Total30Min.ToString("F1"))
            UpdateLabelSafely(ls60, stats.Total60Min.ToString("F1"))

            ' შესრულებული სესიები (lshess, lshesm, lshes30, lshes60)
            UpdateLabelSafely(lshess, stats.CompletedSessions.ToString())
            UpdateLabelSafely(lshesm, stats.CompletedMinutes.ToString())
            UpdateLabelSafely(lshes30, stats.Completed30Min.ToString("F1"))
            UpdateLabelSafely(lshes60, stats.Completed60Min.ToString("F1"))

            ' აღდგენილი სესიები (las, lam, la30, la60)
            UpdateLabelSafely(las, stats.RestoredSessions.ToString())
            UpdateLabelSafely(lam, stats.RestoredMinutes.ToString())
            UpdateLabelSafely(la30, stats.Restored30Min.ToString("F1"))
            UpdateLabelSafely(la60, stats.Restored60Min.ToString("F1"))

            ' გაცდენა არასაპატიო (lgas, lgam, lga30, lga60)
            UpdateLabelSafely(lgas, stats.MissedUnexcusedSessions.ToString())
            UpdateLabelSafely(lgam, stats.MissedUnexcusedMinutes.ToString())
            UpdateLabelSafely(lga30, stats.MissedUnexcused30Min.ToString("F1"))
            UpdateLabelSafely(lga60, stats.MissedUnexcused60Min.ToString("F1"))

            ' გაცდენა საპატიო (lgss, lgsm, lgs30, lgs60)
            UpdateLabelSafely(lgss, stats.MissedExcusedSessions.ToString())
            UpdateLabelSafely(lgsm, stats.MissedExcusedMinutes.ToString())
            UpdateLabelSafely(lgs30, stats.MissedExcused30Min.ToString("F1"))
            UpdateLabelSafely(lgs60, stats.MissedExcused60Min.ToString("F1"))

            ' პროგრამით გატარება (lpgs, lpgm, lpg30, lpg60) - ვფიქრობ lpgm, lpg30, lpg60 იქნება
            UpdateLabelSafely(lpgs, stats.AutoProcessedSessions.ToString())
            UpdateLabelSafely(lpgm, stats.AutoProcessedMinutes.ToString())
            UpdateLabelSafely(lpg30, stats.AutoProcessed30Min.ToString("F1"))
            UpdateLabelSafely(lpg60, stats.AutoProcessed60Min.ToString("F1"))

            ' გაუქმებული (lgaus, lgaum, lgau30, lgau60) - ვფიქრობ ასე იქნება სახელები
            UpdateLabelSafely(lgs, stats.CancelledSessions.ToString())
            UpdateLabelSafely(lgm, stats.CancelledMinutes.ToString())
            UpdateLabelSafely(lg30, stats.Cancelled30Min.ToString("F1"))
            UpdateLabelSafely(lg60, stats.Cancelled60Min.ToString("F1"))

            ' დაგეგმილი ვადაში (lds, ldm, ld30, ld60)
            UpdateLabelSafely(lds, stats.PlannedInTimeSessions.ToString())
            UpdateLabelSafely(ldm, stats.PlannedInTimeMinutes.ToString())
            UpdateLabelSafely(ld30, stats.PlannedInTime30Min.ToString("F1"))
            UpdateLabelSafely(ld60, stats.PlannedInTime60Min.ToString("F1"))

            ' დაგეგმილი ვადაგადაცილებული (ldvs, ldvm, ldv30, ldv60)
            UpdateLabelSafely(ldvs, stats.PlannedOverdueSessions.ToString())
            UpdateLabelSafely(ldvm, stats.PlannedOverdueMinutes.ToString())
            UpdateLabelSafely(ldv30, stats.PlannedOverdue30Min.ToString("F1"))
            UpdateLabelSafely(ldv60, stats.PlannedOverdue60Min.ToString("F1"))

            ' უნიკალური მონაწილეები
            UpdateLabelSafely(lbenes, stats.UniqueBeneficiaries.ToString())
            UpdateLabelSafely(lpers, stats.UniqueTherapists.ToString())

            Debug.WriteLine($"UpdateAllStatisticsLabels: ყველა ლეიბლი განახლებულია")

        Catch ex As Exception
            Debug.WriteLine($"UpdateAllStatisticsLabels: შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ლეიბლის უსაფრთხო განახლება - არ იწყუება თუ ლეიბლი არ არსებობს
    ''' </summary>
    Private Sub UpdateLabelSafely(label As Label, value As String)
        Try
            If label IsNot Nothing Then
                label.Text = value
            Else
                Debug.WriteLine($"UpdateLabelSafely: ლეიბლი არ არსებობს, მნიშვნელობა: {value}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"UpdateLabelSafely: შეცდომა ლეიბლის განახლებისას: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' GBSumInf-ის ლეიბლების განახლება
    ''' </summary>
    ''' <param name="sessionCount">სესიების რაოდენობა</param>
    ''' <param name="totalMinutes">ჯამური წუთები</param>
    ''' <param name="sessions30">30-წუთიანი სესიების ეკვივალენტი</param>
    ''' <param name="sessions60">60-წუთიანი სესიების ეკვივალენტი</param>
    Private Sub UpdateGBSumInfLabels(sessionCount As Integer, totalMinutes As Integer, sessions30 As Double, sessions60 As Double)
        Try
            ' lsr - სეანსების რაოდენობა
            If lsr IsNot Nothing Then
                lsr.Text = sessionCount.ToString()
                Debug.WriteLine($"UC_Schedule: lsr = {sessionCount}")
            End If

            ' lsm - სეანსების ჯამური ხანგძლიობა (წუთებში)
            If lsm IsNot Nothing Then
                lsm.Text = totalMinutes.ToString()
                Debug.WriteLine($"UC_Schedule: lsm = {totalMinutes}")
            End If

            ' ls30 - lsm/30 (30-წუთიანი სესიების ეკვივალენტი)
            If ls30 IsNot Nothing Then
                ls30.Text = sessions30.ToString("F1")
                Debug.WriteLine($"UC_Schedule: ls30 = {sessions30:F1}")
            End If

            ' ls60 - lsm/60 (60-წუთიანი სესიების ეკვივალენტი) 
            If ls60 IsNot Nothing Then
                ls60.Text = sessions60.ToString("F1")
                Debug.WriteLine($"UC_Schedule: ls60 = {sessions60:F1}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateGBSumInfLabels შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გაფილტრული სესიების მონაცემების მიღება სტატისტიკისთვის
    ''' </summary>
    ''' <param name="criteria">ფილტრის კრიტერიუმები</param>
    ''' <returns>გაფილტრული სესიების სია</returns>
    Private Function GetFilteredSessionsForStatistics(criteria As ScheduleDataProcessor.FilterCriteria) As List(Of IList(Of Object))
        Try
            If dataProcessor Is Nothing Then
                Debug.WriteLine("UC_Schedule: dataProcessor არის Nothing")
                Return New List(Of IList(Of Object))()
            End If

            ' ყველა გაფილტრული სესიის მიღება (გვერდების გარეშე)
            Dim result = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)

            Debug.WriteLine($"UC_Schedule: გაფილტრული სესიები სტატისტიკისთვის: {result.Data.Count}")
            Return result.Data

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: GetFilteredSessionsForStatistics შეცდომა: {ex.Message}")
            Return New List(Of IList(Of Object))()
        End Try
    End Function

    ''' <summary>
    ''' სესიების ჯამური ხანგძლივობის გამოთვლა წუთებში
    ''' </summary>
    ''' <param name="sessions">სესიების სია</param>
    ''' <returns>ჯამური ხანგძლივობა წუთებში</returns>
    Private Function CalculateTotalMinutes(sessions As List(Of IList(Of Object))) As Integer
        Try
            Dim totalMinutes As Integer = 0

            For Each session In sessions
                Try
                    ' ხანგძლივობა G სვეტშია (ინდექსი 6)
                    If session.Count > 6 AndAlso session(6) IsNot Nothing Then
                        Dim durationStr = session(6).ToString().Trim()
                        Dim duration As Integer = 0

                        If Integer.TryParse(durationStr, duration) AndAlso duration > 0 Then
                            totalMinutes += duration
                        Else
                            ' თუ ხანგძლივობა არ არის მითითებული, ნაგულისხმევად 60 წუთი
                            totalMinutes += 60
                            Debug.WriteLine($"UC_Schedule: სესიისთვის გამოყენებულია ნაგულისხმევი ხანგძლივობა (60 წთ)")
                        End If
                    End If

                Catch sessionEx As Exception
                    Debug.WriteLine($"UC_Schedule: სესიის ხანგძლივობის გამოთვლის შეცდომა: {sessionEx.Message}")
                    ' შეცდომის შემთხვევაში ნაგულისხმევი 60 წუთი
                    totalMinutes += 60
                End Try
            Next

            Debug.WriteLine($"UC_Schedule: ჯამური ხანგძლივობა: {totalMinutes} წუთი")
            Return totalMinutes

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CalculateTotalMinutes შეცდომა: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' GBSumInf-ის განახლება LoadFilteredSchedule-ის შემდეგ
    ''' ეს მეთოდი უნდა გამოიძახოს LoadFilteredSchedule-ის ბოლოში
    ''' </summary>
    Private Sub UpdateGBSumInfAfterLoad()
        Try
            If filterManager IsNot Nothing AndAlso statisticsService IsNot Nothing Then
                Dim criteria = filterManager.GetFilterCriteria()

                ' ასინქრონული სტატისტიკის განახლება
                System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAsync(criteria))
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateGBSumInfAfterLoad შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region
#Region "ფინანსური გრუპბოქსი"
    ''' <summary>
    ''' GBSumFin პანელის ხილვადობის კონტროლი და ფინანსური სტატისტიკის განახლება
    ''' </summary>
    Private Sub UpdateFinancialSummary()
        Try
            Debug.WriteLine($"UC_Schedule: UpdateFinancialSummary - მომხმარებლის როლი: {userRoleID}")

            ' GBSumFin პანელის ხილვადობის კონტროლი
            If GBSumFin IsNot Nothing Then
                ' მხოლოდ ადმინი (1) და მენეჯერი (2) ხედავს ფინანსურ პანელს
                Dim shouldShowFinancial As Boolean = (userRoleID = 1 OrElse userRoleID = 2)
                GBSumFin.Visible = shouldShowFinancial

                Debug.WriteLine($"UC_Schedule: GBSumFin ხილვადობა: {shouldShowFinancial}")

                If shouldShowFinancial Then
                    ' CheckBox-ების ინიციალიზაცია და ივენთების მიბმა
                    InitializeFinancialCheckBoxes()

                    ' ფინანსური სტატისტიკის გამოთვლა და განახლება
                    CalculateAndUpdateFinancialStatistics()
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateFinancialSummary შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსური CheckBox-ების ინიციალიზაცია და ივენთების მიბმა
    ''' </summary>
    Private Sub InitializeFinancialCheckBoxes()
        Try
            Debug.WriteLine("UC_Schedule: ფინანსური CheckBox-ების ინიციალიზაცია")

            ' ყველა ფინანსური CheckBox-ის სია
            Dim financialCheckBoxes As CheckBox() = {cbs, cba, cbga, cbgs, cbpg, cbg, cbd, cbdv}

            For Each cb In financialCheckBoxes
                If cb IsNot Nothing Then
                    ' ნაგულისხმევად ყველა მონიშნული
                    cb.Checked = True

                    ' ივენთის მიბმა CheckedChanged-ისთვის
                    RemoveHandler cb.CheckedChanged, AddressOf FinancialCheckBox_CheckedChanged
                    AddHandler cb.CheckedChanged, AddressOf FinancialCheckBox_CheckedChanged

                    Debug.WriteLine($"UC_Schedule: CheckBox '{cb.Name}' ინიციალიზებული")
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: InitializeFinancialCheckBoxes შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსური CheckBox-ის CheckedChanged ივენთი
    ''' </summary>
    Private Sub FinancialCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("UC_Schedule: ფინანსური CheckBox შეიცვალა")

            ' ფინანსური სტატისტიკის ხელახალი გამოთვლა
            CalculateAndUpdateFinancialStatistics()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: FinancialCheckBox_CheckedChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსური სტატისტიკის გამოთვლა და განახლება
    ''' </summary>
    Private Sub CalculateAndUpdateFinancialStatistics()
        Try
            If filterManager Is Nothing OrElse dataProcessor Is Nothing Then
                Debug.WriteLine("UC_Schedule: filterManager ან dataProcessor არის Nothing")
                Return
            End If

            Debug.WriteLine("UC_Schedule: ფინანსური სტატისტიკის გამოთვლა დაიწყო")

            ' გაფილტრული მონაცემების მიღება
            Dim criteria = filterManager.GetFilterCriteria()
            Dim filteredData = GetFilteredSessionsForStatistics(criteria)

            ' ფინანსური სტატისტიკის გამოთვლა
            Dim financialStats = CalculateFinancialStatistics(filteredData)

            ' ლეიბლების განახლება
            UpdateFinancialLabels(financialStats)

            Debug.WriteLine("UC_Schedule: ფინანსური სტატისტიკა განახლებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CalculateAndUpdateFinancialStatistics შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსური სტატისტიკის სტრუქტურა
    ''' </summary>
    Private Structure FinancialStatistics
        ' შესრულებული
        Public CompletedPrivateAmount As Decimal
        Public CompletedOtherAmount As Decimal
        Public CompletedTotalAmount As Decimal

        ' აღდგენილი
        Public RestoredPrivateAmount As Decimal
        Public RestoredOtherAmount As Decimal
        Public RestoredTotalAmount As Decimal

        ' გაცდენა არასაპატიო
        Public MissedUnexcusedPrivateAmount As Decimal
        Public MissedUnexcusedOtherAmount As Decimal
        Public MissedUnexcusedTotalAmount As Decimal

        ' გაცდენა საპატიო
        Public MissedExcusedPrivateAmount As Decimal
        Public MissedExcusedOtherAmount As Decimal
        Public MissedExcusedTotalAmount As Decimal

        ' პროგრამით გატარება
        Public AutoProcessedPrivateAmount As Decimal
        Public AutoProcessedOtherAmount As Decimal
        Public AutoProcessedTotalAmount As Decimal

        ' გაუქმებული
        Public CancelledPrivateAmount As Decimal
        Public CancelledOtherAmount As Decimal
        Public CancelledTotalAmount As Decimal

        ' დაგეგმილი მომავალში
        Public PlannedInTimePrivateAmount As Decimal
        Public PlannedInTimeOtherAmount As Decimal
        Public PlannedInTimeTotalAmount As Decimal

        ' დაგეგმილი ვადაგადაცილებული
        Public PlannedOverduePrivateAmount As Decimal
        Public PlannedOverdueOtherAmount As Decimal
        Public PlannedOverdueTotalAmount As Decimal

        ' საერთო ჯამები (მონიშნული CheckBox-ების მიხედვით)
        Public TotalPrivateAmount As Decimal
        Public TotalOtherAmount As Decimal
        Public GrandTotalAmount As Decimal
    End Structure

    ''' <summary>
    ''' ფინანსური სტატისტიკის გამოთვლა
    ''' </summary>
    Private Function CalculateFinancialStatistics(sessions As List(Of IList(Of Object))) As FinancialStatistics
        Try
            Dim stats As New FinancialStatistics()
            Dim currentTime As DateTime = DateTime.Now

            Debug.WriteLine($"CalculateFinancialStatistics: {sessions.Count} სესიის ანალიზი")

            For Each session In sessions
                Try
                    ' ძირითადი ვალიდაცია
                    If session.Count < 14 Then Continue For

                    ' სტატუსის მიღება (M სვეტი - ინდექსი 12)
                    Dim status As String = If(session(12) IsNot Nothing, session(12).ToString().Trim().ToLower(), "")

                    ' ფასის მიღება (L სვეტი - ინდექსი 11)
                    Dim price As Decimal = GetSessionPrice(session)

                    ' დაფინანსების ტიპის მიღება (N სვეტი - ინდექსი 13)
                    Dim funding As String = If(session(13) IsNot Nothing, session(13).ToString().Trim().ToLower(), "")
                    Dim isPrivate As Boolean = (funding = "კერძო")

                    ' სესიის თარიღის მიღება (F სვეტი - ინდექსი 5)
                    Dim sessionDate As DateTime = DateTime.MinValue
                    If session.Count > 5 AndAlso session(5) IsNot Nothing Then
                        Dim dateStr = session(5).ToString().Trim()
                        Dim parsedDate As DateTime
                        Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                        If DateTime.TryParseExact(dateStr, formats, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) OrElse
                           DateTime.TryParse(dateStr, parsedDate) Then
                            sessionDate = parsedDate
                        End If
                    End If

                    ' სტატუსების მიხედვით თანხების დაჯამება
                    Select Case status
                        Case "შესრულებული"
                            If isPrivate Then
                                stats.CompletedPrivateAmount += price
                            Else
                                stats.CompletedOtherAmount += price
                            End If
                            stats.CompletedTotalAmount += price

                        Case "აღდგენა"
                            If isPrivate Then
                                stats.RestoredPrivateAmount += price
                            Else
                                stats.RestoredOtherAmount += price
                            End If
                            stats.RestoredTotalAmount += price

                        Case "გაცდენა არასაპატიო"
                            If isPrivate Then
                                stats.MissedUnexcusedPrivateAmount += price
                            Else
                                stats.MissedUnexcusedOtherAmount += price
                            End If
                            stats.MissedUnexcusedTotalAmount += price

                        Case "გაცდენა საპატიო"
                            If isPrivate Then
                                stats.MissedExcusedPrivateAmount += price
                            Else
                                stats.MissedExcusedOtherAmount += price
                            End If
                            stats.MissedExcusedTotalAmount += price

                        Case "პროგრამით გატარება"
                            If isPrivate Then
                                stats.AutoProcessedPrivateAmount += price
                            Else
                                stats.AutoProcessedOtherAmount += price
                            End If
                            stats.AutoProcessedTotalAmount += price

                        Case "გაუქმება", "გაუქმებული"
                            If isPrivate Then
                                stats.CancelledPrivateAmount += price
                            Else
                                stats.CancelledOtherAmount += price
                            End If
                            stats.CancelledTotalAmount += price

                        Case "დაგეგმილი"
                            ' დაგეგმილი სესიების დაყოფა ვადაში/ვადაგადაცილებულად
                            If sessionDate <= currentTime Then
                                ' ვადაგადაცილებული
                                If isPrivate Then
                                    stats.PlannedOverduePrivateAmount += price
                                Else
                                    stats.PlannedOverdueOtherAmount += price
                                End If
                                stats.PlannedOverdueTotalAmount += price
                            Else
                                ' ვადაში
                                If isPrivate Then
                                    stats.PlannedInTimePrivateAmount += price
                                Else
                                    stats.PlannedInTimeOtherAmount += price
                                End If
                                stats.PlannedInTimeTotalAmount += price
                            End If
                    End Select

                Catch sessionEx As Exception
                    Debug.WriteLine($"CalculateFinancialStatistics: სესიის დამუშავების შეცდომა: {sessionEx.Message}")
                    Continue For
                End Try
            Next

            ' მონიშნული CheckBox-ების მიხედვით საერთო ჯამების გამოთვლა
            CalculateTotalAmounts(stats)

            Debug.WriteLine($"CalculateFinancialStatistics: ფინანსური სტატისტიკა გამოთვლილია")
            Return stats

        Catch ex As Exception
            Debug.WriteLine($"CalculateFinancialStatistics: შეცდომა: {ex.Message}")
            Return New FinancialStatistics()
        End Try
    End Function

    ''' <summary>
    ''' სესიის ფასის მიღება
    ''' </summary>
    Private Function GetSessionPrice(session As IList(Of Object)) As Decimal
        Try
            If session.Count > 11 AndAlso session(11) IsNot Nothing Then
                Dim priceStr = session(11).ToString().Replace(",", ".").Trim()
                Dim price As Decimal = 0

                If Decimal.TryParse(priceStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, price) Then
                    Return price
                End If
            End If

            Return 0

        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' მონიშნული CheckBox-ების მიხედვით საერთო ჯამების გამოთვლა
    ''' </summary>
    Private Sub CalculateTotalAmounts(ByRef stats As FinancialStatistics)
        Try
            stats.TotalPrivateAmount = 0
            stats.TotalOtherAmount = 0

            ' თითოეული CheckBox-ის შემოწმება და შესაბამისი თანხების დამატება
            If CheckBoxIsChecked(cbs) Then ' შესრულებული
                stats.TotalPrivateAmount += stats.CompletedPrivateAmount
                stats.TotalOtherAmount += stats.CompletedOtherAmount
            End If

            If CheckBoxIsChecked(cba) Then ' აღდგენილი
                stats.TotalPrivateAmount += stats.RestoredPrivateAmount
                stats.TotalOtherAmount += stats.RestoredOtherAmount
            End If

            If CheckBoxIsChecked(cbga) Then ' გაცდენა არასაპატიო
                stats.TotalPrivateAmount += stats.MissedUnexcusedPrivateAmount
                stats.TotalOtherAmount += stats.MissedUnexcusedOtherAmount
            End If

            If CheckBoxIsChecked(cbgs) Then ' გაცდენა საპატიო
                stats.TotalPrivateAmount += stats.MissedExcusedPrivateAmount
                stats.TotalOtherAmount += stats.MissedExcusedOtherAmount
            End If

            If CheckBoxIsChecked(cbpg) Then ' პროგრამით გატარება
                stats.TotalPrivateAmount += stats.AutoProcessedPrivateAmount
                stats.TotalOtherAmount += stats.AutoProcessedOtherAmount
            End If

            If CheckBoxIsChecked(cbg) Then ' გაუქმებული
                stats.TotalPrivateAmount += stats.CancelledPrivateAmount
                stats.TotalOtherAmount += stats.CancelledOtherAmount
            End If

            If CheckBoxIsChecked(cbd) Then ' დაგეგმილი მომავალში
                stats.TotalPrivateAmount += stats.PlannedInTimePrivateAmount
                stats.TotalOtherAmount += stats.PlannedInTimeOtherAmount
            End If

            If CheckBoxIsChecked(cbdv) Then ' დაგეგმილი ვადაგადაცილებული
                stats.TotalPrivateAmount += stats.PlannedOverduePrivateAmount
                stats.TotalOtherAmount += stats.PlannedOverdueOtherAmount
            End If

            ' საერთო ჯამი
            stats.GrandTotalAmount = stats.TotalPrivateAmount + stats.TotalOtherAmount

        Catch ex As Exception
            Debug.WriteLine($"CalculateTotalAmounts: შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' CheckBox-ის მდგომარეობის უსაფრთხო შემოწმება
    ''' </summary>
    Private Function CheckBoxIsChecked(checkBox As CheckBox) As Boolean
        Try
            Return checkBox IsNot Nothing AndAlso checkBox.Checked
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ფინანსური ლეიბლების განახლება
    ''' </summary>
    Private Sub UpdateFinancialLabels(stats As FinancialStatistics)
        Try
            Debug.WriteLine("UpdateFinancialLabels: ფინანსური ლეიბლების განახლება")

            ' შესრულებული ლეიბლები და ფერების განახლება
            Dim isCompletedChecked = CheckBoxIsChecked(cbs)
            UpdateFinancialLabelWithColor(lsk, stats.CompletedPrivateAmount, isCompletedChecked)
            UpdateFinancialLabelWithColor(lsa, stats.CompletedOtherAmount, isCompletedChecked)
            UpdateFinancialLabelWithColor(lssum, stats.CompletedTotalAmount, isCompletedChecked)

            ' აღდგენილი ლეიბლები
            Dim isRestoredChecked = CheckBoxIsChecked(cba)
            UpdateFinancialLabelWithColor(lak, stats.RestoredPrivateAmount, isRestoredChecked)
            UpdateFinancialLabelWithColor(laa, stats.RestoredOtherAmount, isRestoredChecked)
            UpdateFinancialLabelWithColor(lasum, stats.RestoredTotalAmount, isRestoredChecked)

            ' გაცდენა არასაპატიო ლეიბლები (ვფიქრობ lgak, lgaa, lgasum იქნება სახელები)
            Dim isMissedUnexcusedChecked = CheckBoxIsChecked(cbga)
            UpdateFinancialLabelWithColor(lgak, stats.MissedUnexcusedPrivateAmount, isMissedUnexcusedChecked)
            UpdateFinancialLabelWithColor(lgaa, stats.MissedUnexcusedOtherAmount, isMissedUnexcusedChecked)
            UpdateFinancialLabelWithColor(lgasum, stats.MissedUnexcusedTotalAmount, isMissedUnexcusedChecked)

            ' გაცდენა საპატიო ლეიბლები (ვფიქრობ lgsk, lgsa, lgssum იქნება)
            Dim isMissedExcusedChecked = CheckBoxIsChecked(cbgs)
            UpdateFinancialLabelWithColor(lgsk, stats.MissedExcusedPrivateAmount, isMissedExcusedChecked)
            UpdateFinancialLabelWithColor(lgsa, stats.MissedExcusedOtherAmount, isMissedExcusedChecked)
            UpdateFinancialLabelWithColor(lgssum, stats.MissedExcusedTotalAmount, isMissedExcusedChecked)

            ' პროგრამით გატარება ლეიბლები (ვფიქრობ lpgk, lpga, lpgsum იქნება)
            Dim isAutoProcessedChecked = CheckBoxIsChecked(cbpg)
            UpdateFinancialLabelWithColor(lpgk, stats.AutoProcessedPrivateAmount, isAutoProcessedChecked)
            UpdateFinancialLabelWithColor(lpga, stats.AutoProcessedOtherAmount, isAutoProcessedChecked)
            UpdateFinancialLabelWithColor(lpgsum, stats.AutoProcessedTotalAmount, isAutoProcessedChecked)

            ' გაუქმებული ლეიბლები (ვფიქრობ lgk, lga, lgsum იქნება)
            Dim isCancelledChecked = CheckBoxIsChecked(cbg)
            UpdateFinancialLabelWithColor(lgk, stats.CancelledPrivateAmount, isCancelledChecked)
            UpdateFinancialLabelWithColor(lga, stats.CancelledOtherAmount, isCancelledChecked)
            UpdateFinancialLabelWithColor(lgsum, stats.CancelledTotalAmount, isCancelledChecked)

            ' დაგეგმილი მომავალში ლეიბლები (ვფიქრობ ldk, lda, ldsum იქნება)
            Dim isPlannedInTimeChecked = CheckBoxIsChecked(cbd)
            UpdateFinancialLabelWithColor(ldk, stats.PlannedInTimePrivateAmount, isPlannedInTimeChecked)
            UpdateFinancialLabelWithColor(lda, stats.PlannedInTimeOtherAmount, isPlannedInTimeChecked)
            UpdateFinancialLabelWithColor(ldsum, stats.PlannedInTimeTotalAmount, isPlannedInTimeChecked)

            ' დაგეგმილი ვადაგადაცილებული ლეიბლები (ვფიქრობ ldvk, ldva, ldvsum იქნება)
            Dim isPlannedOverdueChecked = CheckBoxIsChecked(cbdv)
            UpdateFinancialLabelWithColor(ldvk, stats.PlannedOverduePrivateAmount, isPlannedOverdueChecked)
            UpdateFinancialLabelWithColor(ldva, stats.PlannedOverdueOtherAmount, isPlannedOverdueChecked)
            UpdateFinancialLabelWithColor(ldvsum, stats.PlannedOverdueTotalAmount, isPlannedOverdueChecked)

            ' საერთო ჯამების ლეიბლები - ყოველთვის შავი ფერით
            UpdateFinancialLabelWithColor(lsumk, stats.TotalPrivateAmount, True)
            UpdateFinancialLabelWithColor(lsuma, stats.TotalOtherAmount, True)
            UpdateFinancialLabelWithColor(lsumsum, stats.GrandTotalAmount, True)

            Debug.WriteLine("UpdateFinancialLabels: ყველა ფინანსური ლეიბლი განახლებულია")

        Catch ex As Exception
            Debug.WriteLine($"UpdateFinancialLabels: შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფინანსური ლეიბლის განახლება ფერისა და ფორმატის კონტროლით
    ''' </summary>
    Private Sub UpdateFinancialLabelWithColor(label As Label, amount As Decimal, isEnabled As Boolean)
        Try
            If label IsNot Nothing Then
                ' ფორმატირება: 0.00₾
                label.Text = $"{amount:F2}₾"

                ' ფერის დაყენება: შავი თუ მონიშნულია, ნაცრისფერი თუ არა
                label.ForeColor = If(isEnabled, Color.Black, Color.Gray)
            End If
        Catch ex As Exception
            Debug.WriteLine($"UpdateFinancialLabelWithColor: შეცდომა: {ex.Message}")
        End Try
    End Sub
#End Region
End Class