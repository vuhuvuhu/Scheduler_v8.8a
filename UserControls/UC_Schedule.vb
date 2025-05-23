' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განახლებული განრიგის UserControl
' ახალი სერვისებით: ScheduleStatisticsDisplayService და ScheduleFinancialAnalysisService
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

    ' 🆕 ახალი სერვისები სტატისტიკისა და ფინანსური ანალიზისთვის
    Private statisticsDisplayService As ScheduleStatisticsDisplayService = Nothing
    Private financialAnalysisService As ScheduleFinancialAnalysisService = Nothing

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRoleID As Integer = 6 ' ნაგულისხმევი როლი

    ' გვერდების მონაცემები
    Private currentPage As Integer = 1

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
    ''' მონაცემთა სერვისისა და მომხმარებლის ინფორმაციის დაყენება - შესწორებული ვერსია
    ''' ✅ ამ ვერსიაში პირველ რიგში dataService-იდან მოაქვს მომხმარებლის როლი
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი</param>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი (Optional - თუ არ მითითებულა, dataService-დან მოვიღებთ)</param>
    Public Sub SetDataService(service As IDataService, Optional email As String = "", Optional role As Integer = -1)
        Try
            Debug.WriteLine($"UC_Schedule: SetDataService - email='{email}', role={role}")

            dataService = service
            userEmail = email

            ' 🔧 მთავარი შესწორება: ჯერ dataService-დან ვიღებთ როლს
            If dataService IsNot Nothing AndAlso Not String.IsNullOrEmpty(email) Then
                Try
                    Dim userRoleString As String = dataService.GetUserRole(email)
                    Debug.WriteLine($"UC_Schedule: dataService.GetUserRole('{email}') = '{userRoleString}'")

                    If Not String.IsNullOrEmpty(userRoleString) Then
                        Dim parsedRole As Integer
                        If Integer.TryParse(userRoleString, parsedRole) Then
                            userRoleID = parsedRole
                            Debug.WriteLine($"UC_Schedule: ✅ როლი dataService-დან: {userRoleID}")
                        Else
                            Debug.WriteLine($"UC_Schedule: ❌ როლის პარსინგი ვერ მოხერხდა: '{userRoleString}'")
                            userRoleID = If(role = -1, 6, role) ' ნაგულისხმევი 6, თუ role არ მითითებულა
                        End If
                    Else
                        Debug.WriteLine($"UC_Schedule: ❌ ცარიელი როლი dataService-დან")
                        userRoleID = If(role = -1, 6, role)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"UC_Schedule: ❌ dataService-დან როლის მიღების შეცდომა: {ex.Message}")
                    userRoleID = If(role = -1, 6, role)
                End Try
            Else
                ' თუ dataService არ არის ან email ცარიელია
                userRoleID = If(role = -1, 6, role)
                Debug.WriteLine($"UC_Schedule: dataService ან email არ არის - ვიყენებ role={userRoleID}")
            End If

            Debug.WriteLine($"UC_Schedule: ✅ საბოლოო userRoleID: {userRoleID}")

            ' Label22-ში როლის ჩვენება დიაგნოსტიკისთვის
            If Label22 IsNot Nothing Then
                Label22.Text = $"მომხმარებლის როლი: {userRoleID}"
                Label22.ForeColor = If(userRoleID = 1, Color.Green, If(userRoleID = 2, Color.Blue, Color.Red))
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
            If Label22 IsNot Nothing Then
                Label22.Text = $"SetDataService შეცდომა: {ex.Message}"
                Label22.ForeColor = Color.Red
            End If
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

            ' 🆕 ფინანსური პანელის ხილულობის განახლება
            financialAnalysisService?.SetVisibilityByUserRole(userRoleID)

            ' UI-ს ხელახალი კონფიგურაცია ახალი როლისთვის
            If uiManager IsNot Nothing Then
                LoadFilteredSchedule() ' ეს ახლიდან შექმნის სვეტებს
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ForceSetUserRole შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ყველა ფინანსური CheckBox-ის მონიშვნა/განიშვნა
    ''' </summary>
    ''' <param name="checkAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
    Public Sub SetAllFinancialCheckBoxes(checkAll As Boolean)
        Try
            Debug.WriteLine($"UC_Schedule: ყველა ფინანსური CheckBox-ის დაყენება: {checkAll}")
            financialAnalysisService?.SetAllCheckBoxes(checkAll)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: SetAllFinancialCheckBoxes შეცდომა: {ex.Message}")
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
            'სატესტოდ როლისთვის
            If Label22 IsNot Nothing Then
                Label22.Text = $"მომხმარებლის როლი: {userRoleID}"
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

            ' 🆕 სტატისტიკის ჩვენების სერვისი (GBSumInf გრუპბოქსისთვის)
            If GBSumInf IsNot Nothing Then
                statisticsDisplayService = New ScheduleStatisticsDisplayService(dataService, GBSumInf)
                Debug.WriteLine("UC_Schedule: ScheduleStatisticsDisplayService ინიციალიზებული")
            Else
                Debug.WriteLine("UC_Schedule: GBSumInf გრუპბოქსი არ მოიძებნა")
            End If

            ' 🆕 ფინანსური ანალიზის სერვისი (GBSumFin გრუპბოქსისთვის)
            If GBSumFin IsNot Nothing Then
                financialAnalysisService = New ScheduleFinancialAnalysisService(GBSumFin)
                ' მომხმარებლის როლის მიხედვით ხილულობის დაყენება
                financialAnalysisService.SetVisibilityByUserRole(userRoleID)
                Debug.WriteLine("UC_Schedule: ScheduleFinancialAnalysisService ინიციალიზებული")
            Else
                Debug.WriteLine("UC_Schedule: GBSumFin გრუპბოქსი არ მოიძებნა")
            End If

            ' სტატუსის CheckBox-ების ძებნა და დამატება
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია
            filterManager.InitializeFilters()

            ' ივენთების მიბმა
            BindEvents()

            ' ღილაკების მრგვალი ფორმის დაყენება
            ConfigureRoundButtons()

            ' ComboBox-ების შევსება ფონზე (პერფორმანსისთვის) - განახლებული ლოგიკა
            System.Threading.Tasks.Task.Run(Sub() PopulateComboBoxesAsync())

            Debug.WriteLine("UC_Schedule: სერვისების ინიციალიზაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: InitializeServices შეცდომა: {ex.Message}")
            Throw
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

            ' 🆕 სტატისტიკისა და ფინანსური ანალიზის განახლება (ასინქრონულად)
            System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAndFinancialsAsync(criteria))

            Debug.WriteLine($"UC_Schedule: ჩატვირთულია {result.Data.Count} ჩანაწერი, გვერდი {result.CurrentPage}/{result.TotalPages}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 სტატისტიკისა და ფინანსური ანალიზის ასინქრონული განახლება
    ''' </summary>
    Private Sub UpdateStatisticsAndFinancialsAsync(criteria As ScheduleDataProcessor.FilterCriteria)
        Try
            Debug.WriteLine("UC_Schedule: სტატისტიკისა და ფინანსების ასინქრონული განახლება")

            ' ყველა გაფილტრული მონაცემის მიღება (არა მხოლოდ მიმდინარე გვერდი)
            Dim allFilteredResult = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)
            Dim allFilteredData = allFilteredResult.Data

            Debug.WriteLine($"UC_Schedule: მიღებულია {allFilteredData.Count} გაფილტრული ჩანაწერი სტატისტიკისთვის")

            ' UI Thread-ზე სტატისტიკისა და ფინანსების განახლება
            Me.Invoke(Sub()
                          Try
                              ' 🆕 GBSumInf სტატისტიკის განახლება
                              If statisticsDisplayService IsNot Nothing Then
                                  statisticsDisplayService.UpdateStatisticsAsync(allFilteredData)
                                  Debug.WriteLine("UC_Schedule: სტატისტიკის ჩვენება განახლდა")
                              End If

                              ' 🆕 GBSumFin ფინანსური ანალიზის განახლება (მხოლოდ ადმინისა და მენეჯერისთვის)
                              If financialAnalysisService IsNot Nothing AndAlso (userRoleID = 1 OrElse userRoleID = 2) Then
                                  financialAnalysisService.UpdateFinancialData(allFilteredData)
                                  Debug.WriteLine("UC_Schedule: ფინანსური ანალიზი განახლდა")
                              End If

                          Catch ex As Exception
                              Debug.WriteLine($"UC_Schedule: სტატისტიკის UI განახლების შეცდომა: {ex.Message}")
                          End Try
                      End Sub)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateStatisticsAndFinancialsAsync შეცდომა: {ex.Message}")
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
    ''' წინა გვერდის ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub OnPreviousPageClick(sender As Object, e As EventArgs)
        Try
            If currentPage > 1 Then
                currentPage -= 1
                Debug.WriteLine($"UC_Schedule: წინა გვერდი - {currentPage}")
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPreviousPageClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub OnNextPageClick(sender As Object, e As EventArgs)
        Try
            If BtnNext.Enabled Then
                currentPage += 1
                Debug.WriteLine($"UC_Schedule: შემდეგი გვერდი - {currentPage}")
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnNextPageClick შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "🆕 ღილაკების ივენთ ჰენდლერები - სტატისტიკისა და ფინანსებისთვის"

    ''' <summary>
    ''' 🆕 განახლების ღილაკზე დაჭერის ივენთი (BtnRef)
    ''' </summary>
    Private Sub BtnRef_Click(sender As Object, e As EventArgs) Handles BtnRef.Click
        Try
            Debug.WriteLine("UC_Schedule: განახლების ღილაკზე დაჭერა")
            RefreshData()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtnRef_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ახალი ჩანაწერის დამატების ღილაკზე დაჭერის ივენთი (BtnAddSchedule)
    ''' </summary>
    Private Sub BtnAddSchedule_Click(sender As Object, e As EventArgs) Handles BtnAddSchedule.Click
        Try
            Debug.WriteLine("UC_Schedule: ახალი ჩანაწერის დამატების ღილაკზე დაჭერა")

            ' NewRecordForm-ის გახსნა ახალი ჩანაწერის შექმნისთვის
            Dim currentUserEmail As String = If(String.IsNullOrEmpty(userEmail), "user@example.com", userEmail)

            Using addForm As New NewRecordForm(dataService, "სესია", 0, currentUserEmail, "UC_Schedule")
                Dim result As DialogResult = addForm.ShowDialog()

                If result = DialogResult.OK Then
                    RefreshData()
                    MessageBox.Show("ახალი სესია წარმატებით დაემატა", "წარმატება",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtnAddSchedule_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის შექმნის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ბეჭდვის ღილაკზე დაჭერის ივენთი (btbPrint)
    ''' </summary>
    Private Sub BtbPrint_Click(sender As Object, e As EventArgs) Handles btbPrint.Click
        Try
            Debug.WriteLine("UC_Schedule: ბეჭდვის ღილაკზე დაჭერა")

            ' მიმდინარე გაფილტრული მონაცემების ბეჭდვა
            If dataProcessor IsNot Nothing AndAlso filterManager IsNot Nothing Then
                Dim criteria = filterManager.GetFilterCriteria()
                Dim allData = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)

                ' ბეჭდვის ფუნქციის გამოძახება
                PrintScheduleData(allData.Data)
            Else
                MessageBox.Show("მონაცემები მზად არ არის ბეჭდვისთვის", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtbPrint_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 PDF-ში ექსპორტის ღილაკზე დაჭერის ივენთი (btnToPDF)
    ''' </summary>
    Private Sub BtnToPDF_Click(sender As Object, e As EventArgs) Handles btnToPDF.Click
        Try
            Debug.WriteLine("UC_Schedule: PDF ექსპორტის ღილაკზე დაჭერა")

            ' მიმდინარე გაფილტრული მონაცემების PDF-ში ექსპორტი
            If dataProcessor IsNot Nothing AndAlso filterManager IsNot Nothing Then
                Dim criteria = filterManager.GetFilterCriteria()
                Dim allData = dataProcessor.GetFilteredSchedule(criteria, 1, Integer.MaxValue)

                ' PDF ექსპორტის ფუნქციის გამოძახება
                ExportScheduleToPDF(allData.Data)
            Else
                MessageBox.Show("მონაცემები მზად არ არის PDF ექსპორტისთვის", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtnToPDF_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "🆕 დამხმარე მეთოდები ბეჭდვისა და ექსპორტისთვის"

    ''' <summary>
    ''' 🆕 მონაცემების ბეჭდვა
    ''' </summary>
    ''' <param name="data">ბეჭდვისთვის მონაცემები</param>
    Private Sub PrintScheduleData(data As List(Of IList(Of Object)))
        Try
            Debug.WriteLine($"UC_Schedule: ბეჭდვა დაიწყო {data.Count} ჩანაწერისთვის")

            ' PrintDocument-ის შექმნა
            Dim printDoc As New Printing.PrintDocument()
            Dim currentRowIndex As Integer = 0
            Dim rowsPerPage As Integer = 25 ' ჩვეულებრივ 25 მწკრივი გვერდზე

            AddHandler printDoc.PrintPage, Sub(sender, e)
                                               Try
                                                   ' ფონტების განსაზღვრა
                                                   Dim headerFont As New Font("Arial", 12, FontStyle.Bold)
                                                   Dim dataFont As New Font("Arial", 8)
                                                   Dim brush As New SolidBrush(Color.Black)

                                                   ' სათაურის ბეჭდვა
                                                   Dim yPos As Single = 50
                                                   e.Graphics.DrawString("განრიგის მონაცემები", headerFont, brush, 50, yPos)
                                                   yPos += 30

                                                   e.Graphics.DrawString($"თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}", dataFont, brush, 50, yPos)
                                                   yPos += 30

                                                   ' სვეტების სათაურები
                                                   Dim columnHeaders() As String = {"ID", "თარიღი", "ბენეფიციარი", "თერაპევტი", "სტატუსი"}
                                                   Dim xPos As Single = 50

                                                   For Each header In columnHeaders
                                                       e.Graphics.DrawString(header, headerFont, brush, xPos, yPos)
                                                       xPos += 120
                                                   Next
                                                   yPos += 25

                                                   ' მონაცემების ბეჭდვა
                                                   Dim rowsPrinted As Integer = 0
                                                   While currentRowIndex < data.Count AndAlso rowsPrinted < rowsPerPage
                                                       Dim row = data(currentRowIndex)
                                                       If row.Count >= 5 Then
                                                           xPos = 50
                                                           ' ID
                                                           e.Graphics.DrawString(row(0).ToString(), dataFont, brush, xPos, yPos)
                                                           xPos += 120
                                                           ' თარიღი
                                                           e.Graphics.DrawString(row(5).ToString(), dataFont, brush, xPos, yPos)
                                                           xPos += 120
                                                           ' ბენეფიციარი
                                                           Dim beneficiary = $"{row(3)} {row(4)}"
                                                           e.Graphics.DrawString(beneficiary, dataFont, brush, xPos, yPos)
                                                           xPos += 120
                                                           ' თერაპევტი
                                                           e.Graphics.DrawString(If(row.Count > 8, row(8).ToString(), ""), dataFont, brush, xPos, yPos)
                                                           xPos += 120
                                                           ' სტატუსი
                                                           e.Graphics.DrawString(If(row.Count > 12, row(12).ToString(), ""), dataFont, brush, xPos, yPos)
                                                       End If

                                                       currentRowIndex += 1
                                                       rowsPrinted += 1
                                                       yPos += 20
                                                   End While

                                                   ' შემდეგი გვერდი საჭიროა თუ არა
                                                   e.HasMorePages = (currentRowIndex < data.Count)

                                               Catch ex As Exception
                                                   Debug.WriteLine($"UC_Schedule: PrintPage შეცდომა: {ex.Message}")
                                               End Try
                                           End Sub

            ' ბეჭდვის დიალოგის ჩვენება
            Dim printDialog As New PrintDialog()
            printDialog.Document = printDoc

            If printDialog.ShowDialog() = DialogResult.OK Then
                printDoc.Print()
                Debug.WriteLine("UC_Schedule: ბეჭდვა წარმატებით დასრულდა")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: PrintScheduleData შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 მონაცემების PDF-ში ექსპორტი
    ''' </summary>
    ''' <param name="data">ექსპორტისთვის მონაცემები</param>
    Private Sub ExportScheduleToPDF(data As List(Of IList(Of Object)))
        Try
            Debug.WriteLine($"UC_Schedule: PDF ექსპორტი დაიწყო {data.Count} ჩანაწერისთვის")

            ' SaveFileDialog-ის ჩვენება
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "PDF ფაილები (*.pdf)|*.pdf"
                saveDialog.Title = "განრიგის PDF ექსპორტი"
                saveDialog.FileName = $"Ganrigi_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    ' თუ iTextSharp-ი ხელმისაწვდომია, გამოვიყენოთ ის
                    ' თუ არა, მარტივი ტექსტური ფაილის შექმნა
                    ExportToTextFile(data, saveDialog.FileName.Replace(".pdf", ".txt"))

                    MessageBox.Show($"მონაცემები წარმატებით ექსპორტირდა: {saveDialog.FileName.Replace(".pdf", ".txt")}",
                                  "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportScheduleToPDF შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 მონაცემების ტექსტურ ფაილში ექსპორტი
    ''' </summary>
    ''' <param name="data">ექსპორტისთვის მონაცემები</param>
    ''' <param name="filePath">ფაილის მისამართი</param>
    Private Sub ExportToTextFile(data As List(Of IList(Of Object)), filePath As String)
        Try
            Using writer As New IO.StreamWriter(filePath, False, System.Text.Encoding.UTF8)
                ' სათაური
                writer.WriteLine("განრიგის მონაცემები")
                writer.WriteLine($"თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}")
                writer.WriteLine(New String("="c, 80))
                writer.WriteLine()

                ' სვეტების სათაურები
                writer.WriteLine("ID".PadRight(10) &
                               "თარიღი".PadRight(20) &
                               "ბენეფიციარი".PadRight(25) &
                               "თერაპევტი".PadRight(20) &
                               "სტატუსი".PadRight(15))
                writer.WriteLine(New String("-"c, 80))

                ' მონაცემები
                For Each row In data
                    If row.Count >= 5 Then
                        Dim line = row(0).ToString().PadRight(10) &
                                 If(row.Count > 5, row(5).ToString(), "").PadRight(20) &
                                 $"{row(3)} {row(4)}".PadRight(25) &
                                 If(row.Count > 8, row(8).ToString(), "").PadRight(20) &
                                 If(row.Count > 12, row(12).ToString(), "").PadRight(15)
                        writer.WriteLine(line)
                    End If
                Next

                writer.WriteLine()
                writer.WriteLine(New String("="c, 80))
                writer.WriteLine($"სულ ჩანაწერები: {data.Count}")
            End Using

            Debug.WriteLine($"UC_Schedule: ტექსტური ფაილი შეიქმნა: {filePath}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportToTextFile შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

#End Region

End Class