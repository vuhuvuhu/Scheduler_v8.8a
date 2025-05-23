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

            ' ✅ ღილაკების მრგვალი ფორმის დაყენება
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

    ''' <summary>
    ''' შემაჯამებელი პანელების განახლება
    ''' </summary>
    Private Sub UpdateSummaryPanels(totalSessions As Integer, completedSessions As Integer, totalRevenue As Decimal)
        Try
            ' შეიძლება აქ დაემატოს სტატისტიკის Label-ების განახლება
            ' მაგალითად: lblTotalSessions.Text = totalSessions.ToString()
            Debug.WriteLine($"UC_Schedule: სტატისტიკა - სულ: {totalSessions}, შესრულებული: {completedSessions}, შემოსავალი: {totalRevenue:F2}")
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateSummaryPanels შეცდომა: {ex.Message}")
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

End Class