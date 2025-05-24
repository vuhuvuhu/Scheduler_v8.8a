' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განახლებული განრიგის UserControl - გამარტივებული და ოპტიმიზირებული
' მომხმარებლის ელფოსტისა და როლის სწორი მიღება Form1-დან
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Services
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

            ' ლეიბლების განახლება
            UpdateUserInfoLabels()

            ' ფინანსური პანელის ხილულობის განახლება
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
    ''' კონკრეტულ გვერდზე გადასვლა
    ''' </summary>
    Public Sub GoToPage(pageNumber As Integer)
        Try
            If pageNumber >= 1 Then
                currentPage = pageNumber
                LoadFilteredSchedule()
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

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვსრულებთ ინიციალიზაციას
            If dataService IsNot Nothing Then
                InitializeServices()
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის ლეიბლების განახლება
    ''' </summary>
    Private Sub UpdateUserInfoLabels()
        Try
            ' Label22 - მომხმარებლის როლი
            If Label22 IsNot Nothing Then
                Label22.Text = $"როლი: {userRoleID}"
                Label22.ForeColor = If(userRoleID = 1, Color.Green, If(userRoleID = 2, Color.Blue, Color.Red))
            End If

            ' Label23 - მომხმარებლის ელფოსტა
            If Label23 IsNot Nothing Then
                Label23.Text = $"ელფოსტა: {userEmail}"
                Label23.ForeColor = Color.Black
            End If

            Debug.WriteLine($"UC_Schedule: ლეიბლები განახლდა - როლი: {userRoleID}, ელფოსტა: '{userEmail}'")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: UpdateUserInfoLabels შეცდომა: {ex.Message}")
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

            ' სტატუსის CheckBox-ების ინიციალიზაცია
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
    ''' სტატუსის CheckBox-ების ინიციალიზაცია
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
    ''' ივენთების მიბმა
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

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BindEvents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დაყენება
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
    ''' განრიგის მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadFilteredSchedule()
        Try
            If dataProcessor Is Nothing OrElse uiManager Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_Schedule: სერვისები არ არის ინიციალიზებული")
                Return
            End If

            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule - გვერდი {currentPage}")

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

            ' სტატისტიკისა და ფინანსური ანალიზის განახლება
            System.Threading.Tasks.Task.Run(Sub() UpdateStatisticsAsync(criteria))

            Debug.WriteLine($"UC_Schedule: ჩატვირთულია {result.Data.Count} ჩანაწერი")

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

                              ' ფინანსური ანალიზის განახლება (მხოლოდ ადმინისა და მენეჯერისთვის)
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
    ''' ფილტრის შეცვლის ივენთი
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
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
            ' რედაქტირების ღილაკზე დაჭერის შემოწმება
            If uiManager.IsEditButtonClicked(e) Then
                Dim sessionIdValue = uiManager.GetCellValue("N", e.RowIndex)

                If sessionIdValue IsNot Nothing AndAlso IsNumeric(sessionIdValue) Then
                    Dim sessionId As Integer = CInt(sessionIdValue)
                    Debug.WriteLine($"UC_Schedule: რედაქტირება - სესია ID={sessionId}")

                    ' 🔧 მომხმარებლის ელფოსტის სწორი გადაცემა
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
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnDataGridViewCellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკი
    ''' </summary>
    Private Sub OnPreviousPageClick(sender As Object, e As EventArgs)
        Try
            If currentPage > 1 Then
                currentPage -= 1
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnPreviousPageClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკი
    ''' </summary>
    Private Sub OnNextPageClick(sender As Object, e As EventArgs)
        Try
            If BtnNext.Enabled Then
                currentPage += 1
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnNextPageClick შეცდომა: {ex.Message}")
        End Try
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
        Try
            Debug.WriteLine("UC_Schedule: ახალი სესიის დამატება")

            ' 🔧 მომხმარებლის ელფოსტის სწორი გადაცემა
            Using addForm As New NewRecordForm(dataService, "სესია", 0, userEmail, "UC_Schedule")
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
    ''' 🖨️ ბეჭდვის ღილაკი - DgvSchedule-ის მონაცემების ბეჭდვა
    ''' გაუმჯობესებული ვერსია: სვეტების მონიშვნა + პრინტერის არჩევა + ლანდშაფტი
    ''' </summary>
    Private Sub BtbPrint_Click(sender As Object, e As EventArgs) Handles btbPrint.Click
        Try
            Debug.WriteLine("UC_Schedule: გაუმჯობესებული ბეჭდვის ღილაკზე დაჭერა")

            ' შევამოწმოთ არის თუ არა მონაცემები ბეჭდვისთვის
            If DgvSchedule Is Nothing OrElse DgvSchedule.Rows.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' მომხმარებელს ავარჩევინოთ ბეჭდვის ტიპი
            Dim printTypeResult As DialogResult = MessageBox.Show(
                "რომელი ტიპის ბეჭდვა გსურთ?" & Environment.NewLine & Environment.NewLine &
                "დიახ - გაუმჯობესებული ბეჭდვა (სვეტების მონიშვნა + ლანდშაფტი)" & Environment.NewLine &
                "არა - ჩვეულებრივი ბეჭდვა" & Environment.NewLine &
                "გაუქმება - ოპერაციის შეწყვეტა",
                "ბეჭდვის ტიპის არჩევა",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case printTypeResult
                Case DialogResult.Yes
                    ' 🔧 გაუმჯობესებული ბეჭდვა - სვეტების მონიშვნით და ლანდშაფტით
                    Using advancedPrintService As New AdvancedDataGridViewPrintService(DgvSchedule)
                        advancedPrintService.ShowFullPrintDialog()
                    End Using

                Case DialogResult.No
                    ' ჩვეულებრივი ბეჭდვა (ძველი ვერსია)
                    Using simplePrintService As New DataGridViewPrintService(DgvSchedule)
                        Dim result As DialogResult = MessageBox.Show(
                            "გსურთ ჯერ ნახოთ ბეჭდვის პრევიუ?",
                            "ბეჭდვის პრევიუ",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question)

                        Select Case result
                            Case DialogResult.Yes
                                simplePrintService.ShowPrintPreview()
                            Case DialogResult.No
                                simplePrintService.Print()
                            Case DialogResult.Cancel
                                Debug.WriteLine("UC_Schedule: ჩვეულებრივი ბეჭდვა გაუქმებულია")
                        End Select
                    End Using

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_Schedule: ბეჭდვა გაუქმებულია მომხმარებლის მიერ")

            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: BtbPrint_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class