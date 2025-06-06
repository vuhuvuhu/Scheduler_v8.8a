﻿' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განახლებული განრიგის UserControl - გამარტივებული და ოპტიმიზირებული
' მომხმარებლის ელფოსტისა და როლის სწორი მიღება Form1-დან
' ===========================================
Imports System.Windows.Forms
'Imports Scheduler_v8_8a.Services
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
    ''' განრიგის მონაცემების ჩატვირთვა - შესწორებული ვერსია
    ''' currentPage-ის ორმაგი განახლების თავიდან აცილებით
    ''' </summary>
    Private Sub LoadFilteredSchedule()
        Try
            ' 🔧 შემოწმებები ციკლური გადაკვრის თავიდან ასაცილებლად
            If dataProcessor Is Nothing OrElse uiManager Is Nothing OrElse filterManager Is Nothing Then
                Debug.WriteLine("UC_Schedule: სერვისები არ არის ინიციალიზებული")
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
    ''' 🔧 გამართული ფილტრის შეცვლის ივენთი
    ''' ნავიგაციის დროს გვერდი არ რესეტდება
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine($"UC_Schedule: OnFilterChanged - isNavigating: {isNavigating}, isLoadingData: {isLoadingData}")

            ' 🔧 ნავიგაციის დროს ფილტრის ივენთი არ უნდა გავარდეს
            If isNavigating Then
                Debug.WriteLine("UC_Schedule: ნავიგაციის დროს FilterChanged ივენთი იგნორირებულია")
                Return
            End If

            ' 🔧 მონაცემების დატვირთვის დროს არ ვუშვებთ ციკლს
            If isLoadingData Then
                Debug.WriteLine("UC_Schedule: მონაცემების დატვირთვის დროს FilterChanged ივენთი იგნორირებულია")
                Return
            End If

            ' ფილტრის შეცვლისას გვერდი რესეტდება 1-ზე
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: OnFilterChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვერდის ზომის შეცვლის ივენთი - შესწორებული ვერსია
    ''' </summary>
    Private Sub OnPageSizeChanged()
        Try
            Debug.WriteLine("UC_Schedule: OnPageSizeChanged - გვერდის ზომა შეიცვალა, პირველ გვერდზე გადასვლა")
            currentPage = 1 ' 🔧 გვერდის ზომის შეცვლისას ყოველთვის პირველ გვერდზე
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

    ' ========================================
    ' 🔧 UC_Schedule.vb-ში შეცვალეთ ეს მეთოდები:
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
                    Debug.WriteLine($"UC_Schedule: წარმატებით გადავედი გვერდზე {currentPage}")
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
                    Debug.WriteLine($"UC_Schedule: წარმატებით გადავედი გვერდზე {currentPage}")
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
            ' შევამოწმოთ უკვე გახსნილია თუ არა NewRecordForm
            For Each frm As Form In Application.OpenForms
                If TypeOf frm Is NewRecordForm Then
                    ' თუ უკვე გახსნილია, მოვიტანოთ წინ და გამოვიდეთ მეთოდიდან
                    Debug.WriteLine("BtnAddSchedule_Click: NewRecordForm უკვე გახსნილია, ფოკუსის გადატანა")
                    frm.Focus()
                    Return
                End If
            Next

            ' შევამოწმოთ გვაქვს თუ არა dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Debug.WriteLine("BtnAddSchedule_Click: dataService არ არის ინიციალიზებული")
                Return
            End If

            ' ნაგულისხმევად "სესია" ტიპი
            Dim recordType As String = "სესია"

            ' NewRecordForm-ის გახსნა Add რეჟიმში
            Dim newRecordForm As New NewRecordForm(dataService, recordType, userEmail, "UC_Calendar")
            Dim result = newRecordForm.ShowDialog()

            ' თუ ფორმა დაიხურა OK რეზულტატით, განვაახლოთ მონაცემები
            If result = DialogResult.OK Then
                Debug.WriteLine("BtnAddSchedule_Click: სესია წარმატებით დაემატა")

                ' ჩავტვირთოთ სესიები თავიდან
                RefreshData()

            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnAddSchedule_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ახალი ჩანაწერის ფორმის გახსნის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    ' ========================================
    ' 🔧 დაამატეთ ეს მეთოდები UC_Schedule.vb ფაილში
    ' არსებული btnToPDF_Click მეთოდის ჩანაცვლებად
    ' ========================================

    ''' <summary>
    ''' 📄 PDF ექსპორტის ღილაკი - სვეტების მონიშვნით ბეჭდვის მსგავსად
    ''' გაუმჯობესებული ვერსია: მუშაობს ზუსტად როგორც btbPrint
    ''' </summary>
    Private Sub btnToPDF_Click(sender As Object, e As EventArgs) Handles btnToPDF.Click
        Try
            Debug.WriteLine("UC_Schedule: PDF ექსპორტის ღილაკზე დაჭერა")

            ' შევამოწმოთ არის თუ არა მონაცემები ექსპორტისთვის
            If DgvSchedule Is Nothing OrElse DgvSchedule.Rows.Count = 0 Then
                MessageBox.Show("PDF ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 🔧 მომხმარებელს ავარჩევინოთ PDF ექსპორტის ტიპი (btbPrint-ის ანალოგიურად)
            Dim pdfTypeResult As DialogResult = MessageBox.Show(
            "რომელი ტიპის PDF ექსპორტი გსურთ?" & Environment.NewLine & Environment.NewLine &
            "დიახ - გაუმჯობესებული PDF ექსპორტი (სვეტების მონიშვნით)" & Environment.NewLine &
            "არა - მარტივი PDF ექსპორტი (ყველა ხილული სვეტით)" & Environment.NewLine &
            "გაუქმება - ოპერაციის შეწყვეტა",
            "PDF ექსპორტის ტიპის არჩევა",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question)

            Select Case pdfTypeResult
                Case DialogResult.Yes
                    ' 🔧 გაუმჯობესებული PDF ექსპორტი - სვეტების მონიშვნით (btbPrint-ის ანალოგიური)
                    ExportAdvancedPDFWithColumnSelection()

                Case DialogResult.No
                    ' მარტივი PDF ექსპორტი - HTML ალტერნატივით (ძველი ვერსია)
                    ExportSimplePDFAlternative()

                Case DialogResult.Cancel
                    Debug.WriteLine("UC_Schedule: PDF ექსპორტი გაუქმებულია მომხმარებლის მიერ")

            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: btnToPDF_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 გაუმჯობესებული PDF ექსპორტი სვეტების მონიშვნით - btbPrint-ის ანალოგია
    ''' იყენებს ColumnSelectionForm-ს ზუსტად როგორც AdvancedDataGridViewPrintService
    ''' </summary>
    Private Sub ExportAdvancedPDFWithColumnSelection()
        Try
            Debug.WriteLine("UC_Schedule: გაუმჯობესებული PDF ექსპორტი სვეტების მონიშვნით")

            ' 1. სვეტების მონიშვნის დიალოგის ჩვენება - btbPrint-ის ანალოგიურად
            Dim selectedColumns As List(Of DataGridViewColumn) = Nothing

            Using columnDialog As New ColumnSelectionForm(DgvSchedule)
                Dim selectionResult As DialogResult = columnDialog.ShowDialog()

                If selectionResult <> DialogResult.OK Then
                    Debug.WriteLine("UC_Schedule: სვეტების მონიშვნა გაუქმებულია")
                    Return
                End If

                selectedColumns = columnDialog.GetSelectedColumns()
                Debug.WriteLine($"UC_Schedule: მონიშნულია {selectedColumns.Count} სვეტი PDF ექსპორტისთვის")
            End Using

            If selectedColumns Is Nothing OrElse selectedColumns.Count = 0 Then
                MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი PDF ექსპორტისთვის", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 2. PDF ფაილის შენახვის ადგილის არჩევა
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "PDF ფაილები (*.pdf)|*.pdf"
                saveDialog.Title = "PDF ფაილის შენახვა"
                saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    ' 3. PDF-ის შექმნა მონიშნული სვეტებით
                    CreateAdvancedPDF(saveDialog.FileName, selectedColumns)
                Else
                    Debug.WriteLine("UC_Schedule: ფაილის შენახვის დიალოგი გაუქმებულია")
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportAdvancedPDFWithColumnSelection შეცდომა: {ex.Message}")
            MessageBox.Show($"გაუმჯობესებული PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 მონიშნული სვეტებით PDF-ის შექმნა - btbPrint ლოგიკის ანალოგია
    ''' </summary>
    ''' <param name="filePath">ფაილის მისამართი</param>
    ''' <param name="selectedColumns">მონიშნული სვეტების სია</param>
    Private Sub CreateAdvancedPDF(filePath As String, selectedColumns As List(Of DataGridViewColumn))
        Try
            Debug.WriteLine($"UC_Schedule: CreateAdvancedPDF - {filePath}, სვეტები: {selectedColumns.Count}")

            ' 🔧 ვიყენებთ HTML ალტერნატივა მონიშნული სვეტებით
            ' რადგანაც PDF ბიბლიოთეკა შეიძლება არ იყოს ხელმისაწვდომი
            CreateAdvancedHTMLForPDFExport(filePath.Replace(".pdf", ".html"), selectedColumns)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreateAdvancedPDF შეცდომა: {ex.Message}")
            ' შეცდომის შემთხვევაში HTML ალტერნატივა
            MessageBox.Show("PDF შექმნის შეცდომა. შეიქმნება HTML ფაილი ბრაუზერში PDF-ად ბეჭდვისთვის.",
                       "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CreateAdvancedHTMLForPDFExport(filePath.Replace(".pdf", ".html"), selectedColumns)
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 HTML ფაილის შექმნა PDF-ისთვის მონიშნული სვეტებით
    ''' იყენებს ზუსტად იმ ლოგიკას რაც btbPrint-ში
    ''' </summary>
    ''' <param name="filePath">ფაილის მისამართი</param>
    ''' <param name="selectedColumns">მონიშნული სვეტების სია</param>
    Private Sub CreateAdvancedHTMLForPDFExport(filePath As String, selectedColumns As List(Of DataGridViewColumn))
        Try
            Debug.WriteLine($"UC_Schedule: CreateAdvancedHTMLForPDFExport - სვეტები: {selectedColumns.Count}")

            Dim html As New System.Text.StringBuilder()

            ' HTML დოკუმენტის დასაწყისი - PDF-ისთვის ოპტიმიზირებული
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html lang=""ka"">")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>განრიგის მონაცემები - PDF ექსპორტი</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        @page { size: A4 landscape; margin: 15mm; }")
            html.AppendLine("        @media print { ")
            html.AppendLine("            body { margin: 0; font-size: 10px; }")
            html.AppendLine("            .no-print { display: none; }")
            html.AppendLine("            table { page-break-inside: avoid; }")
            html.AppendLine("            h1 { font-size: 16px; margin: 5px 0; }")
            html.AppendLine("        }")
            html.AppendLine("        body { font-family: 'Sylfaen', Arial, sans-serif; font-size: 11px; }")
            html.AppendLine("        h1 { text-align: center; font-size: 18px; margin: 10px 0; color: #333; }")
            html.AppendLine("        table { width: 100%; border-collapse: collapse; font-size: 9px; }")
            html.AppendLine("        th, td { padding: 3px 2px; border: 1px solid #333; text-align: left; vertical-align: top; }")
            html.AppendLine("        th { background-color: #ddd; font-weight: bold; text-align: center; }")
            html.AppendLine("        tr:nth-child(even) { background-color: #f9f9f9; }")
            html.AppendLine("        .info { text-align: center; margin: 8px 0; font-size: 10px; color: #666; }")
            html.AppendLine("        .button-container { text-align: center; margin: 20px 0; }")
            html.AppendLine("        .pdf-button { padding: 15px 30px; font-size: 16px; background: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer; }")
            html.AppendLine("        .instructions { margin: 10px 0; color: #666; }")
            html.AppendLine("    </style>")
            html.AppendLine("</head>")
            html.AppendLine("<body>")

            ' სათაური
            html.AppendLine("    <h1>განრიგის მონაცემები - PDF ექსპორტი</h1>")

            ' ინფორმაცია
            html.AppendLine("    <div class=""info"">")
            html.AppendLine($"        თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm} | ჩანაწერები: {DgvSchedule.Rows.Count} | მონიშნული სვეტები: {selectedColumns.Count}")
            html.AppendLine("    </div>")

            ' PDF ღილაკი და ინსტრუქციები
            html.AppendLine("    <div class=""no-print button-container"">")
            html.AppendLine("        <button class=""pdf-button"" onclick=""window.print(); setTimeout(() => window.close(), 1000);"">")
            html.AppendLine("            🖨️ PDF-ად ექსპორტი</button>")
            html.AppendLine("        <div class=""instructions"">")
            html.AppendLine("            <p><strong>ინსტრუქცია:</strong></p>")
            html.AppendLine("            <p>1. დააჭირეთ ღილაკს ⬆️</p>")
            html.AppendLine("            <p>2. აირჩიეთ პრინტერად ""Microsoft Print to PDF""</p>")
            html.AppendLine("            <p>3. დააჭირეთ ""Print"" ღილაკს</p>")
            html.AppendLine("            <p>4. აირჩიეთ PDF ფაილის შენახვის ადგილი</p>")
            html.AppendLine("        </div>")
            html.AppendLine("    </div>")

            ' ცხრილი - მხოლოდ მონიშნული სვეტებით
            html.AppendLine("    <table>")
            html.AppendLine("        <thead><tr>")

            ' სათაურები - მხოლოდ მონიშნული სვეტები
            For Each column In selectedColumns
                html.AppendLine($"            <th>{EscapeHtmlText(column.HeaderText)}</th>")
            Next

            html.AppendLine("        </tr></thead>")
            html.AppendLine("        <tbody>")

            ' მონაცემები - მხოლოდ მონიშნული სვეტები
            For rowIndex As Integer = 0 To DgvSchedule.Rows.Count - 1
                html.AppendLine("            <tr>")
                Dim row As DataGridViewRow = DgvSchedule.Rows(rowIndex)

                For Each column In selectedColumns
                    Dim cellValue As String = ""
                    Try
                        If row.Cells(column.Name).Value IsNot Nothing Then
                            cellValue = row.Cells(column.Name).Value.ToString()
                        End If
                    Catch
                        cellValue = ""
                    End Try

                    html.AppendLine($"                <td>{EscapeHtmlText(cellValue)}</td>")
                Next

                html.AppendLine("            </tr>")
            Next

            html.AppendLine("        </tbody>")
            html.AppendLine("    </table>")

            ' ქვედა ინფორმაცია
            html.AppendLine("    <div class=""info"">")
            html.AppendLine($"        შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm} | Scheduler v8.8a | მონიშნული სვეტები: {selectedColumns.Count}")
            html.AppendLine("    </div>")

            html.AppendLine("</body>")
            html.AppendLine("</html>")

            ' ფაილის ჩაწერა
            System.IO.File.WriteAllText(filePath, html.ToString(), System.Text.Encoding.UTF8)

            Debug.WriteLine("UC_Schedule: PDF ექსპორტისთვის HTML ფაილი შეიქმნა")
            MessageBox.Show($"შეიქმნა PDF ექსპორტისთვის ფაილი:{Environment.NewLine}{filePath}" & Environment.NewLine & Environment.NewLine &
                       "ფაილი იხსნება ბრაუზერში. დააჭირეთ 'PDF-ად ექსპორტი' ღილაკს და" & Environment.NewLine &
                       "აირჩიეთ პრინტერად 'Microsoft Print to PDF'",
                       "PDF ექსპორტი", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' ფაილის გახსნა ბრაუზერში
            System.Diagnostics.Process.Start(filePath)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreateAdvancedHTMLForPDFExport შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 მარტივი PDF ექსპორტი - ყველა ხილული სვეტით (ძველი ვერსია)
    ''' </summary>
    Private Sub ExportSimplePDFAlternative()
        Try
            Debug.WriteLine("UC_Schedule: მარტივი PDF ექსპორტი")

            ' PDF ფაილის შენახვის ადგილის არჩევა
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "PDF ფაილები (*.pdf)|*.pdf|HTML ფაილები (*.html)|*.html"
                saveDialog.Title = "ფაილის შენახვა"
                saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    If saveDialog.FileName.EndsWith(".html", StringComparison.OrdinalIgnoreCase) Then
                        ' HTML ფორმატით შენახვა
                        CreateSimpleHTMLFile(saveDialog.FileName)
                    Else
                        ' PDF ალტერნატივა - HTML ფაილით
                        CreateHTMLForPrinting(saveDialog.FileName.Replace(".pdf", ".html"))
                    End If
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportSimplePDFAlternative შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 HTML ტექსტის escape - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Function EscapeHtmlText(text As String) As String
        Try
            If String.IsNullOrEmpty(text) Then
                Return ""
            End If

            ' ძირითადი HTML სიმბოლოების escape
            text = text.Replace("&", "&amp;")  ' ყველაზე პირველ ადგილას!
            text = text.Replace("<", "&lt;")
            text = text.Replace(">", "&gt;")
            text = text.Replace("""", "&quot;")
            text = text.Replace("'", "&#39;")

            ' დამატებითი სიმბოლოები
            text = text.Replace(vbCrLf, "<br>")
            text = text.Replace(vbLf, "<br>")
            text = text.Replace(vbCr, "<br>")

            Return text
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: EscapeHtmlText შეცდომა: {ex.Message}")
            Return If(text, "")
        End Try
    End Function

    ''' <summary>
    ''' პირდაპირ PDF ფაილის შექმნა System.Drawing-ით
    ''' </summary>
    Private Sub CreatePDFDirect(filePath As String)
        Try
            Debug.WriteLine($"UC_Schedule: PDF ფაილის შექმნა - {filePath}")

            ' PrintDocument-ის შექმნა
            Using printDoc As New System.Drawing.Printing.PrintDocument()
                ' PDF პრინტერის მოძიება
                Dim pdfPrinter As String = FindPDFPrinter()

                If String.IsNullOrEmpty(pdfPrinter) Then
                    ' თუ PDF პრინტერი ვერ მოიძებნა, HTML ალტერნატივა
                    MessageBox.Show("PDF პრინტერი ვერ მოიძებნა. შეიქმნება HTML ფაილი ბრაუზერში ბეჭდვისთვის.",
                                   "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    CreateHTMLForPrinting(filePath.Replace(".pdf", ".html"))
                    Return
                End If

                ' PDF პრინტერის დაყენება
                printDoc.PrinterSettings.PrinterName = pdfPrinter
                printDoc.PrinterSettings.PrintToFile = True
                printDoc.PrinterSettings.PrintFileName = filePath

                ' ლანდშაფტ ორიენტაცია
                printDoc.DefaultPageSettings.Landscape = True

                ' ბეჭდვის ივენთების მიბმა
                AddHandler printDoc.PrintPage, AddressOf PrintToPDF

                ' ბეჭდვის დაწყება
                printDoc.Print()

                Debug.WriteLine("UC_Schedule: PDF ფაილი წარმატებით შეიქმნა")
                MessageBox.Show($"PDF ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}",
                               "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' ფაილის გახსნის შეთავაზება
                Dim openResult As DialogResult = MessageBox.Show(
                    "გსურთ PDF ფაილის გახსნა?",
                    "ფაილის გახსნა",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If openResult = DialogResult.Yes Then
                    System.Diagnostics.Process.Start(filePath)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreatePDFDirect შეცდომა: {ex.Message}")
            ' ალტერნატივად HTML ფაილის შექმნა
            MessageBox.Show("PDF შექმნის შეცდომა. შეიქმნება HTML ფაილი ბრაუზერში ბეჭდვისთვის.",
                           "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CreateHTMLForPrinting(filePath.Replace(".pdf", ".html"))
        End Try
    End Sub

    ''' <summary>
    ''' PDF პრინტერის მოძიება სისტემაში
    ''' </summary>
    Private Function FindPDFPrinter() As String
        Try
            Debug.WriteLine("UC_Schedule: PDF პრინტერის ძებნა")

            ' ყველაზე გავრცელებული PDF პრინტერების სახელები
            Dim pdfPrinters As String() = {
                "Microsoft Print to PDF",
                "Microsoft Office Document Image Writer",
                "doPDF",
                "PDFCreator",
                "CutePDF Writer",
                "Foxit Reader PDF Printer",
                "Adobe PDF"
            }

            ' დაყენებული პრინტერების შემოწმება
            For Each printerName In System.Drawing.Printing.PrinterSettings.InstalledPrinters
                For Each pdfName In pdfPrinters
                    If printerName.ToLower().Contains(pdfName.ToLower()) Then
                        Debug.WriteLine($"UC_Schedule: ნაპოვნია PDF პრინტერი - {printerName}")
                        Return printerName
                    End If
                Next
            Next

            Debug.WriteLine("UC_Schedule: PDF პრინტერი ვერ მოიძებნა")
            Return ""

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: FindPDFPrinter შეცდომა: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' PDF-ზე ბეჭდვის ივენთი
    ''' </summary>
    Private Sub PrintToPDF(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs)
        Try
            Debug.WriteLine("UC_Schedule: PDF-ზე ბეჭდვა")

            Dim graphics As Graphics = e.Graphics
            Dim font As New Font("Arial", 10, FontStyle.Regular)
            Dim headerFont As New Font("Arial", 12, FontStyle.Bold)
            Dim titleFont As New Font("Arial", 16, FontStyle.Bold)

            Dim brush As New SolidBrush(Color.Black)
            Dim headerBrush As New SolidBrush(Color.DarkBlue)

            Dim yPosition As Single = 50
            Dim leftMargin As Single = 50
            Dim topMargin As Single = 50

            ' სათაური
            Dim title As String = "განრიგის მონაცემები"
            Dim titleSize As SizeF = graphics.MeasureString(title, titleFont)
            graphics.DrawString(title, titleFont, brush,
                              (e.PageBounds.Width - titleSize.Width) / 2, yPosition)
            yPosition += titleSize.Height + 20

            ' თარიღი და ინფორმაცია
            Dim dateInfo As String = $"თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm} | ჩანაწერები: {DgvSchedule.Rows.Count}"
            graphics.DrawString(dateInfo, font, brush, leftMargin, yPosition)
            yPosition += 30

            ' ხილული სვეტების მიღება (Edit-ის გარდა)
            Dim visibleColumns As New List(Of DataGridViewColumn)
            For Each column As DataGridViewColumn In DgvSchedule.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    visibleColumns.Add(column)
                End If
            Next

            If visibleColumns.Count = 0 Then
                graphics.DrawString("მონაცემები ვერ მოიძებნა", font, brush, leftMargin, yPosition)
                Return
            End If

            ' სვეტების სიგანეების გამოთვლა
            Dim availableWidth As Single = e.PageBounds.Width - (leftMargin * 2)
            Dim columnWidth As Single = availableWidth / visibleColumns.Count
            Dim rowHeight As Single = 25

            ' სათაურები
            Dim xPosition As Single = leftMargin
            For Each column In visibleColumns
                Dim headerRect As New RectangleF(xPosition, yPosition, columnWidth, rowHeight)
                graphics.FillRectangle(Brushes.LightGray, headerRect)
                graphics.DrawRectangle(Pens.Black, xPosition, yPosition, columnWidth, rowHeight)

                ' სათაურის ტექსტი
                Dim headerText As String = TruncateText(column.HeaderText, columnWidth, headerFont, graphics)
                graphics.DrawString(headerText, headerFont, headerBrush,
                                  xPosition + 2, yPosition + 2)
                xPosition += columnWidth
            Next
            yPosition += rowHeight

            ' მონაცემები (მაქსიმუმ 30 მწკრივი პირველი გვერდისთვის)
            Dim maxRows As Integer = Math.Min(DgvSchedule.Rows.Count, 30)

            For rowIndex As Integer = 0 To maxRows - 1
                If yPosition > e.PageBounds.Height - 100 Then
                    ' გვერდი სავსეა
                    Exit For
                End If

                Dim row As DataGridViewRow = DgvSchedule.Rows(rowIndex)
                xPosition = leftMargin

                ' ალტერნაციული ფონი
                If rowIndex Mod 2 = 1 Then
                    graphics.FillRectangle(Brushes.WhiteSmoke,
                                         leftMargin, yPosition, availableWidth, rowHeight)
                End If

                For Each column In visibleColumns
                    ' უჯრის ჩარჩო
                    graphics.DrawRectangle(Pens.LightGray, xPosition, yPosition, columnWidth, rowHeight)

                    ' უჯრის მნიშვნელობა
                    Dim cellValue As String = ""
                    Try
                        If row.Cells(column.Name).Value IsNot Nothing Then
                            cellValue = row.Cells(column.Name).Value.ToString()
                        End If
                    Catch
                        cellValue = ""
                    End Try

                    ' ტექსტის გამოტანა
                    If Not String.IsNullOrEmpty(cellValue) Then
                        Dim truncatedText As String = TruncateText(cellValue, columnWidth - 4, font, graphics)
                        graphics.DrawString(truncatedText, font, brush, xPosition + 2, yPosition + 2)
                    End If

                    xPosition += columnWidth
                Next

                yPosition += rowHeight
            Next

            ' ქვედა ინფორმაცია
            Dim footerY As Single = e.PageBounds.Height - 50
            Dim footerText As String = $"შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm} | Scheduler v8.8a"
            Dim footerSize As SizeF = graphics.MeasureString(footerText, font)
            graphics.DrawString(footerText, font, brush,
                              (e.PageBounds.Width - footerSize.Width) / 2, footerY)

            ' გვერდები (ამჟამად მხოლოდ ერთი გვერდი)
            e.HasMorePages = False

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: PrintToPDF შეცდომა: {ex.Message}")
            e.HasMorePages = False
        End Try
    End Sub

    ''' <summary>
    ''' ტექსტის შემოკლება სვეტის სიგანის მიხედვით
    ''' </summary>
    Private Function TruncateText(text As String, maxWidth As Single, font As Font, graphics As Graphics) As String
        Try
            If String.IsNullOrEmpty(text) Then
                Return ""
            End If

            Dim textSize As SizeF = graphics.MeasureString(text, font)
            If textSize.Width <= maxWidth Then
                Return text
            End If

            ' ტექსტის შემოკლება
            For i = text.Length - 1 To 1 Step -1
                Dim shortText As String = text.Substring(0, i) & "..."
                Dim shortSize As SizeF = graphics.MeasureString(shortText, font)
                If shortSize.Width <= maxWidth Then
                    Return shortText
                End If
            Next

            Return text.Substring(0, 1) & "..."

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: TruncateText შეცდომა: {ex.Message}")
            Return text
        End Try
    End Function

    ''' <summary>
    ''' HTML ფაილის შექმნა ბეჭდვისთვის (ალტერნატივა PDF-ისთვის)
    ''' </summary>
    Private Sub CreateHTMLForPrinting(filePath As String)
        Try
            Debug.WriteLine($"UC_Schedule: HTML ფაილის შექმნა ბეჭდვისთვის - {filePath}")

            Dim html As New System.Text.StringBuilder()

            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html>")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>განრიგის მონაცემები</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        @media print { body { margin: 0; } .no-print { display: none; } }")
            html.AppendLine("        body { font-family: Arial, sans-serif; font-size: 12px; }")
            html.AppendLine("        h1 { text-align: center; font-size: 18px; margin: 10px 0; }")
            html.AppendLine("        table { width: 100%; border-collapse: collapse; font-size: 10px; }")
            html.AppendLine("        th, td { padding: 4px 2px; border: 1px solid #333; text-align: left; }")
            html.AppendLine("        th { background-color: #ddd; font-weight: bold; }")
            html.AppendLine("        tr:nth-child(even) { background-color: #f9f9f9; }")
            html.AppendLine("        .info { text-align: center; margin: 10px 0; font-size: 11px; }")
            html.AppendLine("    </style>")
            html.AppendLine("</head>")
            html.AppendLine("<body>")

            html.AppendLine("    <h1>განრიგის მონაცემები</h1>")
            html.AppendLine($"    <div class=""info"">თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm} | ჩანაწერები: {DgvSchedule.Rows.Count}</div>")

            ' ავტომატური ბეჭდვის ღილაკი
            html.AppendLine("    <div class=""no-print"" style=""text-align: center; margin: 20px;"">")
            html.AppendLine("        <button onclick=""window.print(); window.close();"" style=""padding: 15px 25px; font-size: 16px; background: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer;"">")
            html.AppendLine("            🖨️ ბეჭდვა PDF-ად</button>")
            html.AppendLine("        <p>ღილაკზე დაჭერის შემდეგ აირჩიეთ ""Microsoft Print to PDF""</p>")
            html.AppendLine("    </div>")

            html.AppendLine("    <table>")
            html.AppendLine("        <thead><tr>")

            ' სათაურები
            For Each column As DataGridViewColumn In DgvSchedule.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    html.AppendLine($"            <th>{EscapeHtml(column.HeaderText)}</th>")
                End If
            Next

            html.AppendLine("        </tr></thead>")
            html.AppendLine("        <tbody>")

            ' მონაცემები
            For rowIndex As Integer = 0 To DgvSchedule.Rows.Count - 1
                html.AppendLine("            <tr>")
                Dim row As DataGridViewRow = DgvSchedule.Rows(rowIndex)

                For Each column As DataGridViewColumn In DgvSchedule.Columns
                    If column.Visible AndAlso column.Name <> "Edit" Then
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        html.AppendLine($"                <td>{EscapeHtml(cellValue)}</td>")
                    End If
                Next

                html.AppendLine("            </tr>")
            Next

            html.AppendLine("        </tbody>")
            html.AppendLine("    </table>")
            html.AppendLine($"    <div class=""info"">შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm} | Scheduler v8.8a</div>")
            html.AppendLine("</body>")
            html.AppendLine("</html>")

            ' ფაილის ჩაწერა
            System.IO.File.WriteAllText(filePath, html.ToString(), System.Text.Encoding.UTF8)

            Debug.WriteLine("UC_Schedule: HTML ფაილი ბეჭდვისთვის შეიქმნა")
            MessageBox.Show($"შეიქმნა ფაილი ბეჭდვისთვის:{Environment.NewLine}{filePath}" & Environment.NewLine & Environment.NewLine &
                           "ფაილი იხსნება ბრაუზერში. დააჭირეთ 'ბეჭდვა PDF-ად' ღილაკს და აირჩიეთ 'Microsoft Print to PDF'",
                           "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' ფაილის გახსნა
            System.Diagnostics.Process.Start(filePath)

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreateHTMLForPrinting შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' HTML ტექსტის escape
    ''' </summary>
    Private Function EscapeHtml(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;")
    End Function

    ''' <summary>
    ''' HTML ფაილის ექსპორტი - მარტივი ვერსია
    ''' </summary>
    Private Sub ExportToSimpleHTML()
        Try
            Debug.WriteLine("UC_Schedule: HTML ექსპორტი")

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "HTML ფაილები (*.html)|*.html"
                saveDialog.Title = "HTML ფაილის შენახვა"
                saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.html"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    CreateSimpleHTMLFile(saveDialog.FileName)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportToSimpleHTML შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' CSV ფაილის ექსპორტი - მარტივი ვერსია
    ''' </summary>
    Private Sub ExportToSimpleCSV()
        Try
            Debug.WriteLine("UC_Schedule: CSV ექსპორტი")

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "CSV ფაილები (*.csv)|*.csv"
                saveDialog.Title = "CSV ფაილის შენახვა"
                saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.csv"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    CreateSimpleCSVFile(saveDialog.FileName)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ExportToSimpleCSV შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' მარტივი HTML ფაილის შექმნა
    ''' </summary>
    Private Sub CreateSimpleHTMLFile(filePath As String)
        Try
            Debug.WriteLine($"UC_Schedule: HTML ფაილის შექმნა - {filePath}")

            Dim html As New System.Text.StringBuilder()

            ' HTML დოკუმენტის საწყისი
            html.AppendLine("<!DOCTYPE html>")
            html.AppendLine("<html lang=""ka"">")
            html.AppendLine("<head>")
            html.AppendLine("    <meta charset=""UTF-8"">")
            html.AppendLine("    <title>განრიგის მონაცემები</title>")
            html.AppendLine("    <style>")
            html.AppendLine("        body { font-family: Arial, sans-serif; margin: 20px; }")
            html.AppendLine("        h1 { text-align: center; color: #333; }")
            html.AppendLine("        table { width: 100%; border-collapse: collapse; margin-top: 20px; }")
            html.AppendLine("        th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }")
            html.AppendLine("        th { background-color: #f2f2f2; font-weight: bold; }")
            html.AppendLine("        tr:nth-child(even) { background-color: #f9f9f9; }")
            html.AppendLine("        .info { text-align: center; margin: 20px 0; color: #666; }")
            html.AppendLine("        @media print { .no-print { display: none; } }")
            html.AppendLine("    </style>")
            html.AppendLine("</head>")
            html.AppendLine("<body>")

            ' სათაური
            html.AppendLine("    <h1>განრიგის მონაცემები</h1>")

            ' ინფორმაცია
            html.AppendLine("    <div class=""info"">")
            html.AppendLine($"        <p>თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm} | ჩანაწერები: {DgvSchedule.Rows.Count}</p>")
            html.AppendLine("    </div>")

            ' ღილაკი ბეჭდვისთვის
            html.AppendLine("    <div class=""no-print"" style=""text-align: center; margin: 20px 0;"">")
            html.AppendLine("        <button onclick=""window.print()"" style=""padding: 10px 20px; font-size: 16px;"">🖨️ ბეჭდვა</button>")
            html.AppendLine("    </div>")

            ' ცხრილი
            html.AppendLine("    <table>")

            ' სათაურები - მხოლოდ ხილული სვეტები (Edit-ის გარდა)
            html.AppendLine("        <thead>")
            html.AppendLine("            <tr>")
            For Each column As DataGridViewColumn In DgvSchedule.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    html.AppendLine($"                <th>{EscapeHtmlText(column.HeaderText)}</th>")
                End If
            Next
            html.AppendLine("            </tr>")
            html.AppendLine("        </thead>")

            ' მონაცემები
            html.AppendLine("        <tbody>")
            For rowIndex As Integer = 0 To DgvSchedule.Rows.Count - 1
                Dim row As DataGridViewRow = DgvSchedule.Rows(rowIndex)
                html.AppendLine("            <tr>")

                For Each column As DataGridViewColumn In DgvSchedule.Columns
                    If column.Visible AndAlso column.Name <> "Edit" Then
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        html.AppendLine($"                <td>{EscapeHtmlText(cellValue)}</td>")
                    End If
                Next

                html.AppendLine("            </tr>")
            Next
            html.AppendLine("        </tbody>")
            html.AppendLine("    </table>")

            ' ქვედა ინფორმაცია
            html.AppendLine("    <div class=""info"">")
            html.AppendLine($"        <p>შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm} | Scheduler v8.8a</p>")
            html.AppendLine("    </div>")

            html.AppendLine("</body>")
            html.AppendLine("</html>")

            ' ფაილის ჩაწერა
            System.IO.File.WriteAllText(filePath, html.ToString(), System.Text.Encoding.UTF8)

            Debug.WriteLine("UC_Schedule: HTML ფაილი წარმატებით შეიქმნა")
            MessageBox.Show($"HTML ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}", "წარმატება",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' ფაილის გახსნის შეთავაზება
            Dim openResult As DialogResult = MessageBox.Show(
                "გსურთ HTML ფაილის გახსნა ბრაუზერში?",
                "ფაილის გახსნა",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If openResult = DialogResult.Yes Then
                System.Diagnostics.Process.Start(filePath)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreateSimpleHTMLFile შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' მარტივი CSV ფაილის შექმნა
    ''' </summary>
    Private Sub CreateSimpleCSVFile(filePath As String)
        Try
            Debug.WriteLine($"UC_Schedule: CSV ფაილის შექმნა - {filePath}")

            Dim csv As New System.Text.StringBuilder()

            ' სათაურები - მხოლოდ ხილული სვეტები (Edit-ის გარდა)
            Dim headers As New List(Of String)
            For Each column As DataGridViewColumn In DgvSchedule.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    headers.Add(EscapeCSVText(column.HeaderText))
                End If
            Next
            csv.AppendLine(String.Join(",", headers))

            ' მონაცემები
            For rowIndex As Integer = 0 To DgvSchedule.Rows.Count - 1
                Dim row As DataGridViewRow = DgvSchedule.Rows(rowIndex)
                Dim rowData As New List(Of String)

                For Each column As DataGridViewColumn In DgvSchedule.Columns
                    If column.Visible AndAlso column.Name <> "Edit" Then
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        rowData.Add(EscapeCSVText(cellValue))
                    End If
                Next

                csv.AppendLine(String.Join(",", rowData))
            Next

            ' ფაილის ჩაწერა UTF-8 BOM-ით (Excel-ისთვის)
            Dim utf8WithBom As New System.Text.UTF8Encoding(True)
            System.IO.File.WriteAllText(filePath, csv.ToString(), utf8WithBom)

            Debug.WriteLine("UC_Schedule: CSV ფაილი წარმატებით შეიქმნა")
            MessageBox.Show($"CSV ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}" & Environment.NewLine & Environment.NewLine &
                           "ფაილი შეგიძლიათ გახსნათ Microsoft Excel-ში", "წარმატება",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' ფაილის გახსნის შეთავაზება
            Dim openResult As DialogResult = MessageBox.Show(
                "გსურთ CSV ფაილის გახსნა?",
                "ფაილის გახსნა",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If openResult = DialogResult.Yes Then
                System.Diagnostics.Process.Start(filePath)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: CreateSimpleCSVFile შეცდომა: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' CSV ტექსტის escape
    ''' </summary>
    Private Function EscapeCSVText(text As String) As String
        Try
            If String.IsNullOrEmpty(text) Then
                Return ""
            End If

            ' თუ შეიცავს კომას, ახალ ხაზს ან ციტატებს
            If text.Contains(",") OrElse text.Contains(vbCrLf) OrElse text.Contains(vbLf) OrElse text.Contains("""") Then
                ' ციტატების გადვოება
                text = text.Replace("""", """""")
                ' მთლიანი ველის ციტატებში ჩასმა
                text = """" & text & """"
            End If

            Return text
        Catch
            Return text
        End Try
    End Function

#End Region

End Class