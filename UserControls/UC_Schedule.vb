' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განრიგის UserControl - აჩვენებს და მართავს სესიების განრიგს
' ფურცლების ნავიგაციით და ჩანაწერების რაოდენობის ფილტრაციით
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Schedule
    Inherits UserControl

    ' მონაცემთა სერვისი
    Private dataService As IDataService = Nothing

    ' მომხმარებლის ელფოსტა რედაქტირებისთვის
    Private userEmail As String = ""

    ' ჩატვირთული სესიების სია
    Private allSessions As List(Of SessionModel)
    Private filteredSessions As List(Of SessionModel)

    ' ფურცლების ნავიგაციისთვის
    Private currentPage As Integer = 1
    Private pageSize As Integer = 20 ' ნაგულისხმები ზომა
    Private totalPages As Integer = 1

    ''' <summary>
    ''' მონაცემთა სერვისის და მომხმარებლის email-ის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი (IDataService)</param>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    Public Sub SetDataService(service As IDataService, Optional email As String = "")
        dataService = service
        userEmail = email
        Debug.WriteLine($"UC_Schedule.SetDataService: მონაცემთა სერვისი დაყენებულია, მომხმარებელი: {email}")

        ' თუ კონტროლი უკვე ჩატვირთულია, ვტვირთავთ მონაცემებს
        If Me.IsHandleCreated Then
            LoadSessionsData()
        End If
    End Sub

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_Schedule_Load: კონტროლის ჩატვირთვა დაიწყო")

            ' ფონის ფერების დაყენება
            SetBackgroundColors()

            ' საწყისი თარიღების დაყენება
            InitializeDatePickers()

            ' DataGridView-ის კონფიგურაცია
            ConfigureDataGridView()

            ' რადიობუტონების ინიციალიზაცია
            InitializeRadioButtons()

            ' ნავიგაციის ღილაკების ინიციალიზაცია
            InitializeNavigationButtons()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვტვირთავთ მონაცემებს
            If dataService IsNot Nothing Then
                LoadSessionsData()
            End If

            Debug.WriteLine("UC_Schedule_Load: კონტროლის ჩატვირთვა დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule_Load: შეცდომა - {ex.Message}")
            MessageBox.Show($"კონტროლის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ფონის ფერების დაყენება
    ''' </summary>
    Private Sub SetBackgroundColors()
        Try
            Debug.WriteLine("SetBackgroundColors: ფონის ფერების დაყენება")

            ' ღია გამჭვირვალე თეთრი ფერი
            Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

            ' პანელებისა და გრუპბოქსების ფერები
            pnlFilter.BackColor = transparentWhite
            GBSumInf.BackColor = transparentWhite
            GBSumFin.BackColor = transparentWhite

            Debug.WriteLine("SetBackgroundColors: ფონის ფერები დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"SetBackgroundColors: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თარიღების DatePicker-ების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeDatePickers()
        Try
            Debug.WriteLine("InitializeDatePickers: თარიღების ველების ინიციალიზაცია")

            ' მიმდინარე თარიღი
            Dim currentDate As DateTime = DateTime.Today

            ' მიმდინარე თვის პირველი დღე
            Dim firstDayOfMonth As DateTime = New DateTime(currentDate.Year, currentDate.Month, 1)

            ' მიმდინარე თვის ბოლო დღე
            Dim lastDayOfMonth As DateTime = firstDayOfMonth.AddMonths(1).AddDays(-1)

            ' DatePicker-ების დაყენება
            DtpDan.Value = firstDayOfMonth
            DtpMde.Value = lastDayOfMonth

            Debug.WriteLine($"InitializeDatePickers: საწყისი თარიღი: {firstDayOfMonth:dd.MM.yyyy}, " &
                           $"ბოლო თარიღი: {lastDayOfMonth:dd.MM.yyyy}")

            ' თარიღების შეცვლის ივენთები
            AddHandler DtpDan.ValueChanged, AddressOf DatePickers_ValueChanged
            AddHandler DtpMde.ValueChanged, AddressOf DatePickers_ValueChanged

        Catch ex As Exception
            Debug.WriteLine($"InitializeDatePickers: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' რადიობუტონების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeRadioButtons()
        Try
            Debug.WriteLine("InitializeRadioButtons: რადიობუტონების ინიციალიზაცია")

            ' RB20-ს ღირს ნაგულისხმევად
            RB20.Checked = True
            pageSize = 20

            ' ივენთების მიბმა
            AddHandler RB20.CheckedChanged, AddressOf RadioButton_CheckedChanged
            AddHandler RB50.CheckedChanged, AddressOf RadioButton_CheckedChanged
            AddHandler RB100.CheckedChanged, AddressOf RadioButton_CheckedChanged

            Debug.WriteLine("InitializeRadioButtons: რადიობუტონები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"InitializeRadioButtons: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureDataGridView()
        Try
            Debug.WriteLine("ConfigureDataGridView: DataGridView-ის კონფიგურაცია")

            ' ძირითადი პარამეტრები
            DgvSchedule.AutoGenerateColumns = False
            DgvSchedule.AllowUserToAddRows = False
            DgvSchedule.AllowUserToDeleteRows = False
            DgvSchedule.ReadOnly = True
            DgvSchedule.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DgvSchedule.MultiSelect = False

            ' თავის სტილი
            DgvSchedule.EnableHeadersVisualStyles = False
            DgvSchedule.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 80, 150)
            DgvSchedule.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            DgvSchedule.ColumnHeadersDefaultCellStyle.Font = New Font("Sylfaen", 10, FontStyle.Bold)

            ' მწკრივების სტილი
            DgvSchedule.RowsDefaultCellStyle.BackColor = Color.White
            DgvSchedule.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 250)
            DgvSchedule.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 200, 255)
            DgvSchedule.DefaultCellStyle.SelectionForeColor = Color.Black

            ' სვეტების შექმნა
            CreateDataGridViewColumns()

            Debug.WriteLine("ConfigureDataGridView: DataGridView კონფიგურაცია დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"ConfigureDataGridView: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის სვეტების შექმნა
    ''' </summary>
    Private Sub CreateDataGridViewColumns()
        Try
            Debug.WriteLine("CreateDataGridViewColumns: სვეტების შექმნა")

            ' გავასუფთავოთ არსებული სვეტები
            DgvSchedule.Columns.Clear()

            ' ID სვეტი
            Dim colId As New DataGridViewTextBoxColumn()
            colId.Name = "Id"
            colId.HeaderText = "ID"
            colId.DataPropertyName = "Id"
            colId.Width = 50
            colId.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DgvSchedule.Columns.Add(colId)

            ' თარიღი და დრო
            Dim colDateTime As New DataGridViewTextBoxColumn()
            colDateTime.Name = "DateTime"
            colDateTime.HeaderText = "თარიღი და დრო"
            colDateTime.DataPropertyName = "FormattedDateTime"
            colDateTime.Width = 130
            colDateTime.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DgvSchedule.Columns.Add(colDateTime)

            ' ხანგრძლივობა (მესამე სვეტი)
            Dim colDuration As New DataGridViewTextBoxColumn()
            colDuration.Name = "Duration"
            colDuration.HeaderText = "ხანგრძლივობა (წთ)"
            colDuration.DataPropertyName = "Duration"
            colDuration.Width = 80
            colDuration.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DgvSchedule.Columns.Add(colDuration)

            ' ბენეფიციარი
            Dim colBeneficiary As New DataGridViewTextBoxColumn()
            colBeneficiary.Name = "Beneficiary"
            colBeneficiary.HeaderText = "ბენეფიციარი"
            colBeneficiary.DataPropertyName = "FullName"
            colBeneficiary.Width = 150
            DgvSchedule.Columns.Add(colBeneficiary)

            ' თერაპევტი
            Dim colTherapist As New DataGridViewTextBoxColumn()
            colTherapist.Name = "Therapist"
            colTherapist.HeaderText = "თერაპევტი"
            colTherapist.DataPropertyName = "TherapistName"
            colTherapist.Width = 120
            DgvSchedule.Columns.Add(colTherapist)

            ' თერაპიის ტიპი
            Dim colTherapyType As New DataGridViewTextBoxColumn()
            colTherapyType.Name = "TherapyType"
            colTherapyType.HeaderText = "თერაპიის ტიპი"
            colTherapyType.DataPropertyName = "TherapyType"
            colTherapyType.Width = 100
            DgvSchedule.Columns.Add(colTherapyType)

            ' სივრცე
            Dim colSpace As New DataGridViewTextBoxColumn()
            colSpace.Name = "Space"
            colSpace.HeaderText = "სივრცე"
            colSpace.DataPropertyName = "Space"
            colSpace.Width = 80
            DgvSchedule.Columns.Add(colSpace)

            ' ფასი
            Dim colPrice As New DataGridViewTextBoxColumn()
            colPrice.Name = "Price"
            colPrice.HeaderText = "ფასი"
            colPrice.DataPropertyName = "Price"
            colPrice.Width = 70
            colPrice.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            colPrice.DefaultCellStyle.Format = "F2"
            DgvSchedule.Columns.Add(colPrice)

            ' სტატუსი
            Dim colStatus As New DataGridViewTextBoxColumn()
            colStatus.Name = "Status"
            colStatus.HeaderText = "სტატუსი"
            colStatus.DataPropertyName = "Status"
            colStatus.Width = 100
            colStatus.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DgvSchedule.Columns.Add(colStatus)

            ' დაფინანსება
            Dim colFunding As New DataGridViewTextBoxColumn()
            colFunding.Name = "Funding"
            colFunding.HeaderText = "დაფინანსება"
            colFunding.DataPropertyName = "Funding"
            colFunding.Width = 100
            DgvSchedule.Columns.Add(colFunding)

            ' ავტორი
            Dim colAuthor As New DataGridViewTextBoxColumn()
            colAuthor.Name = "Author"
            colAuthor.HeaderText = "ავტორი"
            colAuthor.DataPropertyName = "Author"
            colAuthor.Width = 100
            DgvSchedule.Columns.Add(colAuthor)

            ' რედაქტირების თარიღი
            Dim colEditDate As New DataGridViewTextBoxColumn()
            colEditDate.Name = "EditDate"
            colEditDate.HeaderText = "რედაქტირების თარიღი"
            colEditDate.DataPropertyName = "FormattedLastEditDate"
            colEditDate.Width = 130
            colEditDate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DgvSchedule.Columns.Add(colEditDate)

            ' რედაქტირების ღილაკი
            Dim colEdit As New DataGridViewButtonColumn()
            colEdit.Name = "EditButton"
            colEdit.HeaderText = "რედაქტირება"
            colEdit.Text = "✎"
            colEdit.UseColumnTextForButtonValue = True
            colEdit.Width = 80
            colEdit.DefaultCellStyle.BackColor = Color.FromArgb(60, 80, 150)
            colEdit.DefaultCellStyle.ForeColor = Color.White
            colEdit.DefaultCellStyle.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
            DgvSchedule.Columns.Add(colEdit)

            ' რედაქტირების ღილაკის ივენთი
            AddHandler DgvSchedule.CellClick, AddressOf DgvSchedule_CellClick

            Debug.WriteLine($"CreateDataGridViewColumns: შეიქმნა {DgvSchedule.Columns.Count} სვეტი")

        Catch ex As Exception
            Debug.WriteLine($"CreateDataGridViewColumns: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeNavigationButtons()
        Try
            Debug.WriteLine("InitializeNavigationButtons: ნავიგაციის ღილაკების ინიციალიზაცია")

            ' ღილაკების სტილი
            BtnPrev.FlatStyle = FlatStyle.Flat
            BtnNext.FlatStyle = FlatStyle.Flat
            BtnPrev.BackColor = Color.FromArgb(60, 80, 150)
            BtnNext.BackColor = Color.FromArgb(60, 80, 150)
            BtnPrev.ForeColor = Color.White
            BtnNext.ForeColor = Color.White

            ' ღილაკების ტექსტი მოშორებული
            BtnPrev.Text = ""
            BtnNext.Text = ""

            ' ღილაკების ივენთები
            AddHandler BtnPrev.Click, AddressOf BtnPrev_Click
            AddHandler BtnNext.Click, AddressOf BtnNext_Click

            ' გვერდის ლეიბლის ინიციალიზაცია
            UpdatePageLabel()

            Debug.WriteLine("InitializeNavigationButtons: ნავიგაციის ღილაკები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"InitializeNavigationButtons: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadSessionsData()
        Try
            Debug.WriteLine("LoadSessionsData: სესიების მონაცემების ჩატვირთვა დაიწყო")

            If dataService Is Nothing Then
                Debug.WriteLine("LoadSessionsData: მონაცემთა სერვისი არ არის დაყენებული")
                Return
            End If

            ' ყველა სესიის მიღება
            allSessions = dataService.GetAllSessions()

            Debug.WriteLine($"LoadSessionsData: ჩატვირთულია {If(allSessions Is Nothing, 0, allSessions.Count)} სესია")

            ' ფილტრაცია თარიღების მიხედვით
            ApplyDateFilter()

        Catch ex As Exception
            Debug.WriteLine($"LoadSessionsData: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიების ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თარიღების ფილტრის გამოყენება
    ''' </summary>
    Private Sub ApplyDateFilter()
        Try
            Debug.WriteLine("ApplyDateFilter: თარიღების ფილტრაცია")

            If allSessions Is Nothing Then
                filteredSessions = New List(Of SessionModel)()
            Else
                ' ფილტრაცია თარიღების დიაპაზონის მიხედვით
                Dim startDate As DateTime = DtpDan.Value.Date
                Dim endDate As DateTime = DtpMde.Value.Date.AddDays(1).AddTicks(-1) ' დღის ბოლომდე

                filteredSessions = allSessions.Where(Function(s) s.DateTime >= startDate AndAlso s.DateTime <= endDate).ToList()

                Debug.WriteLine($"ApplyDateFilter: ფილტრაციის შემდეგ დარჩა {filteredSessions.Count} სესია " &
                               $"({startDate:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy})")
            End If

            ' ფურცლების ნავიგაციის განახლება
            currentPage = 1
            CalculateTotalPages()

            ' DataGridView-ის განახლება
            UpdateDataGridView()

            ' სტატისტიკის განახლება
            UpdateStatistics()

        Catch ex As Exception
            Debug.WriteLine($"ApplyDateFilter: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მთლიანი გვერდების რაოდენობის გამოთვლა
    ''' </summary>
    Private Sub CalculateTotalPages()
        Try
            If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                totalPages = 1
            Else
                totalPages = Math.Ceiling(filteredSessions.Count / pageSize)
            End If

            Debug.WriteLine($"CalculateTotalPages: სულ გვერდები: {totalPages}, ჩანაწერები: {If(filteredSessions Is Nothing, 0, filteredSessions.Count)}, გვერდის ზომა: {pageSize}")

        Catch ex As Exception
            Debug.WriteLine($"CalculateTotalPages: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის განახლება ფილტრირებული მონაცემებით
    ''' </summary>
    Private Sub UpdateDataGridView()
        Try
            Debug.WriteLine("UpdateDataGridView: DataGridView-ის განახლება")

            If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                DgvSchedule.DataSource = Nothing
                Debug.WriteLine("UpdateDataGridView: მონაცემები ცარიელია")
                Return
            End If

            ' მიმდინარე გვერდის მონაცემების მიღება
            Dim startIndex As Integer = (currentPage - 1) * pageSize
            Dim endIndex As Integer = Math.Min(startIndex + pageSize - 1, filteredSessions.Count - 1)

            Dim pageData As List(Of SessionModel) = filteredSessions.GetRange(startIndex, endIndex - startIndex + 1)

            ' მონაცემების წყაროს დაყენება
            DgvSchedule.DataSource = Nothing
            DgvSchedule.DataSource = pageData

            ' სტატუსის ფერების დამატება
            ApplyStatusColors()

            Debug.WriteLine($"UpdateDataGridView: DataGridView განახლდა {pageData.Count} ჩანაწერით (გვერდი {currentPage}/{totalPages})")

            ' გვერდის ლეიბლის განახლება
            UpdatePageLabel()

            ' ღილაკების მდგომარეობის განახლება
            UpdateNavigationButtons()

        Catch ex As Exception
            Debug.WriteLine($"UpdateDataGridView: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის ფერების გამოყენება მწკრივებზე
    ''' </summary>
    Private Sub ApplyStatusColors()
        Try
            For Each row As DataGridViewRow In DgvSchedule.Rows
                If row.DataBoundItem IsNot Nothing Then
                    Dim session As SessionModel = CType(row.DataBoundItem, SessionModel)

                    ' სტატუსის ფერის მიღება
                    Dim statusColor As Color = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)

                    ' მწკრივის ფონის ფერი (რედაქტირების ღილაკის გარდა)
                    For i As Integer = 0 To row.Cells.Count - 2 ' ბოლო სვეტის გარდა
                        row.Cells(i).Style.BackColor = statusColor

                        ' ტექსტის ფერი - მუქი ფონისთვის ღია ტექსტი
                        If statusColor.GetBrightness() < 0.5 Then
                            row.Cells(i).Style.ForeColor = Color.White
                        Else
                            row.Cells(i).Style.ForeColor = Color.Black
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"ApplyStatusColors: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატისტიკის განახლება
    ''' </summary>
    Private Sub UpdateStatistics()
        Try
            Debug.WriteLine("UpdateStatistics: სტატისტიკის განახლება")

            If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                ' ცარიელი სტატისტიკა
                ClearStatistics()
                Return
            End If

            ' ძირითადი სტატისტიკა
            Dim totalSessions As Integer = filteredSessions.Count
            Dim completedSessions As Integer = filteredSessions.Where(Function(s) s.Status = "შესრულებული").Count()
            Dim cancelledSessions As Integer = filteredSessions.Where(Function(s) s.Status = "გაუქმებული").Count()
            Dim plannedSessions As Integer = filteredSessions.Where(Function(s) s.Status = "დაგეგმილი").Count()

            ' ფინანსური სტატისტიკა
            Dim totalRevenue As Decimal = filteredSessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)
            Dim potentialRevenue As Decimal = filteredSessions.Sum(Function(s) s.Price)

            Debug.WriteLine($"UpdateStatistics: სულ={totalSessions}, შესრულებული={completedSessions}, " &
                           $"გაუქმებული={cancelledSessions}, შემოსავალი={totalRevenue:F2}")

            ' სტატისტიკის განახლება UI-ზე (ამ ეტაპზე მხოლოდ Debug)
            Debug.WriteLine("UpdateStatistics: სტატისტიკა განახლებულია")

        Catch ex As Exception
            Debug.WriteLine($"UpdateStatistics: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატისტიკის გასუფთავება
    ''' </summary>
    Private Sub ClearStatistics()
        Try
            Debug.WriteLine("ClearStatistics: სტატისტიკის გასუფთავება")
            ' თანდათან დავამატებთ UI ელემენტების გასუფთავებას
        Catch ex As Exception
            Debug.WriteLine($"ClearStatistics: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვერდის ლეიბლის განახლება
    ''' </summary>
    Private Sub UpdatePageLabel()
        Try
            LPage.Text = $"{currentPage} / {totalPages}"
            Debug.WriteLine($"UpdatePageLabel: {LPage.Text}")
        Catch ex As Exception
            Debug.WriteLine($"UpdatePageLabel: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების მდგომარეობის განახლება
    ''' </summary>
    Private Sub UpdateNavigationButtons()
        Try
            ' წინა ღილაკი
            BtnPrev.Enabled = (currentPage > 1)

            ' შემდეგი ღილაკი
            BtnNext.Enabled = (currentPage < totalPages)

            Debug.WriteLine($"UpdateNavigationButtons: წინა={BtnPrev.Enabled}, შემდეგი={BtnNext.Enabled}")

        Catch ex As Exception
            Debug.WriteLine($"UpdateNavigationButtons: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' რადიობუტონების შეცვლის ივენთი
    ''' </summary>
    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim radioButton As RadioButton = CType(sender, RadioButton)

            If Not radioButton.Checked Then Return

            Debug.WriteLine($"RadioButton_CheckedChanged: {radioButton.Name} არჩეულია")

            ' გვერდის ზომის განახლება
            Select Case radioButton.Name
                Case "RB20"
                    pageSize = 20
                Case "RB50"
                    pageSize = 50
                Case "RB100"
                    pageSize = 100
            End Select

            Debug.WriteLine($"RadioButton_CheckedChanged: ახალი გვერდის ზომა: {pageSize}")

            ' გვერდების ხელახალი გამოთვლა
            currentPage = 1
            CalculateTotalPages()

            ' DataGridView-ის განახლება
            UpdateDataGridView()

        Catch ex As Exception
            Debug.WriteLine($"RadioButton_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თარიღების შეცვლის ივენთი
    ''' </summary>
    Private Sub DatePickers_ValueChanged(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine($"DatePickers_ValueChanged: {DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")

            ' თუ მონაცემები ჩატვირთულია, ვიყენებთ ფილტრაციას
            If allSessions IsNot Nothing Then
                ApplyDateFilter()
            End If

        Catch ex As Exception
            Debug.WriteLine($"DatePickers_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკი
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("BtnPrev_Click: წინა გვერდზე გადასვლა")

            If currentPage > 1 Then
                currentPage -= 1
                UpdateDataGridView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnPrev_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკი
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("BtnNext_Click: შემდეგ გვერდზე გადასვლა")

            If currentPage < totalPages Then
                currentPage += 1
                UpdateDataGridView()
            End If

        Catch ex As Exception
            Debug.WriteLine($"BtnNext_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub DgvSchedule_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' შევამოწმოთ რედაქტირების სვეტზე დაჭერა
            If e.ColumnIndex = DgvSchedule.Columns("EditButton").Index AndAlso e.RowIndex >= 0 Then
                ' მივიღოთ სესიის ID
                Dim sessionId As Integer = CInt(DgvSchedule.Rows(e.RowIndex).Cells("Id").Value)

                Debug.WriteLine($"DgvSchedule_CellClick: რედაქტირების ღილაკზე დაჭერა სესია ID={sessionId}")

                ' NewRecordForm-ის გახსნა რედაქტირების რეჟიმში
                Try
                    ' მოვძებნოთ მომხმარებლის email
                    Dim currentUserEmail As String = If(String.IsNullOrEmpty(userEmail), "user@example.com", userEmail)

                    ' NewRecordForm-ის შექმნა რედაქტირების რეჟიმში
                    ' გამოვიყენოთ არსებული კონსტრუქტორი რედაქტირებისთვის
                    Using editForm As New NewRecordForm(dataService, "სესია", sessionId, currentUserEmail, "UC_Schedule")

                        ' ფორმის ჩვენება
                        Dim result As DialogResult = editForm.ShowDialog()

                        ' თუ წარმატებით შეინახა, განვაახლოთ მონაცემები
                        If result = DialogResult.OK Then
                            RefreshData()
                            Debug.WriteLine($"DgvSchedule_CellClick: სესია {sessionId} წარმატებით განახლდა")
                        End If
                    End Using

                Catch formEx As Exception
                    Debug.WriteLine($"DgvSchedule_CellClick: NewRecordForm-ის გახსნის შეცდომა - {formEx.Message}")
                    MessageBox.Show($"რედაქტირების ფორმის გახსნის შეცდომა: {formEx.Message}", "შეცდომა",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"DgvSchedule_CellClick: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების ხელახალი ჩატვირთვა
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("RefreshData: მონაცემების ხელახალი ჩატვირთვა")

            If dataService IsNot Nothing Then
                LoadSessionsData()
            End If

        Catch ex As Exception
            Debug.WriteLine($"RefreshData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონკრეტულ გვერდზე გადასვლა
    ''' </summary>
    ''' <param name="pageNumber">გვერდის ნომერი</param>
    Public Sub GoToPage(pageNumber As Integer)
        Try
            If pageNumber >= 1 AndAlso pageNumber <= totalPages Then
                currentPage = pageNumber
                UpdateDataGridView()
                Debug.WriteLine($"GoToPage: გადავედი {pageNumber} გვერდზე")
            End If
        Catch ex As Exception
            Debug.WriteLine($"GoToPage: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მიმდინარე გვერდის ნომრის მიღება
    ''' </summary>
    Public ReadOnly Property CurrentPageNumber As Integer
        Get
            Return currentPage
        End Get
    End Property

    ''' <summary>
    ''' სულ გვერდების რაოდენობის მიღება
    ''' </summary>
    Public ReadOnly Property TotalPagesCount As Integer
        Get
            Return totalPages
        End Get
    End Property
End Class