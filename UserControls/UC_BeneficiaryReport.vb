' ===========================================
' 📄 UserControls/UC_BeneficiaryReport.vb - განახლებული ვერსია ინვოისის პრევიუთი
' -------------------------------------------
' ბენეფიციარის ინვოისის ფორმის UserControl
' ✨ ახალი ფუნქციონალი:
' - ეკრანის გაყოფა: 70% DataGridView, 30% ინვოისის პრევიუ
' - რეალური ინვოისის ფორმატის ნიმუში PDF-ის მიხედვით
' - ხანგძლიობის სვეტის ზომის გაზრდა 5 პიკსელით
' 
' 🎯 ხელმისაწვდომია როლებისთვის: 1 (ადმინი), 2 (მენეჯერი), 3
' ===========================================
Imports System.ComponentModel
'Imports Scheduler_v8_8a.Services
'Imports Scheduler_v8_8a.Models
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

    ' 🔧 ციკლური ივენთების თავიდან აცილება - გამართული ლოგიკა
    Private isLoadingData As Boolean = False
    Private isUpdatingBeneficiaryName As Boolean = False      ' სახელის ComboBox განახლებისთვის
    Private isUpdatingBeneficiarySurname As Boolean = False   ' გვარის ComboBox განახლებისთვის
    Private isUpdatingTherapistData As Boolean = False        ' თერაპევტის/თერაპიის ComboBox განახლებისთვის

    ' ✨ ახალი UI კონტროლები ინვოისის პრევიუსთვის
    Private WithEvents pnlMain As Panel                       ' მთავარი კონტეინერი
    Private WithEvents pnlLeft As Panel                       ' მარცხენა პანელი (DataGrid)
    Private WithEvents pnlRight As Panel                      ' მარჯვენა პანელი (პრევიუ)
    Private WithEvents splitter As Splitter                   ' განმყოფი

    ' ინვოისის პრევიუს კონტროლები
    Private WithEvents lblInvoiceTitle As Label               ' ინვოისის სათაური
    Private WithEvents lblCompanyInfo As RichTextBox          ' კომპანიის ინფორმაცია
    Private WithEvents lblBeneficiaryInfo As RichTextBox      ' ბენეფიციარის ინფორმაცია
    Private WithEvents dgvInvoicePreview As DataGridView      ' ინვოისის ცხრილის პრევიუ
    Private WithEvents lblTotalAmount As Label                ' ჯამური თანხა
    Private WithEvents lblTotalWords As Label                 ' თანხა სიტყვიერად
    Private WithEvents lblSignature As Label                  ' ხელმოწერა

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

            ' ✨ ახალი UI-ის შექმნა
            CreateSplitLayoutUI()

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
    ''' ✨ ახალი გაყოფილი UI-ის შექმნა - 70% DataGrid, 30% ინვოისის პრევიუ
    ''' </summary>
    Private Sub CreateSplitLayoutUI()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: გაყოფილი UI-ის შექმნა")

            ' მთავარი კონტეინერის შექმნა
            pnlMain = New Panel()
            pnlMain.Dock = DockStyle.Fill
            pnlMain.BorderStyle = BorderStyle.None
            Me.Controls.Add(pnlMain)

            ' მარცხენა პანელი - DataGridView-სთვის (70%)
            pnlLeft = New Panel()
            pnlLeft.Dock = DockStyle.Left
            pnlLeft.Width = CInt(Me.Width * 0.7) ' 70%
            pnlLeft.BorderStyle = BorderStyle.FixedSingle
            pnlLeft.BackColor = Color.FromArgb(245, 245, 250)
            pnlMain.Controls.Add(pnlLeft)

            ' განმყოფი
            splitter = New Splitter()
            splitter.Dock = DockStyle.Left
            splitter.Width = 5
            splitter.BackColor = Color.Gray
            splitter.BorderStyle = BorderStyle.Fixed3D
            pnlMain.Controls.Add(splitter)

            ' მარჯვენა პანელი - ინვოისის პრევიუსთვის (30%)
            pnlRight = New Panel()
            pnlRight.Dock = DockStyle.Fill ' დანარჩენი ადგილი
            pnlRight.BorderStyle = BorderStyle.FixedSingle
            pnlRight.BackColor = Color.White
            pnlRight.AutoScroll = True
            pnlMain.Controls.Add(pnlRight)

            ' არსებული კონტროლების მარცხენა პანელში გადატანა
            MoveExistingControlsToLeftPanel()

            ' ინვოისის პრევიუს შექმნა მარჯვენა პანელში
            CreateInvoicePreviewUI()

            Debug.WriteLine("UC_BeneficiaryReport: გაყოფილი UI წარმატებით შეიქმნა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateSplitLayoutUI შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ არსებული კონტროლების მარცხენა პანელში გადატანა
    ''' </summary>
    Private Sub MoveExistingControlsToLeftPanel()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: არსებული კონტროლების გადატანა")

            ' ყველა არსებული კონტროლის პოვნა და მარცხენა პანელში გადატანა
            Dim controlsToMove As New List(Of Control)()

            For Each ctrl As Control In Me.Controls
                If ctrl IsNot pnlMain Then
                    controlsToMove.Add(ctrl)
                End If
            Next

            ' კონტროლების გადატანა
            For Each ctrl In controlsToMove
                Me.Controls.Remove(ctrl)
                pnlLeft.Controls.Add(ctrl)
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: გადატანილია {controlsToMove.Count} კონტროლი")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: MoveExistingControlsToLeftPanel შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის პრევიუს UI-ის შექმნა PDF ნიმუშის მიხედვით
    ''' </summary>
    Private Sub CreateInvoicePreviewUI()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის პრევიუს UI-ის შექმნა")

            Dim yPos As Integer = 10
            Dim margin As Integer = 10

            ' ✨ ინვოისის სათაური
            lblInvoiceTitle = New Label()
            lblInvoiceTitle.Text = "ინვოისი მომსახურების გაწევაზე"
            lblInvoiceTitle.Font = New Font("Sylfaen", 12, FontStyle.Bold)
            lblInvoiceTitle.Location = New Point(margin, yPos)
            lblInvoiceTitle.Size = New Size(pnlRight.Width - margin * 2, 25)
            lblInvoiceTitle.TextAlign = ContentAlignment.MiddleCenter
            lblInvoiceTitle.ForeColor = Color.DarkBlue
            pnlRight.Controls.Add(lblInvoiceTitle)
            yPos += 35

            ' ✨ კომპანიის ინფორმაცია
            lblCompanyInfo = New RichTextBox()
            lblCompanyInfo.Text = GetCompanyInfoText()
            lblCompanyInfo.Font = New Font("Sylfaen", 8, FontStyle.Regular)
            lblCompanyInfo.Location = New Point(margin, yPos)
            lblCompanyInfo.Size = New Size(pnlRight.Width - margin * 2, 80)
            lblCompanyInfo.ReadOnly = True
            lblCompanyInfo.BorderStyle = BorderStyle.None
            lblCompanyInfo.BackColor = pnlRight.BackColor
            pnlRight.Controls.Add(lblCompanyInfo)
            yPos += 90

            ' ✨ ბენეფიციარის ინფორმაცია
            lblBeneficiaryInfo = New RichTextBox()
            lblBeneficiaryInfo.Text = "ბენეფიციარი: [არჩეული ბენეფიციარი]" & vbCrLf &
                                    "დაბადების თარიღი:" & vbCrLf &
                                    "საბანკო რეკვიზიტები: პ/ნ:" & vbCrLf &
                                    "კანონიერი წარმომადგენელი:" & vbCrLf &
                                    "წარმომადგენლის პ/ნ:"
            lblBeneficiaryInfo.Font = New Font("Sylfaen", 8, FontStyle.Regular)
            lblBeneficiaryInfo.Location = New Point(margin, yPos)
            lblBeneficiaryInfo.Size = New Size(pnlRight.Width - margin * 2, 70)
            lblBeneficiaryInfo.ReadOnly = True
            lblBeneficiaryInfo.BorderStyle = BorderStyle.FixedSingle
            lblBeneficiaryInfo.BackColor = Color.FromArgb(250, 250, 250)
            pnlRight.Controls.Add(lblBeneficiaryInfo)
            yPos += 80

            ' ✨ ინვოისის ცხრილის პრევიუ
            dgvInvoicePreview = New DataGridView()
            dgvInvoicePreview.Location = New Point(margin, yPos)
            dgvInvoicePreview.Size = New Size(pnlRight.Width - margin * 2, 200)
            dgvInvoicePreview.ReadOnly = True
            dgvInvoicePreview.AllowUserToAddRows = False
            dgvInvoicePreview.AllowUserToDeleteRows = False
            dgvInvoicePreview.RowHeadersVisible = False
            dgvInvoicePreview.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgvInvoicePreview.BackgroundColor = Color.White
            dgvInvoicePreview.GridColor = Color.Black
            dgvInvoicePreview.Font = New Font("Sylfaen", 8, FontStyle.Regular)

            ' ინვოისის ცხრილის სვეტები
            CreateInvoicePreviewColumns()

            pnlRight.Controls.Add(dgvInvoicePreview)
            yPos += 210

            ' ✨ ჯამური თანხა
            lblTotalAmount = New Label()
            lblTotalAmount.Text = "მომსახურების საფასური სულ: 0.00 ₾"
            lblTotalAmount.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            lblTotalAmount.Location = New Point(margin, yPos)
            lblTotalAmount.Size = New Size(pnlRight.Width - margin * 2, 25)
            lblTotalAmount.TextAlign = ContentAlignment.MiddleRight
            lblTotalAmount.ForeColor = Color.DarkGreen
            pnlRight.Controls.Add(lblTotalAmount)
            yPos += 35

            ' ✨ თანხა სიტყვიერად
            lblTotalWords = New Label()
            lblTotalWords.Text = "თანხა სიტყვიერად: ნული ლარი, ნული თეთრი"
            lblTotalWords.Font = New Font("Sylfaen", 9, FontStyle.Italic)
            lblTotalWords.Location = New Point(margin, yPos)
            lblTotalWords.Size = New Size(pnlRight.Width - margin * 2, 25)
            lblTotalWords.TextAlign = ContentAlignment.MiddleLeft
            pnlRight.Controls.Add(lblTotalWords)
            yPos += 35

            ' ✨ ხელმოწერა
            lblSignature = New Label()
            lblSignature.Text = GetSignatureText()
            lblSignature.Font = New Font("Sylfaen", 8, FontStyle.Regular)
            lblSignature.Location = New Point(margin, yPos)
            lblSignature.Size = New Size(pnlRight.Width - margin * 2, 60)
            lblSignature.TextAlign = ContentAlignment.BottomCenter
            pnlRight.Controls.Add(lblSignature)

            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის პრევიუს UI წარმატებით შეიქმნა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateInvoicePreviewUI შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის პრევიუს ცხრილის სვეტების შექმნა
    ''' </summary>
    Private Sub CreateInvoicePreviewColumns()
        Try
            dgvInvoicePreview.Columns.Clear()

            With dgvInvoicePreview.Columns
                .Add("N", "N")           ' ნომერი
                .Add("Date", "თარიღი")   ' თარიღი
                .Add("Therapy", "თერაპია") ' თერაპია
                .Add("Therapist", "თერაპევტი") ' თერაპევტი
                .Add("Amount", "თანხა")   ' თანხა
            End With

            ' სვეტების სიგანეების დაყენება
            With dgvInvoicePreview.Columns
                .Item("N").Width = 30
                .Item("Date").Width = 80
                .Item("Therapy").Width = 100
                .Item("Therapist").Width = 80
                .Item("Amount").Width = 50

                ' ფორმატირება
                .Item("Amount").DefaultCellStyle.Format = "N2"
                .Item("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Item("N").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის პრევიუს სვეტები შეიქმნა")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateInvoicePreviewColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ კომპანიის ინფორმაციის ტექსტის მიღება
    ''' </summary>
    Private Function GetCompanyInfoText() As String
        Return "შპს ""ბავშვთა და მოზარდთა განვითარების, აბილიტაციისა და" & vbCrLf &
               "რეაბილიტაციის ცენტრი - პროსპერო""" & vbCrLf &
               "მომსახურების განწევი: [კომპანიის რეკვიზიტები]" & vbCrLf &
               "მომსახურების მიმღები: [მიმღების ინფორმაცია]"
    End Function

    ''' <summary>
    ''' ✨ ხელმოწერის ტექსტის მიღება
    ''' </summary>
    Private Function GetSignatureText() As String
        Return vbCrLf & "ცენტრის დირექტორი:" & vbCrLf &
               "თეა ჩანადირი MD PhD DBP" & vbCrLf &
               "მედიცინის დოქტორი,განვითარების და ქცევის პედიატრი"
    End Function

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
                Nothing, Nothing, Nothing, Nothing, Nothing, Nothing ' ComboBox-ები ხელით მოვამუშავებთ
            )

            ' სტატუსის CheckBox-ების ინიციალიზაცია
            InitializeStatusCheckBoxes()

            ' ფილტრების ინიციალიზაცია (თარიღების და სტატუსების)
            filterManager.InitializeFilters()

            ' 🔧 მხოლოდ FilterChanged ივენთი გვჭირდება
            AddHandler filterManager.FilterChanged, AddressOf OnFilterChanged

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
                Dim checkBox As CheckBox = FindCheckBoxRecursive(pnlLeft, $"CheckBox{i}")
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
    ''' 🔧 DataGridView-ის კონფიგურაცია ინვოისის ფორმისთვის - განახლებული ვერსია
    ''' ✨ ხანგძლიობის სვეტის ზომა გაზრდილია 5 პიკსელით
    ''' </summary>
    Private Sub ConfigureDataGridView()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: DataGridView-ის კონფიგურაცია (განახლებული)")

            If DgvSessions Is Nothing Then
                Debug.WriteLine("UC_BeneficiaryReport: DgvSessions არ არსებობს")
                Return
            End If

            With DgvSessions
                .AutoGenerateColumns = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .ReadOnly = False  ' 🔧 ჩეკბოქსების რედაქტირებისთვის
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .MultiSelect = True  ' 🔧 მრავალი მწკრივის მონიშვნისთვის

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

                ' 🔧 ჩეკბოქსების რედაქტირების ჩართვა
                .EditMode = DataGridViewEditMode.EditOnKeystroke
            End With

            ' სვეტების შექმნა ინვოისისთვის
            CreateInvoiceColumns()

            Debug.WriteLine("UC_BeneficiaryReport: DataGridView კონფიგურაცია დასრულდა (განახლებული)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ConfigureDataGridView შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ინვოისის სვეტების შექმნა - განახლებული ვერსია
    ''' ✨ ხანგძლიობის სვეტის ზომა 45 პიკსელი (40+5)
    ''' </summary>
    Private Sub CreateInvoiceColumns()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის სვეტების შექმნა (განახლებული)")

            DgvSessions.Columns.Clear()

            With DgvSessions.Columns
                ' ინვოისის ნომერი (არა ბაზის ID)
                .Add("InvoiceN", "N")

                ' თარიღი
                .Add("DateTime", "თარიღი")

                ' ხანგძლივობა - ✨ ზომა გაზრდილია 5 პიკსელით
                .Add("Duration", "ხანგძლ.")

                ' თერაპია
                .Add("TherapyType", "თერაპია")

                ' თერაპევტი
                .Add("Therapist", "თერაპევტი")

                ' 🆕 დაფინანსების სვეტი
                .Add("Funding", "დაფინანსება")

                ' 🆕 სტატუსის სვეტი
                .Add("Status", "სტატუსი")

                ' თანხა
                .Add("Price", "თანხა")

                ' 🆕 ჩეკბოქსის სვეტი - ინვოისში ჩართვის/ამოღების კონტროლი
                Dim excludeCheckColumn As New DataGridViewCheckBoxColumn()
                excludeCheckColumn.Name = "ExcludeFromInvoice"
                excludeCheckColumn.HeaderText = "ამოღება"
                excludeCheckColumn.ToolTipText = "მონიშნე ინვოისიდან ამოსაღებად"
                excludeCheckColumn.Width = 70
                excludeCheckColumn.FalseValue = False
                excludeCheckColumn.TrueValue = True
                excludeCheckColumn.ThreeState = False
                .Add(excludeCheckColumn)

                ' 🆕 რედაქტირების ღილაკი
                Dim editBtn As New DataGridViewButtonColumn()
                editBtn.Name = "EditSession"
                editBtn.HeaderText = ""
                editBtn.Text = "✎"
                editBtn.UseColumnTextForButtonValue = True
                editBtn.ToolTipText = "სესიის რედაქტირება"
                editBtn.Width = 35
                .Add(editBtn)
            End With

            ' სვეტების სიგანეების დაყენება
            SetInvoiceColumnWidths()

            Debug.WriteLine($"UC_BeneficiaryReport: შეიქმნა {DgvSessions.Columns.Count} სვეტი ინვოისისთვის (ჩეკბოქსი და რედაქტირების ღილაკით)")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CreateInvoiceColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ინვოისის სვეტების სიგანეების დაყენება - განახლებული ვერსია
    ''' ✨ ხანგძლიობის სვეტი 45 პიკსელი (40+5)
    ''' </summary>
    Private Sub SetInvoiceColumnWidths()
        Try
            With DgvSessions.Columns
                .Item("InvoiceN").Width = 50          ' N
                .Item("DateTime").Width = 120         ' თარიღი - ოდნავ შემცირებული
                .Item("Duration").Width = 45          ' ✨ ხანგძლივობა - 40+5=45 პიკსელი
                .Item("TherapyType").Width = 180      ' თერაპია - ოდნავ შემცირებული
                .Item("Therapist").Width = 150        ' თერაპევტი - ოდნავ შემცირებული
                .Item("Funding").Width = 100          ' 🆕 დაფინანსება
                .Item("Status").Width = 120           ' 🆕 სტატუსი
                .Item("Price").Width = 80             ' თანხა
                .Item("ExcludeFromInvoice").Width = 70  ' 🆕 ჩეკბოქსი
                .Item("EditSession").Width = 35       ' 🆕 რედაქტირება

                ' ფასის სვეტის ფორმატი
                .Item("Price").DefaultCellStyle.Format = "N2"
                .Item("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                ' სტატუსის სვეტის ფორმატი
                .Item("Status").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                ' ჩეკბოქსის სვეტის ფორმატი
                .Item("ExcludeFromInvoice").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Item("ExcludeFromInvoice").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                ' რედაქტირების ღილაკის ფორმატი
                .Item("EditSession").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის სვეტების სიგანეები დაყენებულია (ხანგძლივობა +5 პიკსელი)")

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

            ' 🔧 ComboBox-ების საწყისი მდგომარეობა - გამართული ლოგიკით
            ResetAllComboBoxes()

            ' DataGridView-ის გასუფთავება
            DgvSessions.Rows.Clear()

            ' ✨ ინვოისის პრევიუს გასუფთავება
            ClearInvoicePreview()

            Debug.WriteLine("UC_BeneficiaryReport: საწყისი ფილტრები დაყენებულია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SetInitialFilters შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ყველა ComboBox-ის რესეტი - გამართული ვერსია
    ''' </summary>
    Private Sub ResetAllComboBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ResetAllComboBoxes")

            ' Flag-ების დაყენება
            isUpdatingBeneficiaryName = True
            isUpdatingBeneficiarySurname = True
            isUpdatingTherapistData = True

            Try
                ' სახელების ComboBox რესეტი
                CBBeneName.Items.Clear()
                CBBeneName.Text = ""
                CBBeneName.SelectedIndex = -1
                CBBeneName.Enabled = False

                ' გვარების ComboBox რესეტი
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Text = ""
                CBBeneSurname.SelectedIndex = -1
                CBBeneSurname.Enabled = False

                ' თერაპევტის ComboBox რესეტი
                CBPer.Items.Clear()
                CBPer.Text = ""
                CBPer.SelectedIndex = -1
                CBPer.Enabled = False

                ' თერაპიის ComboBox რესეტი
                CBTer.Items.Clear()
                CBTer.Text = ""
                CBTer.SelectedIndex = -1
                CBTer.Enabled = False

                Debug.WriteLine("UC_BeneficiaryReport: ყველა ComboBox გასუფთავებულია")

            Finally
                ' Flag-ების ჩართვა ყოველთვის
                isUpdatingBeneficiaryName = False
                isUpdatingBeneficiarySurname = False
                isUpdatingTherapistData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResetAllComboBoxes შეცდომა: {ex.Message}")
            ' Flag-ების ჩართვა შეცდომის შემთხვევაშიც
            isUpdatingBeneficiaryName = False
            isUpdatingBeneficiarySurname = False
            isUpdatingTherapistData = False
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის პრევიუს გასუფთავება
    ''' </summary>
    Private Sub ClearInvoicePreview()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის პრევიუს გასუფთავება")

            ' ბენეფიციარის ინფორმაციის რესეტი
            If lblBeneficiaryInfo IsNot Nothing Then
                lblBeneficiaryInfo.Text = "ბენეფიციარი: [არჩეული ბენეფიციარი]" & vbCrLf &
                                        "დაბადების თარიღი:" & vbCrLf &
                                        "საბანკო რეკვიზიტები: პ/ნ:" & vbCrLf &
                                        "კანონიერი წარმომადგენელი:" & vbCrLf &
                                        "წარმომადგენლის პ/ნ:"
            End If

            ' ცხრილის გასუფთავება
            If dgvInvoicePreview IsNot Nothing Then
                dgvInvoicePreview.Rows.Clear()
            End If

            ' ჯამური თანხის რესეტი
            If lblTotalAmount IsNot Nothing Then
                lblTotalAmount.Text = "მომსახურების საფასური სულ: 0.00 ₾"
            End If

            ' თანხა სიტყვიერად რესეტი
            If lblTotalWords IsNot Nothing Then
                lblTotalWords.Text = "თანხა სიტყვიერად: ნული ლარი, ნული თეთრი"
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ClearInvoicePreview შეცდომა: {ex.Message}")
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

#Region "ფილტრაციის ლოგიკა - გამართული ვერსია"

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

                ' 🔧 ბენეფიციარის სახელების ComboBox-ის განახლება
                PopulateBeneficiaryNamesComboBox()

                ' მონაცემების გასუფთავება (სახელი ახალი არჩევისთვის)
                ClearBeneficiarySpecificData()

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
    ''' 🔧 ბენეფიციარის სახელების ComboBox-ის შევსება - გამართული ვერსია
    ''' </summary>
    Private Sub PopulateBeneficiaryNamesComboBox()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox")

            If CBBeneName Is Nothing OrElse currentFilteredSessions Is Nothing Then Return

            ' 🔧 ტემპორარული flag განახლებისთვის
            isUpdatingBeneficiaryName = True

            Try
                ' 🔧 მიმდინარე არჩეული მნიშვნელობის შენახვა
                Dim selectedValue As String = CBBeneName.SelectedItem?.ToString()

                CBBeneName.Items.Clear()

                If currentFilteredSessions.Count = 0 Then
                    CBBeneName.Enabled = False
                    CBBeneName.Text = ""
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

                ' 🔧 არჩეული მნიშვნელობის აღდგენა (თუ ჯერ კიდევ არსებობს)
                If Not String.IsNullOrEmpty(selectedValue) Then
                    For i As Integer = 0 To CBBeneName.Items.Count - 1
                        If String.Equals(CBBeneName.Items(i).ToString(), selectedValue, StringComparison.OrdinalIgnoreCase) Then
                            CBBeneName.SelectedIndex = i
                            Debug.WriteLine($"UC_BeneficiaryReport: აღდგენილია სახელი: '{selectedValue}'")
                            Exit For
                        End If
                    Next
                End If

                Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {uniqueNames.Count} სახელი")

            Finally
                isUpdatingBeneficiaryName = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PopulateBeneficiaryNamesComboBox შეცდომა: {ex.Message}")
            isUpdatingBeneficiaryName = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ბენეფიციარის სპეციფიკური მონაცემების გასუფთავება
    ''' </summary>
    Private Sub ClearBeneficiarySpecificData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ClearBeneficiarySpecificData")

            ' მონაცემების გასუფთავება
            currentBeneficiaryData = New List(Of SessionModel)()
            LoadSessionsToGrid(currentBeneficiaryData)

            ' ✨ ინვოისის პრევიუს გასუფთავება
            UpdateInvoicePreview()

            ' გვარის, თერაპევტისა და თერაპიის ComboBox-ების გასუფთავება
            isUpdatingBeneficiarySurname = True
            isUpdatingTherapistData = True

            Try
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Text = ""
                CBBeneSurname.SelectedIndex = -1
                CBBeneSurname.Enabled = False

                CBPer.Items.Clear()
                CBPer.Text = ""
                CBPer.SelectedIndex = -1
                CBPer.Enabled = False

                CBTer.Items.Clear()
                CBTer.Text = ""
                CBTer.SelectedIndex = -1
                CBTer.Enabled = False

            Finally
                isUpdatingBeneficiarySurname = False
                isUpdatingTherapistData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ClearBeneficiarySpecificData შეცდომა: {ex.Message}")
            isUpdatingBeneficiarySurname = False
            isUpdatingTherapistData = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ბენეფიციარის გვარების ComboBox-ის განახლება - გამართული ვერსია
    ''' </summary>
    Public Sub UpdateBeneficiarySurnames(selectedName As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateBeneficiarySurnames - სახელი: '{selectedName}'")

            If CBBeneSurname Is Nothing OrElse currentFilteredSessions Is Nothing Then Return

            isUpdatingBeneficiarySurname = True

            Try
                CBBeneSurname.Items.Clear()

                If String.IsNullOrWhiteSpace(selectedName) Then
                    CBBeneSurname.Enabled = False
                    CBBeneSurname.Text = ""
                    CBBeneSurname.SelectedIndex = -1

                    ' თერაპევტისა და თერაპიის ComboBox-ების გასუფთავება
                    ClearTherapistAndTherapyComboBoxes()

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

            Finally
                isUpdatingBeneficiarySurname = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateBeneficiarySurnames შეცდომა: {ex.Message}")
            isUpdatingBeneficiarySurname = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპევტისა და თერაპიის ComboBox-ების გასუფთავება
    ''' </summary>
    Private Sub ClearTherapistAndTherapyComboBoxes()
        Try
            isUpdatingTherapistData = True

            Try
                CBPer.Items.Clear()
                CBPer.Text = ""
                CBPer.SelectedIndex = -1
                CBPer.Enabled = False

                CBTer.Items.Clear()
                CBTer.Text = ""
                CBTer.SelectedIndex = -1
                CBTer.Enabled = False

            Finally
                isUpdatingTherapistData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ClearTherapistAndTherapyComboBoxes შეცდომა: {ex.Message}")
            isUpdatingTherapistData = False
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

                ' თერაპევტისა და თერაპიის ComboBox-ების გასუფთავება
                ClearTherapistAndTherapyComboBoxes()
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

            ' ✨ ინვოისის პრევიუს განახლება
            UpdateInvoicePreview()

            ' თერაპევტების და თერაპიების ComboBox-ების განახლება
            UpdateTherapistAndTherapyComboBoxes()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: FilterBeneficiarySpecificData შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 სესიების ჩატვირთვა DataGridView-ში ინვოისის ფორმატით - განახლებული ვერსია
    ''' ✨ ინვოისის პრევიუს ავტომატური განახლება
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

                ' 🔧 მწკრივის მონაცემები
                Dim rowData As Object() = {
                    invoiceNumber,                                  ' ინვოისის N
                    session.DateTime.ToString("dd.MM.yyyy HH:mm"), ' თარიღი
                    $"{session.Duration}წთ",                        ' ხანგძლივობა - მოკლე ფორმატი
                    session.TherapyType,                            ' თერაპია
                    session.TherapistName,                          ' თერაპევტი
                    session.Funding,                                ' დაფინანსება
                    session.Status,                                 ' სტატუსი
                    session.Price,                                  ' თანხა
                    False,                                          ' ჩეკბოქსი - ნაგულისხმევად არ არის მონიშნული
                    "✎"                                             ' რედაქტირების ღილაკი
                }

                Dim addedRowIndex = DgvSessions.Rows.Add(rowData)

                ' 🔧 სესიის ID-ის შენახვა მწკრივში - რედაქტირებისთვის
                DgvSessions.Rows(addedRowIndex).Tag = session.Id

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                Try
                    Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                    DgvSessions.Rows(addedRowIndex).DefaultCellStyle.BackColor = statusColor
                Catch
                    ' სტატუსის ფერის დაყენების შეცდომის შემთხვევაში ნაგულისხმევი
                End Try
            Next

            ' 🔧 ჩეკბოქსის ივენთის მიბმა
            AddHandler DgvSessions.CellValueChanged, AddressOf DgvSessions_CellValueChanged
            AddHandler DgvSessions.CellClick, AddressOf DgvSessions_CellClick

            ' ✨ ინვოისის პრევიუს ავტომატური განახლება
            UpdateInvoicePreview()

            Debug.WriteLine($"UC_BeneficiaryReport: ჩატვირთულია {sessions.Count} სესია DataGridView-ში")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: LoadSessionsToGrid შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპევტების და თერაპიების ComboBox-ების განახლება - გამართული ვერსია
    ''' </summary>
    Private Sub UpdateTherapistAndTherapyComboBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: UpdateTherapistAndTherapyComboBoxes")

            If currentBeneficiaryData Is Nothing OrElse currentBeneficiaryData.Count = 0 Then
                ClearTherapistAndTherapyComboBoxes()
                Return
            End If

            isUpdatingTherapistData = True

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
                isUpdatingTherapistData = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateTherapistAndTherapyComboBoxes შეცდომა: {ex.Message}")
            isUpdatingTherapistData = False
        End Try
    End Sub

#End Region

#Region "ივენთ ჰენდლერები - გამართული ვერსია"

    ''' <summary>
    ''' ფილტრის შეცვლის ივენთი - FilterManager-დან
    ''' </summary>
    Private Sub OnFilterChanged()
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: OnFilterChanged - isLoadingData: {isLoadingData}")

            ' 🔧 მონაცემების დატვირთვის დროს არ ვუშვებთ ციკლს
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
    ''' 🔧 ბენეფიციარის სახელის ComboBox-ის შეცვლა - გამართული ვერსია
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            ' 🔧 ციკლური განახლებების თავიდან აცილება
            If isUpdatingBeneficiaryName Then
                Debug.WriteLine("UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged იგნორირებულია - სახელები განახლების პროცესშია")
                Return
            End If

            Debug.WriteLine("UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged")

            Dim actualSelectedName As String = If(CBBeneName?.SelectedItem?.ToString(), "")
            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარის სახელი შეიცვალა: '{actualSelectedName}'")

            ' ბენეფიციარის გვარების ComboBox-ის განახლება
            UpdateBeneficiarySurnames(actualSelectedName)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneName_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ბენეფიციარის გვარის ComboBox-ის შეცვლა - გამართული ვერსია
    ''' </summary>
    Private Sub CBBeneSurname_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneSurname.SelectedIndexChanged
        Try
            ' 🔧 ციკლური განახლებების თავიდან აცილება
            If isUpdatingBeneficiarySurname Then
                Debug.WriteLine("UC_BeneficiaryReport: CBBeneSurname_SelectedIndexChanged იგნორირებულია - გვარები განახლების პროცესშია")
                Return
            End If

            Debug.WriteLine("UC_BeneficiaryReport: CBBeneSurname_SelectedIndexChanged")

            ' ბენეფიციარის სპეციფიკური მონაცემების განახლება
            FilterBeneficiarySpecificData()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneSurname_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპევტის ComboBox-ის შეცვლა - გამართული ვერსია
    ''' </summary>
    Private Sub CBPer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPer.SelectedIndexChanged
        Try
            ' 🔧 ციკლური განახლებების თავიდან აცილება
            If isUpdatingTherapistData Then
                Debug.WriteLine("UC_BeneficiaryReport: CBPer_SelectedIndexChanged იგნორირებულია - თერაპევტები განახლების პროცესშია")
                Return
            End If

            Debug.WriteLine("UC_BeneficiaryReport: CBPer_SelectedIndexChanged")

            ' ბენეფიციარის მონაცემების ხელახალი ფილტრაცია თერაპევტის მიხედვით
            ApplyTherapistAndTherapyFilter()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: CBPer_SelectedIndexChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპიის ComboBox-ის შეცვლა - გამართული ვერსია
    ''' </summary>
    Private Sub CBTer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTer.SelectedIndexChanged
        Try
            ' 🔧 ციკლური განახლებების თავიდან აცილება
            If isUpdatingTherapistData Then
                Debug.WriteLine("UC_BeneficiaryReport: CBTer_SelectedIndexChanged იგნორირებულია - თერაპიები განახლების პროცესშია")
                Return
            End If

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

#Region "✨ ინვოისის პრევიუს განახლების მეთოდები"

    ''' <summary>
    ''' ✨ ინვოისის პრევიუს მთლიანი განახლება
    ''' </summary>
    Private Sub UpdateInvoicePreview()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის პრევიუს განახლება")

            ' ბენეფიციარის ინფორმაციის განახლება
            UpdateBeneficiaryInfoInPreview()

            ' ინვოისის ცხრილის განახლება
            UpdateInvoiceTablePreview()

            ' ჯამური თანხის განახლება
            UpdateTotalAmountInPreview()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateInvoicePreview შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ბენეფიციარის ინფორმაციის განახლება პრევიუში
    ''' </summary>
    Private Sub UpdateBeneficiaryInfoInPreview()
        Try
            If lblBeneficiaryInfo Is Nothing Then Return

            Dim beneficiaryName = GetCurrentBeneficiaryName()

            If String.IsNullOrEmpty(beneficiaryName) Then
                lblBeneficiaryInfo.Text = "ბენეფიციარი: [არჩეული ბენეფიციარი]" & vbCrLf &
                                        "დაბადების თარიღი:" & vbCrLf &
                                        "საბანკო რეკვიზიტები: პ/ნ:" & vbCrLf &
                                        "კანონიერი წარმომადგენელი:" & vbCrLf &
                                        "წარმომადგენლის პ/ნ:"
            Else
                lblBeneficiaryInfo.Text = $"ბენეფიციარი: {beneficiaryName}" & vbCrLf &
                                        "დაბადების თარიღი:" & vbCrLf &
                                        "საბანკო რეკვიზიტები: პ/ნ:" & vbCrLf &
                                        "თიბისი ბანკი TBCBGE21" & vbCrLf &
                                        "GE17TB7742036020100004" & vbCrLf &
                                        "კანონიერი წარმომადგენელი:" & vbCrLf &
                                        "წარმომადგენლის პ/ნ:"
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateBeneficiaryInfoInPreview შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის ცხრილის პრევიუს განახლება
    ''' </summary>
    Private Sub UpdateInvoiceTablePreview()
        Try
            If dgvInvoicePreview Is Nothing OrElse DgvSessions Is Nothing Then Return

            dgvInvoicePreview.Rows.Clear()

            ' მხოლოდ ის სესიები, რომლებიც არ არის მონიშნული ამოღებისთვის
            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' ჩეკბოქსის შემოწმება
                    Dim isExcluded As Boolean = False
                    If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), isExcluded)
                    End If

                    ' თუ სესია არ არის გამორიცხული, დავამატოთ პრევიუში
                    If Not isExcluded Then
                        Dim previewRowData As Object() = {
                            row.Cells("InvoiceN").Value,                    ' N
                            row.Cells("DateTime").Value,                    ' თარიღი
                            row.Cells("TherapyType").Value,                 ' თერაპია
                            row.Cells("Therapist").Value,                   ' თერაპევტი
                            row.Cells("Price").Value                        ' თანხა
                        }

                        dgvInvoicePreview.Rows.Add(previewRowData)
                    End If

                Catch ex As Exception
                    Debug.WriteLine($"UC_BeneficiaryReport: UpdateInvoiceTablePreview მწკრივის შეცდომა: {ex.Message}")
                    Continue For
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის პრევიუში ჩატვირთულია {dgvInvoicePreview.Rows.Count} სესია")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateInvoiceTablePreview შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ჯამური თანხის განახლება პრევიუში
    ''' </summary>
    Private Sub UpdateTotalAmountInPreview()
        Try
            If lblTotalAmount Is Nothing OrElse lblTotalWords Is Nothing Then Return

            Dim totalAmount = GetInvoiceTotalAmount()

            ' ჯამური თანხის ლეიბლი
            lblTotalAmount.Text = $"მომსახურების საფასური სულ: {totalAmount:N2} ₾"

            ' თანხა სიტყვიერად
            lblTotalWords.Text = $"თანხა სიტყვიერად: {ConvertAmountToWords(totalAmount)}"

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateTotalAmountInPreview შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ თანხის სიტყვიერ ფორმად გარდაქმნა (მარტივი ვერსია)
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

#End Region

#Region "დამხმარე მეთოდები"

    ''' <summary>
    ''' 🔧 ინვოისის ჯამური თანხის გამოთვლა - გაუმჯობესებული ვერსია
    ''' მხოლოდ ის სესიები ითვლება, რომლებიც არ არის მონიშნული ამოღებისთვის
    ''' </summary>
    ''' <returns>ინვოისის ჯამური თანხა (გამორიცხული სესიების გარეშე)</returns>
    Public Function GetInvoiceTotalAmount() As Decimal
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim total As Decimal = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' 🔧 შევამოწმოთ ჩეკბოქსი - თუ მონიშნულია, არ ჩავითვალოთ
                    Dim isExcluded As Boolean = False
                    If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), isExcluded)
                    End If

                    ' თუ სესია არ არის მონიშნული ამოღებისთვის, ჩავითვალოთ
                    If Not isExcluded AndAlso row.Cells("Price").Value IsNot Nothing Then
                        Dim price As Decimal
                        If Decimal.TryParse(row.Cells("Price").Value.ToString(), price) Then
                            total += price
                        End If
                    End If

                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის ჯამური თანხა (გამორიცხულების გარეშე): {total:N2}")
            Return total

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetInvoiceTotalAmount შეცდომა: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 🔧 ინვოისის სესიების რაოდენობის მიღება - გაუმჯობესებული ვერსია
    ''' მხოლოდ ის სესიები ითვლება, რომლებიც არ არის მონიშნული ამოღებისთვის
    ''' </summary>
    ''' <returns>ინვოისში ჩართული სესიების რაოდენობა</returns>
    Public Function GetInvoiceSessionCount() As Integer
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim includedCount As Integer = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    ' 🔧 შევამოწმოთ ჩეკბოქსი - თუ მონიშნულია, არ ჩავითვალოთ
                    Dim isExcluded As Boolean = False
                    If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), isExcluded)
                    End If

                    ' თუ სესია არ არის მონიშნული ამოღებისთვის, ჩავითვალოთ
                    If Not isExcluded Then
                        includedCount += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისში ჩართული სესიები: {includedCount}")
            Return includedCount

        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 🆕 გამორიცხული სესიების რაოდენობის მიღება
    ''' </summary>
    ''' <returns>გამორიცხული სესიების რაოდენობა</returns>
    Public Function GetExcludedSessionCount() As Integer
        Try
            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then
                Return 0
            End If

            Dim excludedCount As Integer = 0

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim isExcluded As Boolean = False
                    If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), isExcluded)
                    End If

                    If isExcluded Then
                        excludedCount += 1
                    End If

                Catch
                    Continue For
                End Try
            Next

            Return excludedCount

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

#Region "🆕 DataGridView ივენთ ჰენდლერები - ჩეკბოქსი და რედაქტირება"

    ''' <summary>
    ''' 🔧 DataGridView უჯრაზე დაჭერის ივენთი - რედაქტირების ღილაკისთვის
    ''' </summary>
    Private Sub DgvSessions_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' შევამოწმოთ არის თუ არა ვალიდური მწკრივი და სვეტი
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

            ' 🔧 რედაქტირების ღილაკზე დაჭერა
            If DgvSessions.Columns(e.ColumnIndex).Name = "EditSession" Then
                Debug.WriteLine($"UC_BeneficiaryReport: რედაქტირების ღილაკზე დაჭერა - მწკრივი {e.RowIndex}")

                ' სესიის ID-ის მიღება მწკრივის Tag-დან
                Dim sessionId As Integer = 0
                If DgvSessions.Rows(e.RowIndex).Tag IsNot Nothing Then
                    Integer.TryParse(DgvSessions.Rows(e.RowIndex).Tag.ToString(), sessionId)
                End If

                If sessionId > 0 Then
                    EditSession(sessionId)
                Else
                    MessageBox.Show("სესიის ID ვერ მოიძებნა", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: DgvSessions_CellClick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 DataGridView უჯრის მნიშვნელობის შეცვლის ივენთი - ჩეკბოქსისთვის
    ''' ✨ ინვოისის პრევიუს ავტომატური განახლება
    ''' </summary>
    Private Sub DgvSessions_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' შევამოწმოთ არის თუ არა ჩეკბოქსის სვეტი
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
               DgvSessions.Columns(e.ColumnIndex).Name = "ExcludeFromInvoice" Then

                Debug.WriteLine($"UC_BeneficiaryReport: ჩეკბოქსი შეიცვალა - მწკრივი {e.RowIndex}")

                ' მწკრივის ვიზუალური სტილის განახლება
                UpdateRowVisualStyle(e.RowIndex)

                ' ✨ ინვოისის პრევიუს განახლება
                UpdateInvoicePreview()

                ' ინვოისის ჯამების ხელახალი გამოთვლა (თუ საჭიროა)
                OnInvoiceTotalsChanged()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: DgvSessions_CellValueChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 მწკრივის ვიზუალური სტილის განახლება ჩეკბოქსის მდგომარეობის მიხედვით
    ''' </summary>
    ''' <param name="rowIndex">მწკრივის ინდექსი</param>
    Private Sub UpdateRowVisualStyle(rowIndex As Integer)
        Try
            If rowIndex < 0 OrElse rowIndex >= DgvSessions.Rows.Count Then Return

            Dim row = DgvSessions.Rows(rowIndex)

            ' ჩეკბოქსის მდგომარეობის შემოწმება
            Dim isExcluded As Boolean = False
            If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), isExcluded)
            End If

            If isExcluded Then
                ' 🔧 მონიშნული მწკრივი - ნაცრისფერი ტექსტი და ღია ფონი
                row.DefaultCellStyle.ForeColor = Color.Gray
                row.DefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240)
                row.DefaultCellStyle.Font = New Font(DgvSessions.DefaultCellStyle.Font, FontStyle.Strikeout)

                Debug.WriteLine($"UC_BeneficiaryReport: მწკრივი {rowIndex} მონიშნულია ამოღებისთვის")
            Else
                ' 🔧 ჩვეულებრივი მწკრივი - სტანდარტული სტილი
                row.DefaultCellStyle.ForeColor = Color.Black
                row.DefaultCellStyle.Font = New Font(DgvSessions.DefaultCellStyle.Font, FontStyle.Regular)

                ' სტატუსის ფერის აღდგენა
                Try
                    If currentBeneficiaryData IsNot Nothing AndAlso rowIndex < currentBeneficiaryData.Count Then
                        Dim session = currentBeneficiaryData(rowIndex)
                        Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.BackColor = statusColor
                    Else
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                Catch
                    row.DefaultCellStyle.BackColor = Color.White
                End Try

                Debug.WriteLine($"UC_BeneficiaryReport: მწკრივი {rowIndex} აღდგენილია ჩვეულებრივ მდგომარეობაში")
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: UpdateRowVisualStyle შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 სესიის რედაქტირების ფუნქცია
    ''' </summary>
    ''' <param name="sessionId">სესიის ID</param>
    Private Sub EditSession(sessionId As Integer)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: სესიის რედაქტირება - ID={sessionId}")

            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' რედაქტირების ფორმის გახსნა - UC_Schedule-ის ანალოგია
            Using editForm As New NewRecordForm(dataService, "სესია", sessionId, userEmail, "UC_BeneficiaryReport")
                Dim result As DialogResult = editForm.ShowDialog()

                If result = DialogResult.OK Then
                    Debug.WriteLine($"UC_BeneficiaryReport: სესია ID={sessionId} წარმატებით განახლდა")

                    ' მონაცემების განახლება
                    RefreshData()

                    MessageBox.Show($"სესია წარმატებით განახლდა", "წარმატება",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: EditSession შეცდომა: {ex.Message}")
            MessageBox.Show($"სესიის რედაქტირების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ინვოისის ჯამების შეცვლის ივენთი - მომავალი გამოყენებისთვის
    ''' (მაგალითად Label-ების განახლება ან სხვა UI კომპონენტების ინფორმირება)
    ''' </summary>
    Private Sub OnInvoiceTotalsChanged()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის ჯამები შეიცვალა")

            ' აქ შეიძლება დავამატოთ ლოგიკა:
            ' - Label-ების განახლება ჯამური თანხით
            ' - სტატისტიკის პანელების განახლება
            ' - ივენთის გაშვება მშობელი ფორმისთვის

            Dim totalAmount = GetInvoiceTotalAmount()
            Dim includedSessions = GetInvoiceSessionCount()
            Dim excludedSessions = GetExcludedSessionCount()

            Debug.WriteLine($"UC_BeneficiaryReport: ჯამური თანხა: {totalAmount:N2}, ჩართული: {includedSessions}, გამორიცხული: {excludedSessions}")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: OnInvoiceTotalsChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ყველა სესიის ჩართვა ინვოისში (ყველა ჩეკბოქსის გასუფთავება)
    ''' </summary>
    Public Sub IncludeAllSessions()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სესიის ჩართვა ინვოისში")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("ExcludeFromInvoice").Value = False
                    UpdateRowVisualStyle(row.Index)
                Catch
                    Continue For
                End Try
            Next

            ' ✨ ინვოისის პრევიუს განახლება
            UpdateInvoicePreview()
            OnInvoiceTotalsChanged()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: IncludeAllSessions შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 ყველა სესიის ამოღება ინვოისიდან (ყველა ჩეკბოქსის მონიშვნა)
    ''' </summary>
    Public Sub ExcludeAllSessions()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სესიის ამოღება ინვოისიდან")

            If DgvSessions Is Nothing OrElse DgvSessions.Rows.Count = 0 Then Return

            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    row.Cells("ExcludeFromInvoice").Value = True
                    UpdateRowVisualStyle(row.Index)
                Catch
                    Continue For
                End Try
            Next

            ' ✨ ინვოისის პრევიუს განახლება
            UpdateInvoicePreview()
            OnInvoiceTotalsChanged()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ExcludeAllSessions შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🆕 მონიშნული სესიების ჩართვა/ამოღება ინვოისში/ინვოისიდან
    ''' </summary>
    Public Sub ToggleSelectedSessions()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: მონიშნული სესიების ჩართვა/ამოღება")

            If DgvSessions Is Nothing OrElse DgvSessions.SelectedRows.Count = 0 Then
                MessageBox.Show("აირჩიეთ სესიები ჩართვის/ამოღების კონტროლისთვის", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            For Each row As DataGridViewRow In DgvSessions.SelectedRows
                Try
                    ' მიმდინარე მდგომარეობის შებრუნება
                    Dim currentValue As Boolean = False
                    If row.Cells("ExcludeFromInvoice").Value IsNot Nothing Then
                        Boolean.TryParse(row.Cells("ExcludeFromInvoice").Value.ToString(), currentValue)
                    End If

                    row.Cells("ExcludeFromInvoice").Value = Not currentValue
                    UpdateRowVisualStyle(row.Index)

                Catch
                    Continue For
                End Try
            Next

            ' ✨ ინვოისის პრევიუს განახლება
            UpdateInvoicePreview()
            OnInvoiceTotalsChanged()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ToggleSelectedSessions შეცდომა: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "✨ ახალი საჯარო მეთოდები ინვოისის ექსპორტისთვის"

    ''' <summary>
    ''' ✨ ინვოისის PDF-ად ექსპორტი (მომავალში)
    ''' </summary>
    Public Sub ExportInvoiceToPDF()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის PDF ექსპორტი")

            ' ინვოისის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("ინვოისის ექსპორტისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' მომავალში: PDF ექსპორტის ლოგიკა
            MessageBox.Show("PDF ექსპორტის ფუნქციონალი მომავალში დაემატება", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ExportInvoiceToPDF შეცდომა: {ex.Message}")
            MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის ბეჭდვა
    ''' </summary>
    Public Sub PrintInvoice()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის ბეჭდვა")

            ' ინვოისის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("ინვოისის ბეჭდვისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' მომავალში: ბეჭდვის ლოგიკა
            MessageBox.Show("ინვოისის ბეჭდვის ფუნქციონალი მომავალში დაემატება", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: PrintInvoice შეცდომა: {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის HTML-ად ექსპორტი
    ''' </summary>
    Public Sub ExportInvoiceToHTML()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის HTML ექსპორტი")

            ' ინვოისის ვალიდობის შემოწმება
            If Not IsInvoiceValid() Then
                MessageBox.Show("ინვოისის ექსპორტისთვის აირჩიეთ ბენეფიციარი და მინიმუმ ერთი სესია", "ინფორმაცია",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' მომავალში: HTML ექსპორტის ლოგიკა
            MessageBox.Show("HTML ექსპორტის ფუნქციონალი მომავალში დაემატება", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ExportInvoiceToHTML შეცდომა: {ex.Message}")
            MessageBox.Show($"HTML ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის მონაცემების JSON ფორმატში მიღება (API-სთვის)
    ''' </summary>
    Public Function GetInvoiceDataAsJSON() As String
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ინვოისის JSON მონაცემები")

            If Not IsInvoiceValid() Then
                Return "{""error"": ""ინვოისი არავალიდურია""}"
            End If

            ' მომავალში: JSON სერიალიზაციის ლოგიკა
            Return "{""message"": ""JSON ექსპორტი მომავალში დაემატება""}"

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: GetInvoiceDataAsJSON შეცდომა: {ex.Message}")
            Return $"{{""error"": ""{ex.Message}""}}"
        End Try
    End Function

#End Region

#Region "✨ Splitter ივენთ ჰენდლერები"

    ''' <summary>
    ''' ✨ Splitter-ის მოძრაობისას მარჯვენა პანელის ელემენტების ზომის განახლება
    ''' </summary>
    Private Sub splitter_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles splitter.SplitterMoved
        Try
            Debug.WriteLine("UC_BeneficiaryReport: Splitter გადაადგილდა")

            ' ინვოისის პრევიუს ელემენტების ზომის განახლება
            ResizeInvoicePreviewElements()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: splitter_SplitterMoved შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ ინვოისის პრევიუს ელემენტების ზომის განახლება
    ''' </summary>
    Private Sub ResizeInvoicePreviewElements()
        Try
            If pnlRight Is Nothing Then Return

            Dim newWidth As Integer = pnlRight.Width - 20 ' მარჯინები

            ' ყველა ელემენტის ზომის განახლება
            If lblInvoiceTitle IsNot Nothing Then
                lblInvoiceTitle.Width = newWidth
            End If

            If lblCompanyInfo IsNot Nothing Then
                lblCompanyInfo.Width = newWidth
            End If

            If lblBeneficiaryInfo IsNot Nothing Then
                lblBeneficiaryInfo.Width = newWidth
            End If

            If dgvInvoicePreview IsNot Nothing Then
                dgvInvoicePreview.Width = newWidth
            End If

            If lblTotalAmount IsNot Nothing Then
                lblTotalAmount.Width = newWidth
            End If

            If lblTotalWords IsNot Nothing Then
                lblTotalWords.Width = newWidth
            End If

            If lblSignature IsNot Nothing Then
                lblSignature.Width = newWidth
            End If

            Debug.WriteLine($"UC_BeneficiaryReport: ინვოისის პრევიუს ელემენტები გადაზომეთილია - ახალი სიგანე: {newWidth}")

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: ResizeInvoicePreviewElements შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ✨ UserControl-ის ზომის შეცვლისას
    ''' </summary>
    Private Sub UC_BeneficiaryReport_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        Try
            ' მარცხენა პანელის ზომის განახლება (70%)
            If pnlLeft IsNot Nothing AndAlso Me.Width > 0 Then
                pnlLeft.Width = CInt(Me.Width * 0.7)
            End If

            ' ინვოისის პრევიუს ელემენტების ზომის განახლება
            ResizeInvoicePreviewElements()

        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport: SizeChanged შეცდომა: {ex.Message}")
        End Try
    End Sub

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

            ' ✨ ახალი UI კონტროლების განთავისუფლება
            lblInvoiceTitle?.Dispose()
            lblCompanyInfo?.Dispose()
            lblBeneficiaryInfo?.Dispose()
            dgvInvoicePreview?.Dispose()
            lblTotalAmount?.Dispose()
            lblTotalWords?.Dispose()
            lblSignature?.Dispose()

            pnlLeft?.Dispose()
            pnlRight?.Dispose()
            splitter?.Dispose()
            pnlMain?.Dispose()

        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

End Class