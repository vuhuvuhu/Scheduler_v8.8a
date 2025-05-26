' ===========================================
' 📄 UserControls/UC_TherapistReport.vb
' -------------------------------------------
' თერაპევტის ანგარიშის UserControl
' შექმნილია ხელით, Designer.vb-ს გარეშე
' 
' 🎯 ხელმისაწვდომია მხოლოდ როლებისთვის: 1 (ადმინი), 2 (მენეჯერი)
' ===========================================
Imports System.ComponentModel
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports System.Text
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

''' <summary>
''' თერაპევტის ანგარიშის UserControl
''' საშუალებას აძლევს მენეჯერებს და ადმინისტრატორებს მოძებნონ თერაპევტი 
''' და ნახონ მისი სესიების დეტალური ანგარიში და შემოსავლების ანალიზი
''' </summary>
Public Class UC_TherapistReport
    Inherits UserControl

    ' UI კონტროლები - ხელით შექმნილი
    Private WithEvents TxtSearch As TextBox
    Private WithEvents BtnSearch As Button
    Private WithEvents DgvSessions As DataGridView
    Private WithEvents BtnExportToExcel As Button
    Private WithEvents BtnPrint As Button
    Private WithEvents BtnGenerateReport As Button
    Private WithEvents BtnMonthlyAnalysis As Button

    ' ინფორმაციის ლეიბლები
    Private LblTherapistName As Label
    Private LblTotalSessions As Label
    Private LblCompletedSessions As Label
    Private LblCancelledSessions As Label
    Private LblPendingSessions As Label
    Private LblTotalRevenue As Label
    Private LblAveragePrice As Label
    Private LblBeneficiariesCount As Label
    Private LblTherapyTypes As Label
    Private LblMostCommonTherapy As Label
    Private LblActiveMonths As Label
    Private LblFirstSession As Label
    Private LblLastSession As Label

    ' სერვისები და მონაცემები
    Private dataService As IDataService
    Private currentTherapist As TherapistInfo
    Private therapistSessions As List(Of SessionModel)
    Private printService As AdvancedDataGridViewPrintService

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRole As String = ""

    ' თვისობრივი ანალიზისთვის
    Private monthlyStats As Dictionary(Of String, MonthlyStatistics)

    ''' <summary>
    ''' თერაპევტის ინფორმაციის კლასი
    ''' </summary>
    Public Class TherapistInfo
        Public Property Name As String
        Public Property TotalSessions As Integer
        Public Property CompletedSessions As Integer
        Public Property CancelledSessions As Integer
        Public Property PendingSessions As Integer
        Public Property TotalRevenue As Decimal
        Public Property AverageSessionPrice As Decimal
        Public Property LastSessionDate As DateTime?
        Public Property FirstSessionDate As DateTime?
        Public Property ActiveMonths As Integer
        Public Property TherapyTypes As List(Of String)
        Public Property MostCommonTherapyType As String
        Public Property BeneficiariesCount As Integer
    End Class

    ''' <summary>
    ''' თვიური სტატისტიკის კლასი
    ''' </summary>
    Public Class MonthlyStatistics
        Public Property Year As Integer
        Public Property Month As Integer
        Public Property SessionsCount As Integer
        Public Property Revenue As Decimal
        Public Property CompletedSessions As Integer
        Public Property CancelledSessions As Integer
        Public Property UniqueClients As Integer

        Public ReadOnly Property MonthName As String
            Get
                Dim georgianMonths() As String = {
                    "იანვარი", "თებერვალი", "მარტი", "აპრილი", "მაისი", "ივნისი",
                    "ივლისი", "აგვისტო", "სექტემბერი", "ოქტომბერი", "ნოემბერი", "დეკემბერი"
                }
                If Month >= 1 AndAlso Month <= 12 Then
                    Return georgianMonths(Month - 1)
                End If
                Return Month.ToString()
            End Get
        End Property

        Public ReadOnly Property DisplayName As String
            Get
                Return $"{MonthName} {Year}"
            End Get
        End Property
    End Class

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    Public Sub New()
        therapistSessions = New List(Of SessionModel)()
        monthlyStats = New Dictionary(Of String, MonthlyStatistics)()
        InitializeControls()
        Debug.WriteLine("UC_TherapistReport: ინიციალიზებულია")
    End Sub

    ''' <summary>
    ''' UI კონტროლების ხელით შექმნა და განთავსება
    ''' </summary>
    Private Sub InitializeControls()
        Try
            ' UserControl-ის ძირითადი პარამეტრები
            Me.Size = New Size(1100, 800)
            Me.BackColor = Color.FromArgb(240, 245, 250)
            Me.Font = New Font("Sylfaen", 10)

            ' ძიების სექცია
            CreateSearchSection()

            ' ინფორმაციის სექცია  
            CreateInfoSection()

            ' სესიების DataGridView
            CreateSessionsGrid()

            ' ღილაკების სექცია
            CreateButtonsSection()

            ' საწყისი მდგომარეობა
            ResetForm()

            Debug.WriteLine("UC_TherapistReport: UI კონტროლები შექმნილია")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport InitializeControls: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ძიების სექციის შექმნა
    ''' </summary>
    Private Sub CreateSearchSection()
        ' ძიების ლეიბლი
        Dim lblSearch As New Label()
        lblSearch.Text = "თერაპევტის ძიება (სახელი):"
        lblSearch.Location = New Point(20, 20)
        lblSearch.Size = New Size(300, 25)
        lblSearch.Font = New Font("Sylfaen", 11, FontStyle.Bold)
        Me.Controls.Add(lblSearch)

        ' ძიების TextBox
        TxtSearch = New TextBox()
        TxtSearch.Location = New Point(20, 50)
        TxtSearch.Size = New Size(300, 25)
        TxtSearch.Font = New Font("Sylfaen", 10)
        Me.Controls.Add(TxtSearch)

        ' ძიების ღილაკი
        BtnSearch = New Button()
        BtnSearch.Text = "ძიება"
        BtnSearch.Location = New Point(330, 48)
        BtnSearch.Size = New Size(80, 29)
        BtnSearch.Font = New Font("Sylfaen", 9, FontStyle.Bold)
        BtnSearch.UseVisualStyleBackColor = True
        Me.Controls.Add(BtnSearch)
    End Sub

    ''' <summary>
    ''' ინფორმაციის სექციის შექმნა
    ''' </summary>
    Private Sub CreateInfoSection()
        Dim yPos As Integer = 90

        ' თერაპევტის სახელი
        LblTherapistName = New Label()
        LblTherapistName.Text = "თერაპევტი: -"
        LblTherapistName.Location = New Point(20, yPos)
        LblTherapistName.Size = New Size(400, 25)
        LblTherapistName.Font = New Font("Sylfaen", 12, FontStyle.Bold)
        LblTherapistName.ForeColor = Color.DarkBlue
        Me.Controls.Add(LblTherapistName)

        yPos += 35

        ' სტატისტიკის პანელი
        Dim statsPanel As New Panel()
        statsPanel.Location = New Point(20, yPos)
        statsPanel.Size = New Size(1050, 120)
        statsPanel.BackColor = Color.FromArgb(220, 235, 250)
        statsPanel.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(statsPanel)

        ' სტატისტიკის ლეიბლების შექმნა (3 რიგად)
        ' პირველი რიგი
        LblTotalSessions = New Label()
        LblTotalSessions.Text = "სულ სესიები: 0"
        LblTotalSessions.Location = New Point(10, 10)
        LblTotalSessions.Size = New Size(160, 20)
        statsPanel.Controls.Add(LblTotalSessions)

        LblCompletedSessions = New Label()
        LblCompletedSessions.Text = "შესრულებული: 0"
        LblCompletedSessions.Location = New Point(180, 10)
        LblCompletedSessions.Size = New Size(160, 20)
        statsPanel.Controls.Add(LblCompletedSessions)

        LblCancelledSessions = New Label()
        LblCancelledSessions.Text = "გაუქმებული: 0"
        LblCancelledSessions.Location = New Point(350, 10)
        LblCancelledSessions.Size = New Size(160, 20)
        statsPanel.Controls.Add(LblCancelledSessions)

        LblPendingSessions = New Label()
        LblPendingSessions.Text = "დაგეგმილი: 0"
        LblPendingSessions.Location = New Point(520, 10)
        LblPendingSessions.Size = New Size(160, 20)
        statsPanel.Controls.Add(LblPendingSessions)

        LblActiveMonths = New Label()
        LblActiveMonths.Text = "აქტიური თვეები: 0"
        LblActiveMonths.Location = New Point(690, 10)
        LblActiveMonths.Size = New Size(200, 20)
        statsPanel.Controls.Add(LblActiveMonths)

        ' მეორე რიგი - შემოსავლები
        LblTotalRevenue = New Label()
        LblTotalRevenue.Text = "მთლიანი შემოსავალი: 0.00 ₾"
        LblTotalRevenue.Location = New Point(10, 40)
        LblTotalRevenue.Size = New Size(200, 20)
        LblTotalRevenue.Font = New Font("Sylfaen", 10, FontStyle.Bold)
        LblTotalRevenue.ForeColor = Color.DarkGreen
        statsPanel.Controls.Add(LblTotalRevenue)

        LblAveragePrice = New Label()
        LblAveragePrice.Text = "საშუალო ფასი: 0.00 ₾"
        LblAveragePrice.Location = New Point(220, 40)
        LblAveragePrice.Size = New Size(200, 20)
        LblAveragePrice.Font = New Font("Sylfaen", 10, FontStyle.Bold)
        LblAveragePrice.ForeColor = Color.DarkGreen
        statsPanel.Controls.Add(LblAveragePrice)

        LblBeneficiariesCount = New Label()
        LblBeneficiariesCount.Text = "უნიკალური ბენეფიციარები: 0"
        LblBeneficiariesCount.Location = New Point(430, 40)
        LblBeneficiariesCount.Size = New Size(220, 20)
        statsPanel.Controls.Add(LblBeneficiariesCount)

        LblTherapyTypes = New Label()
        LblTherapyTypes.Text = "თერაპიის ტიპები: 0"
        LblTherapyTypes.Location = New Point(660, 40)
        LblTherapyTypes.Size = New Size(200, 20)
        statsPanel.Controls.Add(LblTherapyTypes)

        ' მესამე რიგი
        LblMostCommonTherapy = New Label()
        LblMostCommonTherapy.Text = "ძირითადი თერაპია: -"
        LblMostCommonTherapy.Location = New Point(10, 70)
        LblMostCommonTherapy.Size = New Size(250, 20)
        statsPanel.Controls.Add(LblMostCommonTherapy)

        LblFirstSession = New Label()
        LblFirstSession.Text = "პირველი სესია: -"
        LblFirstSession.Location = New Point(270, 70)
        LblFirstSession.Size = New Size(200, 20)
        statsPanel.Controls.Add(LblFirstSession)

        LblLastSession = New Label()
        LblLastSession.Text = "ბოლო სესია: -"
        LblLastSession.Location = New Point(480, 70)
        LblLastSession.Size = New Size(200, 20)
        statsPanel.Controls.Add(LblLastSession)
    End Sub

    ''' <summary>
    ''' სესიების DataGridView-ის შექმნა
    ''' </summary>
    Private Sub CreateSessionsGrid()
        DgvSessions = New DataGridView()
        DgvSessions.Location = New Point(20, 280)
        DgvSessions.Size = New Size(1050, 380)
        DgvSessions.AutoGenerateColumns = False
        DgvSessions.AllowUserToAddRows = False
        DgvSessions.AllowUserToDeleteRows = False
        DgvSessions.ReadOnly = True
        DgvSessions.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DgvSessions.MultiSelect = False
        DgvSessions.RowHeadersVisible = False
        DgvSessions.BackgroundColor = Color.White
        DgvSessions.GridColor = Color.LightGray
        DgvSessions.DefaultCellStyle.Font = New Font("Sylfaen", 9)
        DgvSessions.ColumnHeadersDefaultCellStyle.Font = New Font("Sylfaen", 10, FontStyle.Bold)
        DgvSessions.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(220, 230, 250)
        DgvSessions.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 245, 250)
        Me.Controls.Add(DgvSessions)

        ' სვეტების შექმნა
        CreateDataGridColumns()
    End Sub

    ''' <summary>
    ''' ღილაკების სექციის შექმნა
    ''' </summary>
    Private Sub CreateButtonsSection()
        Dim yPos As Integer = 670

        BtnExportToExcel = New Button()
        BtnExportToExcel.Text = "Excel-ში ექსპორტი"
        BtnExportToExcel.Location = New Point(20, yPos)
        BtnExportToExcel.Size = New Size(150, 35)
        BtnExportToExcel.Font = New Font("Sylfaen", 9)
        BtnExportToExcel.UseVisualStyleBackColor = True
        BtnExportToExcel.Enabled = False
        Me.Controls.Add(BtnExportToExcel)

        BtnPrint = New Button()
        BtnPrint.Text = "ბეჭდვა"
        BtnPrint.Location = New Point(180, yPos)
        BtnPrint.Size = New Size(100, 35)
        BtnPrint.Font = New Font("Sylfaen", 9)
        BtnPrint.UseVisualStyleBackColor = True
        BtnPrint.Enabled = False
        Me.Controls.Add(BtnPrint)

        BtnGenerateReport = New Button()
        BtnGenerateReport.Text = "დეტალური რეპორტი"
        BtnGenerateReport.Location = New Point(290, yPos)
        BtnGenerateReport.Size = New Size(150, 35)
        BtnGenerateReport.Font = New Font("Sylfaen", 9)
        BtnGenerateReport.UseVisualStyleBackColor = True
        BtnGenerateReport.Enabled = False
        Me.Controls.Add(BtnGenerateReport)

        BtnMonthlyAnalysis = New Button()
        BtnMonthlyAnalysis.Text = "თვიური ანალიზი"
        BtnMonthlyAnalysis.Location = New Point(450, yPos)
        BtnMonthlyAnalysis.Size = New Size(130, 35)
        BtnMonthlyAnalysis.Font = New Font("Sylfaen", 9)
        BtnMonthlyAnalysis.UseVisualStyleBackColor = True
        BtnMonthlyAnalysis.Enabled = False
        Me.Controls.Add(BtnMonthlyAnalysis)
    End Sub

    ''' <summary>
    ''' DataGridView-ის სვეტების შექმნა
    ''' </summary>
    Private Sub CreateDataGridColumns()
        Try
            DgvSessions.Columns.Clear()

            ' ID სვეტი
            Dim colId As New DataGridViewTextBoxColumn With {
                .Name = "Id",
                .HeaderText = "ID",
                .DataPropertyName = "Id",
                .Width = 50,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
            }
            DgvSessions.Columns.Add(colId)

            ' თარიღი სვეტი
            Dim colDate As New DataGridViewTextBoxColumn With {
                .Name = "DateTime",
                .HeaderText = "თარიღი",
                .DataPropertyName = "FormattedDateTime",
                .Width = 120
            }
            DgvSessions.Columns.Add(colDate)

            ' ბენეფიციარი სვეტი
            Dim colBeneficiary As New DataGridViewTextBoxColumn With {
                .Name = "Beneficiary",
                .HeaderText = "ბენეფიციარი",
                .DataPropertyName = "FullName",
                .Width = 180
            }
            DgvSessions.Columns.Add(colBeneficiary)

            ' თერაპიის ტიპი სვეტი
            Dim colTherapyType As New DataGridViewTextBoxColumn With {
                .Name = "TherapyType",
                .HeaderText = "თერაპიის ტიპი",
                .DataPropertyName = "TherapyType",
                .Width = 150
            }
            DgvSessions.Columns.Add(colTherapyType)

            ' ხანგრძლივობა სვეტი
            Dim colDuration As New DataGridViewTextBoxColumn With {
                .Name = "Duration",
                .HeaderText = "ხანგრძლივობა (წთ)",
                .DataPropertyName = "Duration",
                .Width = 90,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
            }
            DgvSessions.Columns.Add(colDuration)

            ' ფასი სვეტი
            Dim colPrice As New DataGridViewTextBoxColumn With {
                .Name = "Price",
                .HeaderText = "ფასი (₾)",
                .DataPropertyName = "Price",
                .Width = 80,
                .DefaultCellStyle = New DataGridViewCellStyle With {
                    .Alignment = DataGridViewContentAlignment.MiddleRight,
                    .Format = "N2"
                }
            }
            DgvSessions.Columns.Add(colPrice)

            ' სტატუსი სვეტი
            Dim colStatus As New DataGridViewTextBoxColumn With {
                .Name = "Status",
                .HeaderText = "სტატუსი",
                .DataPropertyName = "Status",
                .Width = 90,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
            }
            DgvSessions.Columns.Add(colStatus)

            ' დაფინანსება სვეტი
            Dim colFunding As New DataGridViewTextBoxColumn With {
                .Name = "Funding",
                .HeaderText = "დაფინანსება",
                .DataPropertyName = "Funding",
                .Width = 90
            }
            DgvSessions.Columns.Add(colFunding)

            ' კომენტარები სვეტი
            Dim colComments As New DataGridViewTextBoxColumn With {
                .Name = "Comments",
                .HeaderText = "კომენტარები",
                .DataPropertyName = "Comments",
                .Width = 150,
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            }
            DgvSessions.Columns.Add(colComments)

            Debug.WriteLine("UC_TherapistReport: DataGridView სვეტები შეიქმნა")

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport CreateDataGridColumns: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    ''' <param name="service">IDataService ობიექტი</param>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_TherapistReport: მონაცემთა სერვისი მითითებულია")
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის მითითება
    ''' </summary>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub SetUserInfo(email As String, role As String)
        userEmail = email
        userRole = role
        Debug.WriteLine($"UC_TherapistReport: მომხმარებლის ინფორმაცია - Email: {email}, Role: {role}")
    End Sub

    ''' <summary>
    ''' ძიების ღილაკზე დაჭერა
    ''' </summary>
    Private Async Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Try
            ' ძიების ტექსტის შემოწმება
            Dim searchText As String = TxtSearch.Text.Trim()

            If String.IsNullOrEmpty(searchText) Then
                MessageBox.Show("გთხოვთ შეიყვანოთ თერაპევტის სახელი ძიებისთვის",
                              "ძიება", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TxtSearch.Focus()
                Return
            End If

            ' ძიების ინდიკატორის ჩართვა
            BtnSearch.Enabled = False
            BtnSearch.Text = "იძებნება..."

            ' ძიების შესრულება ასინქრონულად
            Await Task.Run(Sub() SearchTherapist(searchText))

        Catch ex As Exception
            Debug.WriteLine($"UC_TherapistReport BtnSearch_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ძიების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ღილაკის აღდგენა
            BtnSearch.Enabled = True
            BtnSearch.Text = "ძიება"
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის ძიება
    ''' </summary>
    ''' <param name="searchText">ძიების ტექსტი</param>
    Private Sub SearchTherapist(searchText As String)
        Try
            Debug.WriteLine($"SearchTherapist: იწყება ძიება '{searchText}'-სთვის")

            If dataService Is Nothing Then
                Me.Invoke(Sub() MessageBox.Show("მონაცემთა სერვისი არ არის ხელმისაწვდომი", "შეცდომა",
                                              MessageBoxButtons.OK, MessageBoxIcon.Error))
                Return
            End If

            ' ყველა სესიის წამოღება
            Dim allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"SearchTherapist: წამოღებულია {allSessions.Count} სესია")

            ' თერაპევტის სესიების ფილტრაცია
            Dim matchingSessions = allSessions.Where(Function(s)
                                                         Return s.TherapistName.ToLower().Contains(searchText.ToLower())
                                                     End Function).ToList()

            Debug.WriteLine($"SearchTherapist: ნაპოვნია {matchingSessions.Count} შესაბამისი სესია")

            ' UI-ის განახლება მთავარ თრედზე
            Me.Invoke(Sub()
                          Try
                              If matchingSessions.Count = 0 Then
                                  MessageBox.Show($"თერაპევტი '{searchText}' ვერ მოიძებნა", "ძიების შედეგი",
                                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                                  ResetForm()
                                  Return
                              End If

                              ' თერაპევტის ინფორმაციის შექმნა
                              CreateTherapistInfo(matchingSessions)

                              ' სესიების ჩვენება
                              LoadTherapistSessions(matchingSessions)

                              ' სტატისტიკის განახლება
                              UpdateStatistics()

                              ' თვიური ანალიზის შექმნა
                              CreateMonthlyAnalysis(matchingSessions)

                              Debug.WriteLine($"SearchTherapist: თერაპევტის ანგარიში წარმატებით ჩაიტვირთა")

                          Catch uiEx As Exception
                              Debug.WriteLine($"SearchTherapist UI განახლება: შეცდომა - {uiEx.Message}")
                              MessageBox.Show($"UI განახლების შეცდომა: {uiEx.Message}", "შეცდომა",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                          End Try
                      End Sub)

        Catch ex As Exception
            Debug.WriteLine($"SearchTherapist: შეცდომა - {ex.Message}")
            Me.Invoke(Sub() MessageBox.Show($"ძიების შეცდომა: {ex.Message}", "შეცდომა",
                                          MessageBoxButtons.OK, MessageBoxIcon.Error))
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის ინფორმაციის შექმნა სესიების საფუძველზე
    ''' </summary>
    ''' <param name="sessions">თერაპევტის სესიები</param>
    Private Sub CreateTherapistInfo(sessions As List(Of SessionModel))
        Try
            If sessions.Count = 0 Then Return

            ' პირველი სესიიდან თერაპევტის სახელის აღება
            Dim therapistName = sessions.First().TherapistName

            currentTherapist = New TherapistInfo()
            currentTherapist.Name = therapistName

            ' ძირითადი სტატისტიკა
            currentTherapist.TotalSessions = sessions.Count
            currentTherapist.CompletedSessions = sessions.Where(Function(s) s.Status = "შესრულებული").Count()
            currentTherapist.CancelledSessions = sessions.Where(Function(s) s.Status = "გაუქმებული").Count()
            currentTherapist.PendingSessions = sessions.Where(Function(s) s.Status = "დაგეგმილი").Count()

            ' შემოსავლების გამოთვლა (მხოლოდ შესრულებული სესიებიდან)
            Dim completedSessions = sessions.Where(Function(s) s.Status = "შესრულებული").ToList()
            currentTherapist.TotalRevenue = completedSessions.Sum(Function(s) s.Price)

            If completedSessions.Count > 0 Then
                currentTherapist.AverageSessionPrice = currentTherapist.TotalRevenue / completedSessions.Count
            Else
                currentTherapist.AverageSessionPrice = 0
            End If

            ' თარიღების დადგენა
            If sessions.Count > 0 Then
                currentTherapist.FirstSessionDate = sessions.Min(Function(s) s.DateTime)
                currentTherapist.LastSessionDate = sessions.Max(Function(s) s.DateTime)
            End If

            ' უნიკალური ბენეფიციარების რაოდენობა
            currentTherapist.BeneficiariesCount = sessions.GroupBy(Function(s) s.FullName).Count()

            ' თერაპიის ტიპების ანალიზი
            currentTherapist.TherapyTypes = sessions.GroupBy(Function(s) s.TherapyType) _
                                                  .Select(Function(g) g.Key) _
                                                  .Where(Function(t) Not String.IsNullOrEmpty(t)) _
                                                  .ToList()

            ' ყველაზე ხშირი თერაპიის ტიპი
            If sessions.Count > 0 Then
                Dim mostCommon = sessions.Where(Function(s) Not String.IsNullOrEmpty(s.TherapyType)) _
                                       .GroupBy(Function(s) s.TherapyType) _
                                       .OrderByDescending(Function(g) g.Count()) _
                                       .FirstOrDefault()

                currentTherapist.MostCommonTherapyType = If(mostCommon?.Key, "მითითებული არ არის")
            Else
                currentTherapist.MostCommonTherapyType = "მითითებული არ არის"
            End If

            ' აქტიური თვეების რაოდენობა
            If sessions.Count > 0 Then
                currentTherapist.ActiveMonths = sessions.GroupBy(Function(s) New With {s.DateTime.Year, s.DateTime.Month}).Count()
            Else
                currentTherapist.ActiveMonths = 0
            End If

            Debug.WriteLine($"CreateTherapistInfo: შეიქმნა თერაპევტის ინფორმაცია - {currentTherapist.Name}")

        Catch ex As Exception
            Debug.WriteLine($"CreateTherapistInfo: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის სესიების ჩატვირთვა DataGridView-ში
    ''' </summary>
    ''' <param name="sessions">სესიების სია</param>
    Private Sub LoadTherapistSessions(sessions As List(Of SessionModel))
        Try
            ' სესიების დალაგება თარიღის მიხედვით (ახალიდან ძველისკენ)
            therapistSessions = sessions.OrderByDescending(Function(s) s.DateTime).ToList()

            ' DataGridView-ის განახლება
            DgvSessions.DataSource = Nothing
            DgvSessions.DataSource = therapistSessions

            ' მწკრივების ფერების განახლება არსებული SessionStatusColors კლასის გამოყენებით
            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim session = TryCast(row.DataBoundItem, SessionModel)
                    If session IsNot Nothing Then
                        ' გამოვიყენოთ არსებული SessionStatusColors კლასი
                        Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.BackColor = statusColor

                        ' ასევე ჩარჩოს ფერიც დავაყენოთ
                        Dim borderColor = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.SelectionBackColor = borderColor
                    End If
                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"LoadTherapistSessions: ჩაიტვირთა {therapistSessions.Count} სესია")

        Catch ex As Exception
            Debug.WriteLine($"LoadTherapistSessions: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სტატისტიკის განახლება
    ''' </summary>
    Private Sub UpdateStatistics()
        Try
            If currentTherapist Is Nothing Then Return

            ' თერაპევტის ინფორმაცია
            LblTherapistName.Text = $"თერაპევტი: {currentTherapist.Name}"

            ' ძირითადი სტატისტიკა
            LblTotalSessions.Text = $"სულ სესიები: {currentTherapist.TotalSessions}"
            LblCompletedSessions.Text = $"შესრულებული: {currentTherapist.CompletedSessions}"
            LblCancelledSessions.Text = $"გაუქმებული: {currentTherapist.CancelledSessions}"
            LblPendingSessions.Text = $"დაგეგმილი: {currentTherapist.PendingSessions}"

            ' შემოსავლები
            LblTotalRevenue.Text = $"მთლიანი შემოსავალი: {currentTherapist.TotalRevenue:N2} ₾"
            LblAveragePrice.Text = $"საშუალო ფასი: {currentTherapist.AverageSessionPrice:N2} ₾"

            ' კლიენტები და თერაპიები
            LblBeneficiariesCount.Text = $"უნიკალური ბენეფიციარები: {currentTherapist.BeneficiariesCount}"
            LblTherapyTypes.Text = $"თერაპიის ტიპები: {currentTherapist.TherapyTypes.Count}"
            LblMostCommonTherapy.Text = $"ძირითადი თერაპია: {currentTherapist.MostCommonTherapyType}"

            ' აქტივობა
            LblActiveMonths.Text = $"აქტიური თვეები: {currentTherapist.ActiveMonths}"

            ' თარიღები
            If currentTherapist.FirstSessionDate.HasValue Then
                LblFirstSession.Text = $"პირველი სესია: {currentTherapist.FirstSessionDate.Value:dd.MM.yyyy}"
            Else
                LblFirstSession.Text = "პირველი სესია: -"
            End If

            If currentTherapist.LastSessionDate.HasValue Then
                LblLastSession.Text = $"ბოლო სესია: {currentTherapist.LastSessionDate.Value:dd.MM.yyyy}"
            Else
                LblLastSession.Text = "ბოლო სესია: -"
            End If

            ' ღილაკების ააქტიურება
            BtnExportToExcel.Enabled = True
            BtnPrint.Enabled = True
            BtnGenerateReport.Enabled = True
            BtnMonthlyAnalysis.Enabled = True

            Debug.WriteLine("UpdateStatistics: სტატისტიკა განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"UpdateStatistics: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თვიური ანალიზის შექმნა
    ''' </summary>
    ''' <param name="sessions">სესიების სია</param>
    Private Sub CreateMonthlyAnalysis(sessions As List(Of SessionModel))
        Try
            monthlyStats.Clear()

            ' სესიების ჯგუფირება თვეების მიხედვით
            Dim monthlyGroups = sessions.GroupBy(Function(s) New With {s.DateTime.Year, s.DateTime.Month}) _
                                      .OrderBy(Function(g) g.Key.Year) _
                                      .ThenBy(Function(g) g.Key.Month)

            For Each group In monthlyGroups
                Dim key = $"{group.Key.Year}-{group.Key.Month:D2}"
                Dim sessionsInMonth = group.ToList()

                Dim stats As New MonthlyStatistics()
                stats.Year = group.Key.Year
                stats.Month = group.Key.Month
                stats.SessionsCount = sessionsInMonth.Count
                stats.CompletedSessions = sessionsInMonth.Where(Function(s) s.Status = "შესრულებული").Count()
                stats.CancelledSessions = sessionsInMonth.Where(Function(s) s.Status = "გაუქმებული").Count()
                stats.Revenue = sessionsInMonth.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)
                stats.UniqueClients = sessionsInMonth.GroupBy(Function(s) s.FullName).Count()

                monthlyStats(key) = stats
            Next

            Debug.WriteLine($"CreateMonthlyAnalysis: შეიქმნა {monthlyStats.Count} თვის ანალიზი")

        Catch ex As Exception
            Debug.WriteLine($"CreateMonthlyAnalysis: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფორმის გასუფთავება
    ''' </summary>
    Private Sub ResetForm()
        Try
            ' თერაპევტის ინფორმაციის გასუფთავება
            currentTherapist = Nothing
            therapistSessions.Clear()
            monthlyStats.Clear()

            ' UI ელემენტების გასუფთავება
            LblTherapistName.Text = "თერაპევტი: -"
            LblTotalSessions.Text = "სულ სესიები: 0"
            LblCompletedSessions.Text = "შესრულებული: 0"
            LblCancelledSessions.Text = "გაუქმებული: 0"
            LblPendingSessions.Text = "დაგეგმილი: 0"
            LblTotalRevenue.Text = "მთლიანი შემოსავალი: 0.00 ₾"
            LblAveragePrice.Text = "საშუალო ფასი: 0.00 ₾"
            LblBeneficiariesCount.Text = "უნიკალური ბენეფიციარები: 0"
            LblTherapyTypes.Text = "თერაპიის ტიპები: 0"
            LblMostCommonTherapy.Text = "ძირითადი თერაპია: -"
            LblActiveMonths.Text = "აქტიური თვეები: 0"
            LblFirstSession.Text = "პირველი სესია: -"
            LblLastSession.Text = "ბოლო სესია: -"

            ' DataGridView-ის გასუფთავება
            DgvSessions.DataSource = Nothing

            ' ღილაკების გათიშვა
            BtnExportToExcel.Enabled = False
            BtnPrint.Enabled = False
            BtnGenerateReport.Enabled = False
            BtnMonthlyAnalysis.Enabled = False

            Debug.WriteLine("ResetForm: ფორმა გაიწმინდა")

        Catch ex As Exception
            Debug.WriteLine($"ResetForm: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ' ღილაკების ივენთ ჰენდლერები მარტივი იმპლემენტაციით
    Private Sub BtnExportToExcel_Click(sender As Object, e As EventArgs) Handles BtnExportToExcel.Click
        MessageBox.Show("Excel ექსპორტი განხორციელდება მალე", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        If printService IsNot Nothing Then printService.Dispose()
        printService = New AdvancedDataGridViewPrintService(DgvSessions)
        printService.ShowFullPrintDialog()
    End Sub

    Private Sub BtnGenerateReport_Click(sender As Object, e As EventArgs) Handles BtnGenerateReport.Click
        MessageBox.Show("დეტალური რეპორტი განხორციელდება მალე", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnMonthlyAnalysis_Click(sender As Object, e As EventArgs) Handles BtnMonthlyAnalysis.Click
        MessageBox.Show("თვიური ანალიზი განხორციელდება მალე", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Enter ღილაკზე დაჭერისას ძიების ჩართვა
    ''' </summary>
    Private Sub TxtSearch_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtSearch.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            BtnSearch.PerformClick()
        End If
    End Sub

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            printService?.Dispose()
            therapistSessions?.Clear()
            monthlyStats?.Clear()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class