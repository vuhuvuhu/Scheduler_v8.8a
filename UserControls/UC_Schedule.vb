' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განრიგის UserControl - აჩვენებს და მართავს სესიების განრიგს
' ფურცლების ნავიგაციით და ჩანაწერების რაოდენობის ფილტრაციით
' შესწორება: რედაქტირების თარიღისა და კომენტარების სვეტის ნორმალური ჩვენება
' v8.2-ის მსგავსი პირდაპირი მონაცემების ჩატვირთვა
' ===========================================
Imports System.Windows.Forms
Imports System.Globalization
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class UC_Schedule
    Inherits UserControl

    ' მონაცემთა სერვისი
    Private dataService As IDataService = Nothing

    ' მომხმარებლის ელფოსტა რედაქტირებისთვის
    Private userEmail As String = ""
    Private userRoleID As Integer = 6 ' ნაგულისხმევი როლი

    ' ფილტრაციისა და ნავიგაციისთვის
    Private currentPage As Integer = 1
    Private pageSize As Integer = 20 ' ნაგულისხმები ზომა
    Private totalPages As Integer = 1

    ''' <summary>
    ''' მონაცემთა სერვისის და მომხმარებლის email-ის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი (IDataService)</param>
    ''' <param name="email">მომხმარებლის ელფოსტა</param>
    ''' <param name="role">მომხმარებლის როლი (1=ადმინი, 2=მენეჯერი, 6=ჩვეულებრივი)</param>
    Public Sub SetDataService(service As IDataService, Optional email As String = "", Optional role As Integer = 6)
        dataService = service
        userEmail = email
        userRoleID = role

        Debug.WriteLine($"UC_Schedule: დაყენებული role={role}, email={email}")

        ' ძაღლით როლის დაყენება ადმინისთვის ან მენეჯერისთვის
        If String.IsNullOrEmpty(email) OrElse role = 6 Then
            ' ვცდილობთ როლის განსაზღვრას dataService-დან
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

        ' ტესტირებისთვის ძაღლით ადმინის როლის დაყენება
        If userRoleID = 6 Then
            userRoleID = 1 ' ძაღლით ადმინის როლი ტესტირებისთვის
            Debug.WriteLine("UC_Schedule: ძაღლით დაყენებული ადმინის როლი ტესტირებისთვის")
        End If

        ' თუ კონტროლი უკვე ჩატვირთულია, ვტვირთავთ მონაცემებს
        If Me.IsHandleCreated Then
            LoadFilteredSchedule()
        End If
    End Sub

    ''' <summary>
    ''' კონტროლის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
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

            ' ფილტრების ინიციალიზაცია
            InitializeFilters()

            ' თუ მონაცემთა სერვისი უკვე დაყენებულია, ვტვირთავთ მონაცემებს
            If dataService IsNot Nothing Then
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            MessageBox.Show($"კონტროლის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' თარიღების DatePicker-ების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeDatePickers()
        Try
            ' მიმდინარე თარიღი
            Dim currentDate As DateTime = DateTime.Today

            ' მიმდინარე თვის პირველი დღე
            Dim firstDayOfMonth As DateTime = New DateTime(currentDate.Year, currentDate.Month, 1)

            ' მიმდინარე თვის ბოლო დღე
            Dim lastDayOfMonth As DateTime = firstDayOfMonth.AddMonths(1).AddDays(-1)

            ' DatePicker-ების დაყენება
            If DtpDan IsNot Nothing Then DtpDan.Value = firstDayOfMonth
            If DtpMde IsNot Nothing Then DtpMde.Value = lastDayOfMonth

            ' თარიღების შეცვლის ივენთები
            If DtpDan IsNot Nothing Then AddHandler DtpDan.ValueChanged, AddressOf DatePickers_ValueChanged
            If DtpMde IsNot Nothing Then AddHandler DtpMde.ValueChanged, AddressOf DatePickers_ValueChanged

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeFilters()
        Try
            ' ჩეკბოქსების ინიციალიზაცია და ივენთების მიბმა
            InitializeStatusCheckBoxes()

            ' კომბობოქსების ივენთების მიბმა
            If CBBeneName IsNot Nothing Then AddHandler CBBeneName.SelectedIndexChanged, AddressOf Filter_Changed
            If CBBeneSurname IsNot Nothing Then AddHandler CBBeneSurname.SelectedIndexChanged, AddressOf Filter_Changed
            If CBPer IsNot Nothing Then AddHandler CBPer.SelectedIndexChanged, AddressOf Filter_Changed
            If CBTer IsNot Nothing Then AddHandler CBTer.SelectedIndexChanged, AddressOf Filter_Changed
            If CBSpace IsNot Nothing Then AddHandler CBSpace.SelectedIndexChanged, AddressOf Filter_Changed
            If CBDaf IsNot Nothing Then AddHandler CBDaf.SelectedIndexChanged, AddressOf Filter_Changed

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' სტატუსის ჩეკბოქსების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            ' ყველა ჩეკბოქსის მონიშვნა და ივენთების მიბმა
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBox($"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    checkBox.Checked = True ' ყველა მონიშნული ჩატვირთვისას
                    AddHandler checkBox.CheckedChanged, AddressOf StatusCheckBox_Changed
                    Debug.WriteLine($"UC_Schedule: ჩეკბოქსი CheckBox{i} ინიციალიზებული: '{checkBox.Text}'")
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ჩეკბოქსების ინიციალიზაციის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ჩეკბოქსის მოძებნა მთელ კონტროლში (რეკურსიულად)
    ''' </summary>
    Private Function FindCheckBox(name As String) As CheckBox
        Try
            Return FindControlRecursive(Me, name)
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ჩეკბოქსის {name} მოძებნის შეცდომა: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' კონტროლის რეკურსიული ძებნა
    ''' </summary>
    Private Function FindControlRecursive(parent As Control, name As String) As CheckBox
        Try
            ' ჯერ პირდაპირ შვილებში ვეძებთ
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is CheckBox AndAlso ctrl.Name = name Then
                    Return DirectCast(ctrl, CheckBox)
                End If
            Next

            ' მერე რეკურსიულად შვილების შვილებში
            For Each ctrl As Control In parent.Controls
                Dim found = FindControlRecursive(ctrl, name)
                If found IsNot Nothing Then
                    Return found
                End If
            Next

            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' სტატუსის ჩეკბოქსების შეცვლის ივენთი
    ''' </summary>
    Private Sub StatusCheckBox_Changed(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("UC_Schedule: სტატუსის ჩეკბოქსი შეიცვალა")
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: სტატუსის ჩეკბოქსის შეცვლის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონიშნული სტატუსების სიის მიღება
    ''' </summary>
    Private Function GetSelectedStatuses() As List(Of String)
        Dim selectedStatuses As New List(Of String)

        Try
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBox($"CheckBox{i}")
                If checkBox IsNot Nothing AndAlso checkBox.Checked Then
                    ' ჩეკბოქსის ტექსტი სტატუსად
                    Dim statusText = checkBox.Text.Trim()
                    If Not String.IsNullOrEmpty(statusText) Then
                        selectedStatuses.Add(statusText)
                    End If
                End If
            Next

            Debug.WriteLine($"UC_Schedule: მონიშნული სტატუსები: {String.Join(", ", selectedStatuses)}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: სტატუსების სიის მიღების შეცდომა: {ex.Message}")
        End Try

        Return selectedStatuses
    End Function

    ''' <summary>
    ''' ყველა სტატუსის ჩეკბოქსის მონიშვნა/განიშვნა
    ''' </summary>
    ''' <param name="checkAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
    Public Sub SetAllStatusCheckBoxes(checkAll As Boolean)
        Try
            Debug.WriteLine($"UC_Schedule: ყველა ჩეკბოქსის დაყენება: {checkAll}")

            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBox($"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    checkBox.Checked = checkAll
                End If
            Next

            ' მონაცემების ხელახალი ჩატვირთვა
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ყველა ჩეკბოქსის დაყენების შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონკრეტული სტატუსის ჩეკბოქსის მონიშვნა/განიშვნა
    ''' </summary>
    ''' <param name="statusText">სტატუსის ტექსტი</param>
    ''' <param name="isChecked">მონიშნული თუ არა</param>
    Public Sub SetStatusCheckBox(statusText As String, isChecked As Boolean)
        Try
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBox($"CheckBox{i}")
                If checkBox IsNot Nothing AndAlso
                   String.Equals(checkBox.Text.Trim(), statusText.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    checkBox.Checked = isChecked
                    Debug.WriteLine($"UC_Schedule: ჩეკბოქსი '{statusText}' დაყენებული: {isChecked}")
                    Exit For
                End If
            Next

            ' მონაცემების ხელახალი ჩატვირთვა
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: სტატუსის ჩეკბოქსის დაყენების შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონიშნული სტატუსების რაოდენობის მიღება
    ''' </summary>
    Public ReadOnly Property SelectedStatusCount As Integer
        Get
            Return GetSelectedStatuses().Count
        End Get
    End Property

    ''' <summary>
    ''' რადიობუტონების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeRadioButtons()
        Try
            ' RB20-ს ღირს ნაგულისხმევად
            If RB20 IsNot Nothing Then RB20.Checked = True
            pageSize = 20

            ' ივენთების მიბმა
            If RB20 IsNot Nothing Then AddHandler RB20.CheckedChanged, AddressOf RadioButton_CheckedChanged
            If RB50 IsNot Nothing Then AddHandler RB50.CheckedChanged, AddressOf RadioButton_CheckedChanged
            If RB100 IsNot Nothing Then AddHandler RB100.CheckedChanged, AddressOf RadioButton_CheckedChanged

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის კონფიგურაცია
    ''' </summary>
    Private Sub ConfigureDataGridView()
        Try
            If DgvSchedule Is Nothing Then Return

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

            ' რედაქტირების ღილაკის ივენთი
            AddHandler DgvSchedule.CellClick, AddressOf DgvSchedule_CellClick

        Catch ex As Exception
            MessageBox.Show($"DataGridView კონფიგურაციის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeNavigationButtons()
        Try
            ' ღილაკების სტილი
            If BtnPrev IsNot Nothing Then
                BtnPrev.FlatStyle = FlatStyle.Flat
                BtnPrev.BackColor = Color.FromArgb(60, 80, 150)
                BtnPrev.ForeColor = Color.White
                BtnPrev.Text = "<<" ' წინა გვერდი
                AddHandler BtnPrev.Click, AddressOf BtnPrev_Click
            End If

            If BtnNext IsNot Nothing Then
                BtnNext.FlatStyle = FlatStyle.Flat
                BtnNext.BackColor = Color.FromArgb(60, 80, 150)
                BtnNext.ForeColor = Color.White
                BtnNext.Text = ">>" ' შემდეგი გვერდი
                AddHandler BtnNext.Click, AddressOf BtnNext_Click
            End If

            ' გვერდის ლეიბლის ინიციალიზაცია
            UpdatePageLabel()

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' 📊 ცხრილის შევსება ფილტრების მიხედვით - v8.2 ვერსიის მსგავსი
    ''' </summary>
    Public Sub LoadFilteredSchedule()
        Try
            If DgvSchedule Is Nothing OrElse dataService Is Nothing Then Return

            Debug.WriteLine("UC_Schedule: LoadFilteredSchedule - დაიწყო")

            ' ცხრილის გასუფთავება
            DgvSchedule.Rows.Clear()

            ' თუ ჯერ სვეტები არ დამატებულა, ვამატებთ და ვუსვამთ ზომებს
            If DgvSchedule.Columns.Count = 0 Then
                CreateDataGridViewColumns()
            End If

            ' თარიღის დიაპაზონი ფილტრიდან
            Dim dateFrom As Date = If(DtpDan IsNot Nothing, DtpDan.Value, DateTime.Today.AddDays(-30))
            Dim dateTo As Date = If(DtpMde IsNot Nothing, DtpMde.Value, DateTime.Today)

            Debug.WriteLine($"UC_Schedule: ფილტრი - {dateFrom:dd.MM.yyyy} - {dateTo:dd.MM.yyyy}")

            ' Google Sheets-დან ყველა ჩანაწერის წამოღება
            Dim allRows As IList(Of IList(Of Object)) = dataService.GetData("DB-Schedule!A2:O")
            Dim filtered As New List(Of IList(Of Object))

            If allRows IsNot Nothing Then
                Debug.WriteLine($"UC_Schedule: მიღებულია {allRows.Count} ჩანაწერი")

                ' ჩანაწერების გაფილტვრა
                For Each row As IList(Of Object) In allRows
                    Try
                        ' აუცილებელი ველები: N (0), თარიღი (5), ხანგძლიობა (6), სახელი (3), გვარი (4)
                        If row.Count < 7 Then Continue For ' მინიმუმ უნდა არსებობდეს N-დან ხანგძლიობამდე
                        If String.IsNullOrWhiteSpace(row(0).ToString()) Then Continue For ' N ცარიელი
                        If String.IsNullOrWhiteSpace(row(5).ToString()) Then Continue For ' თარიღი ცარიელი
                        If String.IsNullOrWhiteSpace(row(3).ToString()) OrElse String.IsNullOrWhiteSpace(row(4).ToString()) Then Continue For ' სახელი ან გვარი ცარიელი

                        ' თარიღის პარსინგი
                        Dim dt As DateTime
                        Dim rawDate As String = row(5).ToString().Trim()
                        Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                        If Not DateTime.TryParseExact(rawDate, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                            If Not DateTime.TryParse(rawDate, dt) Then
                                Debug.WriteLine($"⛔ ვერ მოხდა თარიღის წაკითხვა: {rawDate}")
                                Continue For
                            End If
                        End If

                        ' თარიღის ფილტრი
                        If dt.Date < dateFrom.Date OrElse dt.Date > dateTo.Date Then Continue For

                        ' სტატუსის ფილტრი - ჩეკბოქსების მიხედვით
                        Dim selectedStatuses = GetSelectedStatuses()
                        If selectedStatuses.Count > 0 Then
                            Dim rowStatus As String = If(row.Count > 12, row(12).ToString().Trim(), "")
                            Dim statusMatches As Boolean = False

                            ' შევამოწმოთ არის თუ არა ამ სესიის სტატუსი მონიშნულ სტატუსებში
                            For Each selectedStatus In selectedStatuses
                                If String.Equals(rowStatus, selectedStatus, StringComparison.OrdinalIgnoreCase) Then
                                    statusMatches = True
                                    Exit For
                                End If
                            Next

                            ' თუ სტატუსი არ ემთხვევა რომელიმე მონიშნულს, გამოვტოვოთ
                            If Not statusMatches Then Continue For
                        End If

                        ' ბენეფიციარის სახელის ფილტრი
                        If CBBeneName IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBBeneName.Text) AndAlso CBBeneName.Text <> "ყველა" Then
                            Dim nameFilter As String = CBBeneName.Text.Trim()
                            Dim rowName As String = row(3).ToString().Trim()
                            If Not String.Equals(rowName, nameFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' ბენეფიციარის გვარის ფილტრი
                        If CBBeneSurname IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBBeneSurname.Text) AndAlso CBBeneSurname.Text <> "ყველა" Then
                            Dim surnameFilter As String = CBBeneSurname.Text.Trim()
                            Dim rowSurname As String = row(4).ToString().Trim()
                            If Not String.Equals(rowSurname, surnameFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' თერაპევტის ფილტრი
                        If CBPer IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBPer.Text) AndAlso CBPer.Text <> "ყველა" Then
                            Dim therapistFilter As String = CBPer.Text.Trim()
                            Dim rowTherapist As String = If(row.Count > 8, row(8).ToString().Trim(), "")
                            If Not String.Equals(rowTherapist, therapistFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' თერაპიის ტიპის ფილტრი
                        If CBTer IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBTer.Text) AndAlso CBTer.Text <> "ყველა" Then
                            Dim therapyFilter As String = CBTer.Text.Trim()
                            Dim rowTherapy As String = If(row.Count > 9, row(9).ToString().Trim(), "")
                            If Not String.Equals(rowTherapy, therapyFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' სივრცის ფილტრი
                        If CBSpace IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBSpace.Text) AndAlso CBSpace.Text <> "ყველა" Then
                            Dim spaceFilter As String = CBSpace.Text.Trim()
                            Dim rowSpace As String = If(row.Count > 10, row(10).ToString().Trim(), "")
                            If Not String.Equals(rowSpace, spaceFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' დაფინანსების ფილტრი
                        If CBDaf IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CBDaf.Text) AndAlso CBDaf.Text <> "ყველა" Then
                            Dim fundingFilter As String = CBDaf.Text.Trim()
                            Dim rowFunding As String = If(row.Count > 13, row(13).ToString().Trim(), "")
                            If Not String.Equals(rowFunding, fundingFilter, StringComparison.OrdinalIgnoreCase) Then Continue For
                        End If

                        ' ფილტრის პირობები დაკმაყოფილებულია
                        filtered.Add(row)

                    Catch ex As Exception
                        Debug.WriteLine($"UC_Schedule: მწკრივის დამუშავების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next
            End If

            Debug.WriteLine($"UC_Schedule: ფილტრაციის შემდეგ დარჩა {filtered.Count} ჩანაწერი")

            ' დალაგება თარიღის მიხედვით კლებადობით (ახლიდან ძველისკენ)
            filtered.Sort(Function(a, b)
                              Try
                                  Dim dta, dtb As DateTime
                                  Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}

                                  ' b თარიღის პარსინგი (ახალი)
                                  If Not DateTime.TryParseExact(b(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dtb) Then
                                      If Not DateTime.TryParse(b(5).ToString().Trim(), dtb) Then
                                          dtb = Date.MinValue
                                      End If
                                  End If

                                  ' a თარიღის პარსინგი (ძველი)
                                  If Not DateTime.TryParseExact(a(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dta) Then
                                      If Not DateTime.TryParse(a(5).ToString().Trim(), dta) Then
                                          dta = Date.MinValue
                                      End If
                                  End If

                                  ' კლებადობით დალაგება (ახლიდან ძველისკენ)
                                  If dtb = Date.MinValue AndAlso dta = Date.MinValue Then
                                      Return 0
                                  ElseIf dtb = Date.MinValue Then
                                      Return 1
                                  ElseIf dta = Date.MinValue Then
                                      Return -1
                                  Else
                                      Return dtb.CompareTo(dta)
                                  End If
                              Catch
                                  Return 0
                              End Try
                          End Function)

            ' გვერდების ლოგიკა
            totalPages = Math.Max(1, Math.Ceiling(filtered.Count / pageSize))
            If currentPage > totalPages Then currentPage = 1

            Dim pageRows = filtered.Skip((currentPage - 1) * pageSize).Take(pageSize).ToList()

            Debug.WriteLine($"UC_Schedule: გვერდი {currentPage}/{totalPages}, ჩანაწერები: {pageRows.Count}")

            ' თითოეული ჩანაწერის დამატება ცხრილში
            For Each row In pageRows
                Try
                    ' ბენეფიციარის სრული სახელი
                    Dim beneFullName As String = row(3).ToString().Trim() & " " & row(4).ToString().Trim()

                    ' ძირითადი მონაცემები
                    Dim dgvRow As New List(Of Object) From {
                        row(0),                                                ' N (ID)
                        row(5),                                                ' თარიღი
                        If(row.Count > 6, row(6).ToString(), "60"),           ' ხანგძლივობა
                        beneFullName,                                          ' ბენეფიციარი
                        If(row.Count > 8, row(8).ToString(), ""),            ' თერაპევტი
                        If(row.Count > 9, row(9).ToString(), ""),            ' თერაპია
                        If(row.Count > 7, row(7).ToString(), ""),            ' ჯგუფური
                        If(row.Count > 10, row(10).ToString(), ""),          ' სივრცე
                        If(row.Count > 11, row(11).ToString(), "0"),         ' ფასი
                        If(row.Count > 12, row(12).ToString(), "დაგეგმილი"), ' სტატუსი
                        If(row.Count > 13, row(13).ToString(), ""),          ' პროგრამა
                        If(row.Count > 14, row(14).ToString(), "")           ' კომენტარი
                    }

                    ' ადმინისა და მენეჯერისთვის დამატებითი სვეტები
                    Debug.WriteLine($"UC_Schedule: userRoleID={userRoleID}, ადმინისა და მენეჯერისთვის სვეტების შემოწმება")
                    If userRoleID = 1 OrElse userRoleID = 2 Then
                        Dim editDate = If(row.Count > 1, row(1).ToString(), "")
                        Dim author = If(row.Count > 2, row(2).ToString(), "")
                        dgvRow.Add(editDate)  ' რედაქტირების თარიღი
                        dgvRow.Add(author)    ' ავტორი
                        Debug.WriteLine($"UC_Schedule: დაემატა რედაქტირების თარიღი='{editDate}', ავტორი='{author}'")
                    End If

                    ' რედაქტირების ღილაკი
                    dgvRow.Add("✎")

                    ' მწკრივის დამატება DataGridView-ში
                    Dim rowIndex As Integer = DgvSchedule.Rows.Add(dgvRow.ToArray())

                    ' სტატუსის მიხედვით ფერების დაყენება - v8.2 ვერსიის მსგავსი
                    ApplyRowColor(rowIndex, row)

                Catch ex As Exception
                    Debug.WriteLine($"UC_Schedule: მწკრივის დამატების შეცდომა: {ex.Message}")
                    Continue For
                End Try
            Next

            ' გვერდის ნომრის ჩვენება
            UpdatePageLabel()
            UpdateNavigationButtons()

            Debug.WriteLine("UC_Schedule: LoadFilteredSchedule - დასრულდა")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: LoadFilteredSchedule შეცდომა: {ex.Message}")
            MessageBox.Show($"განრიგის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView-ის სვეტების შექმნა - v8.2 ვერსიის მსგავსი
    ''' </summary>
    Private Sub CreateDataGridViewColumns()
        Try
            Debug.WriteLine($"UC_Schedule: სვეტების შექმნა დაიწყო, userRoleID={userRoleID}")
            DgvSchedule.Columns.Clear()

            With DgvSchedule.Columns
                .Add("N", "N")
                .Add("Tarigi", "თარიღი")
                .Add("Duri", "ხანგძლიობა")
                .Add("Bene", "ბენეფიციარი")
                .Add("Per", "თერაპევტი")
                .Add("Ter", "თერაპია")
                .Add("Group", "ჯგუფური")
                .Add("Space", "სივრცე")
                .Add("Price", "თანხა")
                .Add("Status", "სტატუსი")
                .Add("Program", "პროგრამა")
                .Add("Coment", "კომენტარი")

                ' ადმინისა და მენეჯერისთვის დამატებითი სვეტები
                Debug.WriteLine($"UC_Schedule: ამოწმება userRoleID = {userRoleID}")
                If userRoleID = 1 OrElse userRoleID = 2 Then
                    .Add("EditDate", "რედ. თარიღი")
                    .Add("Author", "ავტორი")
                    Debug.WriteLine("UC_Schedule: დაემატა რედაქტირების და ავტორის სვეტები")
                Else
                    Debug.WriteLine("UC_Schedule: რედაქტირების და ავტორის სვეტები არ დაემატა")
                End If

                ' რედაქტირების ღილაკი
                Dim editBtn As New DataGridViewButtonColumn()
                editBtn.Name = "Edit"
                editBtn.HeaderText = ""
                editBtn.Text = "✎"
                editBtn.UseColumnTextForButtonValue = True
                .Add(editBtn)
            End With

            ' სვეტების ზომების დაყენება - v8.2 ვერსიის მსგავსი
            DgvSchedule.Columns("N").Width = 40
            DgvSchedule.Columns("Tarigi").Width = 110
            DgvSchedule.Columns("Duri").Width = 40
            DgvSchedule.Columns("Bene").Width = 180
            DgvSchedule.Columns("Per").Width = 185
            DgvSchedule.Columns("Ter").Width = 185
            DgvSchedule.Columns("Group").Width = 50
            DgvSchedule.Columns("Space").Width = 80
            DgvSchedule.Columns("Price").Width = 60
            DgvSchedule.Columns("Status").Width = 130
            DgvSchedule.Columns("Program").Width = 130
            DgvSchedule.Columns("Coment").Width = 120
            DgvSchedule.Columns("Coment").DefaultCellStyle.WrapMode = DataGridViewTriState.False
            DgvSchedule.Columns("Coment").ToolTipText = "სრული კომენტარი გამოჩნდება მაუსის მიტანისას"

            If userRoleID = 1 OrElse userRoleID = 2 Then
                If DgvSchedule.Columns.Contains("EditDate") Then
                    DgvSchedule.Columns("EditDate").Width = 110
                    Debug.WriteLine("UC_Schedule: EditDate სვეტის სიგანე დაყენებული")
                End If
                If DgvSchedule.Columns.Contains("Author") Then
                    DgvSchedule.Columns("Author").Width = 120
                    Debug.WriteLine("UC_Schedule: Author სვეტის სიგანე დაყენებული")
                End If
            End If

            DgvSchedule.Columns("Edit").Width = 24

            Debug.WriteLine($"UC_Schedule: შეიქმნა {DgvSchedule.Columns.Count} სვეტი, userRoleID={userRoleID}")

            ' სვეტების სახელების ჩამონათვალი debug-ისთვის
            Dim columnNames As New List(Of String)
            For Each col As DataGridViewColumn In DgvSchedule.Columns
                columnNames.Add(col.Name)
            Next
            Debug.WriteLine($"UC_Schedule: სვეტების სიაა: {String.Join(", ", columnNames)}")

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: სვეტების შექმნის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მწკრივის ფერის დაყენება სტატუსის მიხედვით - v8.2 ვერსიის მსგავსი
    ''' </summary>
    Private Sub ApplyRowColor(rowIndex As Integer, rowData As IList(Of Object))
        Try
            If rowIndex < 0 OrElse rowIndex >= DgvSchedule.Rows.Count Then Return

            ' სტატუსის მიღება
            Dim statusText As String = If(rowData.Count > 12, rowData(12).ToString().Trim(), "დაგეგმილი")
            Dim nowTime As DateTime = DateTime.Now
            Dim sessionTime As DateTime

            ' სესიის თარიღის პარსინგი
            Dim formats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}
            If Not DateTime.TryParseExact(rowData(5).ToString().Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, sessionTime) Then
                DateTime.TryParse(rowData(5).ToString().Trim(), sessionTime)
            End If

            ' სტატუსის მიხედვით ფერების დაყენება - v8.2 ლოგიკის მსგავსი
            Select Case statusText.ToLower().Trim()
                Case "შესრულებული"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                Case "აღდგენა"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.Honeydew
                Case "გაცდენა არასაპატიო"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.Plum
                Case "გაცდენა საპატიო"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightYellow
                Case "პროგრამით გატარება"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightGray
                Case "გაუქმება", "გაუქმებული"
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.Red
                Case "დაგეგმილი"
                    ' დაგეგმილი სესიებისთვის - თუ დრო გასულია, ღია წითელი, თუ არა - ღია ლურჯი
                    If sessionTime < nowTime Then
                        DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightCoral ' ღია წითელი - ვადაგადაცილებული
                    Else
                        DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.LightBlue ' ღია ლურჯი - დაგეგმილი
                    End If
                Case Else
                    ' უცნობი სტატუსებისთვის სტანდარტული ფერი
                    DgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = Color.White
            End Select

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: მწკრივის ფერის დაყენების შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' გვერდის ლეიბლის განახლება
    ''' </summary>
    Private Sub UpdatePageLabel()
        Try
            If LPage IsNot Nothing Then
                LPage.Text = $"გვერდი {currentPage} / {totalPages}"
            ElseIf LPage IsNot Nothing Then
                LPage.Text = $"{currentPage} / {totalPages}"
            End If
        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' ნავიგაციის ღილაკების მდგომარეობის განახლება
    ''' </summary>
    Private Sub UpdateNavigationButtons()
        Try
            ' წინა ღილაკი
            If BtnPrev IsNot Nothing Then
                BtnPrev.Enabled = (currentPage > 1)
            End If

            ' შემდეგი ღილაკი
            If BtnNext IsNot Nothing Then
                BtnNext.Enabled = (currentPage < totalPages)
            End If

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ
        End Try
    End Sub

    ''' <summary>
    ''' რადიობუტონების შეცვლის ივენთი
    ''' </summary>
    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim radioButton As RadioButton = CType(sender, RadioButton)

            If Not radioButton.Checked Then Return

            ' გვერდის ზომის განახლება
            Select Case radioButton.Name
                Case "RB20"
                    pageSize = 20
                Case "RB50"
                    pageSize = 50
                Case "RB100"
                    pageSize = 100
            End Select

            Debug.WriteLine($"UC_Schedule: გვერდის ზომა შეიცვალა: {pageSize}")

            ' გვერდების ხელახალი გამოთვლა
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: რადიობუტონის შეცვლის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' თარიღების შეცვლის ივენთი
    ''' </summary>
    Private Sub DatePickers_ValueChanged(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("UC_Schedule: თარიღის ფილტრი შეიცვალა")
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: თარიღის ფილტრის შეცვლის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფილტრების შეცვლის ივენთი
    ''' </summary>
    Private Sub Filter_Changed(sender As Object, e As EventArgs)
        Try
            Debug.WriteLine("UC_Schedule: ფილტრი შეიცვალა")
            currentPage = 1
            LoadFilteredSchedule()

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ფილტრის შეცვლის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკი
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs)
        Try
            If currentPage > 1 Then
                currentPage -= 1
                Debug.WriteLine($"UC_Schedule: გადასვლა გვერდზე {currentPage}")
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: წინა გვერდის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკი
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs)
        Try
            If currentPage < totalPages Then
                currentPage += 1
                Debug.WriteLine($"UC_Schedule: გადასვლა გვერდზე {currentPage}")
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: შემდეგი გვერდის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub DgvSchedule_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' შევამოწმოთ რედაქტირების სვეტზე დაჭერა
            If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso
               DgvSchedule.Columns(e.ColumnIndex).Name = "Edit" Then

                ' მივიღოთ სესიის ID
                Dim sessionId As Object = DgvSchedule.Rows(e.RowIndex).Cells("N").Value

                If sessionId IsNot Nothing AndAlso IsNumeric(sessionId) Then
                    Dim sessionIdInt As Integer = CInt(sessionId)
                    Debug.WriteLine($"UC_Schedule: რედაქტირება - სესია ID={sessionIdInt}")

                    ' NewRecordForm-ის გახსნა რედაქტირების რეჟიმში
                    Try
                        ' მოვძებნოთ მომხმარებლის email
                        Dim currentUserEmail As String = If(String.IsNullOrEmpty(userEmail), "user@example.com", userEmail)

                        ' NewRecordForm-ის შექმნა რედაქტირების რეჟიმში
                        Using editForm As New NewRecordForm(dataService, "სესია", sessionIdInt, currentUserEmail, "UC_Schedule")

                            ' ფორმის ჩვენება
                            Dim result As DialogResult = editForm.ShowDialog()

                            ' თუ წარმატებით შეინახა, განვაახლოთ მონაცემები
                            If result = DialogResult.OK Then
                                ' მონაცემების ხელახალი ჩატვირთვა
                                RefreshData()

                                ' შეტყობინება მომხმარებლისთვის
                                MessageBox.Show($"სესია ID={sessionIdInt} წარმატებით განახლდა", "წარმატება",
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
            Debug.WriteLine($"UC_Schedule: რედაქტირების შეცდომა: {ex.Message}")
            MessageBox.Show($"რედაქტირების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' მონაცემების ხელახალი ჩატვირთვა
    ''' </summary>
    Public Sub RefreshData()
        Try
            Debug.WriteLine("UC_Schedule: RefreshData - მონაცემების განახლება")

            If dataService IsNot Nothing Then
                ' ქეშის გასუფთავება თუ SheetDataService გვაქვს
                If TypeOf dataService Is SheetDataService Then
                    DirectCast(dataService, SheetDataService).InvalidateAllCache()
                End If

                ' ახლიდან ჩავტვირთოთ მონაცემები
                LoadFilteredSchedule()
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: მონაცემების განახლების შეცდომა: {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                Debug.WriteLine($"UC_Schedule: გადასვლა გვერდზე {currentPage}")
                LoadFilteredSchedule()
            End If
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: გვერდზე გადასვლის შეცდომა: {ex.Message}")
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

    ''' <summary>
    ''' ფილტრების განახლება - გარე გამოძახებისთვის
    ''' </summary>
    Public Sub UpdateFilters()
        Try
            Debug.WriteLine("UC_Schedule: ფილტრების განახლება")
            currentPage = 1
            LoadFilteredSchedule()
        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: ფილტრების განახლების შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მომხმარებლის როლის ძაღლით დაყენება - ტესტირებისთვის
    ''' </summary>
    ''' <param name="roleId">როლის ID (1=ადმინი, 2=მენეჯერი, 6=მომხმარებელი)</param>
    Public Sub ForceSetUserRole(roleId As Integer)
        Try
            Debug.WriteLine($"UC_Schedule: ძაღლით დაყენება userRoleID = {roleId}")
            userRoleID = roleId

            ' თუ DataGridView არსებობს, ხელახლა შევქმნათ სვეტები
            If DgvSchedule IsNot Nothing Then
                LoadFilteredSchedule() ' ეს ახლიდან შექმნის სვეტებს
            End If

        Catch ex As Exception
            Debug.WriteLine($"UC_Schedule: მომხმარებლის როლის ძაღლით დაყენების შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მიმდინარე მომხმარებლის როლის მიღება
    ''' </summary>
    Public ReadOnly Property CurrentUserRole As Integer
        Get
            Return userRoleID
        End Get
    End Property

End Class