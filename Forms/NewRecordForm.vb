' ===========================================
' 📄 Forms/NewRecordForm.vb
' -------------------------------------------
' ახალი ჩანაწერის ფორმა - გამოიყენება ახალი სესიის, დაბადების დღის ან დავალების დასამატებლად
' ასევე არსებული ჩანაწერების რედაქტირებისთვის
' ===========================================
Imports System.Globalization
Imports System.Threading
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class NewRecordForm
    ' მონაცემთა სერვისის მითითება
    Private ReadOnly dataService As IDataService
    ' ჩანაწერის ტიპი (მაგ: "სესია", "დაბადების დღე", "დავალება")
    Private ReadOnly recordType As String
    ' ფორმის რეჟიმი: True - დამატება, False - რედაქტირება
    Private _isAddMode As Boolean = True
    ' ინფორმაციის წყარო ფორმის/კონტროლის სახელი
    Private _sourceControl As String = ""
    ' რედაქტირების რეჟიმში ჩანაწერის ID
    Private _editRecordId As Integer = 0
    ' მომხმარებლის ელ.ფოსტა
    Private _userEmail As String = ""
    ' ფლაგი რეკურსიის თავიდან ასაცილებლად
    Private isUpdating As Boolean = False

    ''' <summary>
    ''' კონსტრუქტორი ახალი ჩანაწერის შესაქმნელად
    ''' </summary>
    ''' <param name="dataService">მონაცემთა სერვისი</param>
    ''' <param name="recordType">ჩანაწერის ტიპი (მაგ: "სესია", "დაბადების დღე", "დავალება")</param>
    ''' <param name="userEmail">მომხმარებლის ელ-ფოსტა</param>
    ''' <param name="sourceControl">ინფორმაციის წყარო ფორმის/კონტროლის სახელი</param>
    Public Sub New(dataService As IDataService, recordType As String, userEmail As String, Optional sourceControl As String = "UC_Home")
        InitializeComponent()
        Me.dataService = dataService
        Me.recordType = recordType
        _sourceControl = sourceControl
        _isAddMode = True  ' დაყენება Add რეჟიმში
        _userEmail = userEmail

        ' ფორმის ინიციალიზაცია ტიპის მიხედვით
        Me.Text = $"ახალი {recordType} - დამატება"
    End Sub

    ''' <summary>
    ''' კონსტრუქტორი არსებული ჩანაწერის რედაქტირებისთვის
    ''' </summary>
    ''' <param name="dataService">მონაცემთა სერვისი</param>
    ''' <param name="recordType">ჩანაწერის ტიპი</param>
    ''' <param name="recordId">ჩანაწერის ID</param>
    ''' <param name="userEmail">მომხმარებლის ელ-ფოსტა</param>
    ''' <param name="sourceControl">ინფორმაციის წყარო ფორმის/კონტროლის სახელი</param>
    Public Sub New(dataService As IDataService, recordType As String, recordId As Integer, userEmail As String, Optional sourceControl As String = "UC_Home")
        InitializeComponent()
        Me.dataService = dataService
        Me.recordType = recordType
        _sourceControl = sourceControl
        _isAddMode = False  ' დაყენება Edit რეჟიმში
        _editRecordId = recordId
        _userEmail = userEmail

        ' ფორმის ინიციალიზაცია ტიპის მიხედვით
        Me.Text = $"არსებული {recordType} - რედაქტირება"
    End Sub


    ''' <summary>
    ''' პოულობს მაქსიმალურ ID-ს DB-Schedule ფურცლიდან
    ''' </summary>
    ''' <returns>მაქსიმალური ID</returns>
    Private Function GetMaxRecordId() As Integer
        Try
            ' ჩანაწერის ტიპის მიხედვით სწორი ფურცლის შერჩევა
            Dim sheetName As String = ""

            Select Case recordType.ToLower()
                Case "სესია"
                    sheetName = "DB-Schedule"
                Case "დაბადების დღე"
                    sheetName = "DB-Personal"
                Case "დავალება"
                    sheetName = "DB-Tasks"
                Case Else
                    sheetName = "DB-Schedule"
            End Select

            ' A სვეტის წაკითხვა ფურცლიდან
            Dim rows = dataService.GetData($"{sheetName}!A2:A")
            Dim maxId As Integer = 0

            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString()) Then
                        Dim id As Integer
                        If Integer.TryParse(row(0).ToString(), id) Then
                            If id > maxId Then
                                maxId = id
                            End If
                        End If
                    End If
                Next
            End If

            Return maxId
        Catch ex As Exception
            MessageBox.Show($"შეცდომა მაქსიმალური ID-ის მოძიებისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function
    ' აქ უკვე არსებულ კოდს ვამატებთ შემდეგ მეთოდებს და ივენთებს:

    ''' <summary>
    ''' ფორმის ჩატვირთვა - დამატებულია საათის, წუთის და ხანგძლივობის ელემენტების ინიციალიზაცია
    ''' </summary>
    Private Sub NewRecordForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' მიმდინარე თარიღის და დროის დაყენება
            LNow.Text = DateTime.Now.ToString("dd.MM.yyyy HH:mm")

            ' ავტორიზირებული მომხმარებლის ელ.ფოსტის გამოყენება
            LAutor.Text = _userEmail

            ' ჩანაწერის ნომრის (ID) მოპოვება რეჟიმის მიხედვით
            If _isAddMode Then
                ' ახალი ჩანაწერისთვის - მივიღოთ მაქსიმალური ID + 1
                Dim maxId As Integer = GetMaxRecordId()
                LN.Text = (maxId + 1).ToString()
            Else
                ' რედაქტირების რეჟიმისთვის - გამოვიყენოთ არსებული ID
                LN.Text = _editRecordId.ToString()
            End If

            ' ComboBox-ების რედაქტირების შეზღუდვა - მხოლოდ ჩამონათვალიდან არჩევა
            CBBeneName.DropDownStyle = ComboBoxStyle.DropDownList
            CBBeneSurname.DropDownStyle = ComboBoxStyle.DropDownList
            CBPer.DropDownStyle = ComboBoxStyle.DropDownList
            CBTer.DropDownStyle = ComboBoxStyle.DropDownList
            CBDaf.DropDownStyle = ComboBoxStyle.DropDownList

            ' ბენეფიციარის სახელების ჩატვირთვა
            LoadBeneNames()

            ' თერაპევტების სიის ჩატვირთვა
            LoadTherapists()

            ' თერაპიის ტიპების ჩატვირთვა
            LoadTherapyTypes()

            ' დაფინანსების პროგრამების ჩატვირთვა
            LoadFundingPrograms()

            ' CBBeneSurname თავიდან დაბლოკილია
            CBBeneSurname.Enabled = False

            ' DateTimePicker (DTP1) კონფიგურაცია ქართული ფორმატით
            ConfigureDateTimePicker()

            ' საათის, წუთის და ხანგძლივობის ლეიბლების ინიციალიზაცია
            InitializeTimeAndDurationLabels()

            ' TCost ტექსტბოქსის ინიციალიზაცია
            TCost.Text = "0"

            ' სივრცეების ღილაკების საწყისი ფერების დაყენება
            InitializeSpaceButtons()

            ' რადიობუტონების ხილვადობის კონტროლი თარიღის მიხედვით
            UpdateStatusVisibility()

            ' მორგებული კომბობოქსების დაყენება
            SetupCustomComboBoxes()

            ' თარიღისა და დროის ცვლილების ივენთების დამატება
            ' ValueChanged ივენთს აღარ ვიყენებთ, მის ნაცვლად CloseUp ივენთს ვიყენებთ
            RemoveHandler DTP1.ValueChanged, AddressOf DTP1_ValueChanged
            AddHandler DTP1.CloseUp, AddressOf DTP1_CloseUp
            AddHandler THour.TextChanged, AddressOf DateTimeChanged
            AddHandler TMin.TextChanged, AddressOf DateTimeChanged

            ' სივრცეების ღილაკებზე click ივენთების მიბმა
            AttachSpaceButtonHandlers()

            ' რადიობუტონების ივენთების მიბმა
            AttachRadioButtonHandlers()

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ფორმის ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ფორმის დახურვისას მივხსნათ ივენთის მიბმები
    ''' </summary>
    Private Sub NewRecordForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            ' მოვხსნათ ყველა მიბმული ივენთი რესურსების გასათავისუფლებლად
            RemoveHandler DTP1.CloseUp, AddressOf DTP1_CloseUp
            RemoveHandler THour.TextChanged, AddressOf DateTimeChanged
            RemoveHandler TMin.TextChanged, AddressOf DateTimeChanged

            ' სივრცის ღილაკებიდან ივენთების მოხსნა
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") Then
                    RemoveHandler btn.Click, AddressOf SpaceButton_Click
                End If
            Next

            ' რადიობუტონებიდან ივენთების მოხსნა
            For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                If rb.Name.StartsWith("RB") Then
                    RemoveHandler rb.CheckedChanged, AddressOf StatusRadioButton_CheckedChanged
                End If
            Next

            ' მოვხსნათ კომბობოქსების DrawItem ივენთები
            RemoveHandler CBBeneName.DrawItem, AddressOf ComboBox_DrawItem
            RemoveHandler CBBeneSurname.DrawItem, AddressOf ComboBox_DrawItem

        Catch ex As Exception
            ' უბრალოდ გავაგრძელოთ, შეცდომა აქ კრიტიკული არ არის
            Debug.WriteLine($"FormClosing ივენთების მოხსნის შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' საათის, წუთის და ხანგძლივობის ლეიბლების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeTimeAndDurationLabels()
        ' საათი - საწყისი მნიშვნელობა 12
        THour.Text = "12"

        ' წუთი - საწყისი მნიშვნელობა 0
        TMin.Text = "00"

        ' ხანგძლივობა - საწყისი მნიშვნელობა 60
        TDur.Text = "60"

        ' თუ კონტროლები ტექსტბოქსებია, მაშინ გამოვიყენოთ TextAlign
        If TypeOf THour Is TextBox Then
            DirectCast(THour, TextBox).TextAlign = HorizontalAlignment.Center
            DirectCast(TMin, TextBox).TextAlign = HorizontalAlignment.Center
            DirectCast(TDur, TextBox).TextAlign = HorizontalAlignment.Center
        End If

        ' ლეიბლების საზღვრები
        THour.BorderStyle = BorderStyle.FixedSingle
        TMin.BorderStyle = BorderStyle.FixedSingle
        TDur.BorderStyle = BorderStyle.FixedSingle
    End Sub

    ''' <summary>
    ''' DateTimePicker-ის კონფიგურაცია ქართული ფორმატით
    ''' </summary>
    Private Sub ConfigureDateTimePicker()
        Try
            ' დავიმპორტოთ ქართული კულტურა
            Dim georgianCulture As New Globalization.CultureInfo("ka-GE")

            ' საწყისი თარიღის დაყენება მიმდინარე დღეზე
            DTP1.Value = DateTime.Today

            ' ფორმატის დაყენება ქართული სტილით
            DTP1.Format = DateTimePickerFormat.Custom
            DTP1.CustomFormat = "dd MMMM yyyy 'წელი'"

            ' კულტურის დაყენება
            DTP1.CalendarFont = New Font("Sylfaen", 10) ' ქართული შრიფტი
            Thread.CurrentThread.CurrentCulture = georgianCulture

            ' ვცადოთ ფორმატირებული დღევანდელი თარიღი
            Dim formattedDate = DateTime.Today.ToString("dd MMMM yyyy", georgianCulture)
            Debug.WriteLine($"ფორმატირებული თარიღი: {formattedDate}")

        Catch ex As Exception
            MessageBox.Show($"შეცდომა DateTimePicker-ის კონფიგურაციისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    '''     სივრცეების ღილაკების საწყისი ფერების დაყენება
    ''' </summary>
    Private Sub SetupCustomComboBoxes()
        ' შევცვალოთ კომბობოქსების DrawMode, რომ ხელით ვხატოთ
        CBBeneName.DrawMode = DrawMode.OwnerDrawFixed
        CBBeneSurname.DrawMode = DrawMode.OwnerDrawFixed
        CBPer.DrawMode = DrawMode.OwnerDrawFixed ' დავამატეთ CBPer

        ' მივაბათ DrawItem ივენთი CBBeneName-სთვის
        AddHandler CBBeneName.DrawItem, AddressOf ComboBox_DrawItem

        ' მივაბათ DrawItem ივენთი CBBeneSurname-სთვის
        AddHandler CBBeneSurname.DrawItem, AddressOf ComboBox_DrawItem

        ' მივაბათ DrawItem ივენთი CBPer-სთვის
        AddHandler CBPer.DrawItem, AddressOf ComboBox_DrawItem

        Debug.WriteLine("კომბობოქსების DrawMode შეცვლილია")
    End Sub

    ' დამატებითი მეთოდი კომბობოქსის ელემენტების დასახატად
    Private Sub ComboBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        Try
            If e.Index < 0 Then Return

            Dim combo As ComboBox = DirectCast(sender, ComboBox)

            ' გამოვიყენოთ კომბობოქსის არსებული ფერი (BackColor)
            Dim backColor As Color = combo.BackColor

            ' არჩეული ელემენტისთვის გამოვიყენოთ სტანდარტული ჰაილაითი
            If (e.State And DrawItemState.Selected) > 0 Then
                ' სისტემური ფერის გამოყენება არჩეული ელემენტისთვის
                e.DrawBackground()
            Else
                ' კომბობოქსის ფერის გამოყენება ჩვეულებრივი ელემენტებისთვის
                Using brush As New SolidBrush(backColor)
                    e.Graphics.FillRectangle(brush, e.Bounds)
                End Using
            End If

            ' ტექსტის დახატვა
            If e.Index < combo.Items.Count Then
                Dim itemText As String = combo.Items(e.Index).ToString()
                Dim textBrush As New SolidBrush(e.ForeColor)

                ' ვხატავთ ტექსტს
                e.Graphics.DrawString(itemText, e.Font, textBrush,
                                 New RectangleF(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height))

                textBrush.Dispose()
            End If

            ' ფოკუსის მართკუთხედი (თუ საჭიროა)
            If (e.State And DrawItemState.Focus) > 0 Then
                e.DrawFocusRectangle()
            End If
        Catch ex As Exception
            Debug.WriteLine($"ComboBox_DrawItem შეცდომა: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' ბენეფიციარის სახელების ჩატვირთვა DB-Bene ფურცლიდან - განახლებული ვერსია
    ''' </summary>
    Private Sub LoadBeneNames()
        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' B სვეტის წაკითხვა DB-Bene ფურცლიდან
            Dim rows = dataService.GetData("DB-Bene!B2:B")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("ბენეფიციარების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' უნიკალური სახელების შეგროვება HashSet-ის გამოყენებით
            Dim uniqueNames As New HashSet(Of String)

            For Each row In rows
                If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) Then
                    uniqueNames.Add(row(0).ToString().Trim())
                End If
            Next

            ' გასუფთავება ComboBox-ის
            CBBeneName.Items.Clear()

            ' დავამატოთ პირველი ელემენტი - მინიშნება
            CBBeneName.Items.Add("- აირჩიეთ ბენეფიციარის სახელი -")

            ' უნიკალური სახელების სორტირება ანბანის მიხედვით და დამატება CBBeneName-ში
            Dim sortedNamesList As List(Of String) = uniqueNames.OrderBy(Function(name) name).ToList()

            For i As Integer = 0 To sortedNamesList.Count - 1
                CBBeneName.Items.Add(sortedNamesList(i))
            Next

            ' ავირჩიოთ პირველი ელემენტი (მინიშნება)
            CBBeneName.SelectedIndex = 0

            ' საწყისი ფონის ფერი უნდა იყოს ჩვეულებრივი
            CBBeneName.BackColor = SystemColors.Window
            CBBeneSurname.BackColor = SystemColors.Window

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ბენეფიციარის სახელების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneName-ს SelectedIndexChanged ივენთი - ბენეფიციარის სახელის არჩევისას
    ''' აქტიურდება CBBeneSurname და ტვირთავს შესაბამის გვარებს
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            ' გავასუფთაოთ CBBeneSurname-ის Items კოლექცია
            CBBeneSurname.Items.Clear()

            ' თუ ინდექსი ვალიდურია და არაა პირველი ელემენტი, ჩავტვირთოთ შესაბამისი გვარები
            If CBBeneName.SelectedIndex > 0 Then
                Dim selectedName As String = CBBeneName.SelectedItem.ToString()
                LoadBeneSurnames(selectedName)

                ' გავააქტიუროთ გვარების ჩამონათვალი
                CBBeneSurname.Enabled = True

                ' ბენეფიციარის მდგომარეობის შემოწმება
                CheckBeneficiaryAvailability()
            Else
                ' ინდექსი არავალიდურია ან არჩეულია პირველი ელემენტი, გამოვრთოთ გვარების ჩამონათვალი
                CBBeneSurname.Enabled = False

                ' გავწმინდოთ ბექგრაუნდი და შეტყობინება
                CBBeneName.BackColor = SystemColors.Window
                CBBeneSurname.BackColor = SystemColors.Window
                LMsgBene.Text = "აირჩიეთ ბენეფიციარი"
                LMsgBene.ForeColor = Color.Black
            End If

        Catch ex As Exception
            MessageBox.Show($"შეცდომა სახელის არჩევისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneName-ში ტექსტის შეცვლისას ივენთი
    ''' </summary>
    Private Sub CBBeneName_TextChanged(sender As Object, e As EventArgs) Handles CBBeneName.TextChanged
        ' გავასუფთაოთ CBBeneSurname-ის ტექსტი და Items კოლექცია
        CBBeneSurname.Text = ""
        CBBeneSurname.Items.Clear()

        ' თუ არჩეული სახელი აღარ ემთხვევა ტექსტს, გვარების ჩამონათვალი დაბლოკილია
        If CBBeneName.SelectedIndex < 0 OrElse CBBeneName.Text <> CBBeneName.SelectedItem.ToString() Then
            CBBeneSurname.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' ბენეფიციარის გვარების ჩატვირთვა არჩეული სახელის შესაბამისად - განახლებული
    ''' </summary>
    ''' <param name="selectedName">არჩეული სახელი</param>
    Private Sub LoadBeneSurnames(selectedName As String)
        Try
            ' B და C სვეტების წაკითხვა DB-Bene ფურცლიდან
            Dim rows = dataService.GetData("DB-Bene!B2:C")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("ბენეფიციარების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' მოძებნა და შეგროვება შესაბამისი გვარების
            Dim matchingSurnames As New HashSet(Of String)

            For Each row In rows
                ' შევამოწმოთ არის თუ არა საკმარისი სვეტები და ემთხვევა თუ არა სახელი
                If row.Count >= 2 AndAlso
               Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) AndAlso
               row(0).ToString().Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrEmpty(row(1)?.ToString().Trim()) Then

                    matchingSurnames.Add(row(1).ToString().Trim())
                End If
            Next

            ' დაალაგეთ გვარები ანბანის მიხედვით და დაამატეთ ComboBox-ში
            Dim sortedSurnamesList As List(Of String) = matchingSurnames.OrderBy(Function(surname) surname).ToList()

            For i As Integer = 0 To sortedSurnamesList.Count - 1
                CBBeneSurname.Items.Add(sortedSurnamesList(i))
            Next

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ბენეფიციარის გვარების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' CBBeneSurname-ს SelectedIndexChanged ივენთი - ბენეფიციარის გვარის არჩევისას
    ''' </summary>
    Private Sub CBBeneSurname_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneSurname.SelectedIndexChanged
        ' ბენეფიციარის მდგომარეობის შემოწმება
        CheckBeneficiaryAvailability()
    End Sub

    ''' <summary>
    ''' CBPer-ს SelectedIndexChanged ივენთი - თერაპევტის არჩევისას
    ''' </summary>
    Private Sub CBPer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPer.SelectedIndexChanged
        Try
            ' შევამოწმოთ თერაპევტის მდგომარეობა
            CheckTherapistAvailability()
        Catch ex As Exception
            MessageBox.Show($"შეცდომა თერაპევტის არჩევისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNHourDown ღილაკზე დაჭერა - საათის შემცირება 1-ით
    ''' </summary>
    Private Sub BTNHourDown_Click(sender As Object, e As EventArgs) Handles BTNHourDown.Click
        Try
            Dim hour As Integer
            If Integer.TryParse(THour.Text, hour) Then
                ' შევამციროთ საათი 1-ით, მინიმუმ 0
                hour -= 1
                If hour < 0 Then hour = 0

                ' განახლება და ფორმატირება 2 ციფრად
                THour.Text = hour.ToString("00")
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                THour.Text = "00"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა საათის შემცირებისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNHourUp ღილაკზე დაჭერა - საათის გაზრდა 1-ით
    ''' </summary>
    Private Sub BTNHourUp_Click(sender As Object, e As EventArgs) Handles BTNHourUp.Click
        Try
            Dim hour As Integer
            If Integer.TryParse(THour.Text, hour) Then
                ' გავზარდოთ საათი 1-ით, მაქსიმუმ 23
                hour += 1
                If hour > 23 Then hour = 23

                ' განახლება და ფორმატირება 2 ციფრად
                THour.Text = hour.ToString("00")
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                THour.Text = "00"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა საათის გაზრდისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNMinDown ღილაკზე დაჭერა - წუთის შემცირება 5-ით
    ''' </summary>
    Private Sub BTNMinDown_Click(sender As Object, e As EventArgs) Handles BTNMinDown.Click
        Try
            Dim minute As Integer
            If Integer.TryParse(TMin.Text, minute) Then
                ' შევამციროთ წუთი 5-ით, მინიმუმ 0
                minute -= 5
                If minute < 0 Then minute = 0

                ' განახლება და ფორმატირება 2 ციფრად
                TMin.Text = minute.ToString("00")
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                TMin.Text = "00"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა წუთის შემცირებისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNMinUp ღილაკზე დაჭერა - წუთის გაზრდა 5-ით
    ''' </summary>
    Private Sub BTNMinUp_Click(sender As Object, e As EventArgs) Handles BTNMinUp.Click
        Try
            Dim minute As Integer
            If Integer.TryParse(TMin.Text, minute) Then
                ' გავზარდოთ წუთი 5-ით, მაქსიმუმ 55
                minute += 5
                If minute > 55 Then minute = 55

                ' განახლება და ფორმატირება 2 ციფრად
                TMin.Text = minute.ToString("00")
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                TMin.Text = "00"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა წუთის გაზრდისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNDurDown ღილაკზე დაჭერა - ხანგძლივობის შემცირება 5-ით
    ''' </summary>
    Private Sub BTNDurDown_Click(sender As Object, e As EventArgs) Handles BTNDurDown.Click
        Try
            Dim duration As Integer
            If Integer.TryParse(TDur.Text, duration) Then
                ' შევამციროთ ხანგძლივობა 5-ით, მინიმუმ 0
                duration -= 5
                If duration < 0 Then duration = 0

                ' განახლება
                TDur.Text = duration.ToString()
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                TDur.Text = "0"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა ხანგძლივობის შემცირებისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BTNDurUp ღილაკზე დაჭერა - ხანგძლივობის გაზრდა 5-ით
    ''' </summary>
    Private Sub BTNDurUp_Click(sender As Object, e As EventArgs) Handles BTNDurUp.Click
        Try
            Dim duration As Integer
            If Integer.TryParse(TDur.Text, duration) Then
                ' გავზარდოთ ხანგძლივობა 5-ით (მაქსიმუმი არ არის მითითებული)
                duration += 5

                ' განახლება
                TDur.Text = duration.ToString()
            Else
                ' თუ ტექსტი არ არის რიცხვი, დავაყენოთ 0
                TDur.Text = "0"
            End If
        Catch ex As Exception
            MessageBox.Show($"შეცდომა ხანგძლივობის გაზრდისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტების სიის ჩატვირთვა DB-Personal გვერდიდან - განახლებული
    ''' </summary>
    Private Sub LoadTherapists()
        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' B, C და H სვეტების წაკითხვა DB-Personal გვერდიდან
            Dim rows = dataService.GetData("DB-Personal!B2:H")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("თერაპევტების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' გავასუფთაოთ კომბობოქსი
            CBPer.Items.Clear()

            ' დავამატოთ პირველი ელემენტი - მინიშნება
            CBPer.Items.Add("- აირჩიეთ თერაპევტი -")

            ' თერაპევტების სიის შექმნა
            Dim therapists As New List(Of String)

            For Each row In rows
                ' შევამოწმოთ არის თუ არა საკმარისი სვეტები და არის თუ არა "active"
                If row.Count >= 7 AndAlso
               Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) AndAlso  ' B სვეტი - სახელი
               Not String.IsNullOrEmpty(row(1)?.ToString().Trim()) AndAlso  ' C სვეტი - გვარი
               row(6)?.ToString().Trim().ToLower() = "active" Then          ' H სვეტი - სტატუსი

                    ' შექმენით სახელი + გვარი ფორმატში
                    Dim therapistName = $"{row(0).ToString().Trim()} {row(1).ToString().Trim()}"
                    therapists.Add(therapistName)
                End If
            Next

            ' დავალაგოთ ანბანის მიხედვით
            therapists.Sort()

            ' დავამატოთ კომბობოქსში
            For Each therapist In therapists
                CBPer.Items.Add(therapist)
            Next

            ' ავირჩიოთ პირველი ელემენტი (მინიშნება)
            CBPer.SelectedIndex = 0

        Catch ex As Exception
            MessageBox.Show($"შეცდომა თერაპევტების სიის ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' თერაპიის ტიპების ჩატვირთვა DB-Therapy ფურცლიდან - განახლებული
    ''' </summary>
    Private Sub LoadTherapyTypes()
        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' B სვეტის წაკითხვა DB-Therapy ფურცლიდან
            Dim rows = dataService.GetData("DB-Therapy!B2:B")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("თერაპიის ტიპების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' გავასუფთაოთ კომბობოქსი
            CBTer.Items.Clear()

            ' დავამატოთ პირველი ელემენტი - მინიშნება
            CBTer.Items.Add("- აირჩიეთ თერაპიის ტიპი -")

            ' დავამატოთ თერაპიის ტიპები
            For Each row In rows
                If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) Then
                    CBTer.Items.Add(row(0).ToString().Trim())
                End If
            Next

            ' ავირჩიოთ პირველი ელემენტი (მინიშნება)
            CBTer.SelectedIndex = 0

        Catch ex As Exception
            MessageBox.Show($"შეცდომა თერაპიის ტიპების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დაფინანსების პროგრამების ჩატვირთვა DB-Program ფურცლიდან - განახლებული
    ''' </summary>
    Private Sub LoadFundingPrograms()
        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' B სვეტის წაკითხვა DB-Program ფურცლიდან
            Dim rows = dataService.GetData("DB-Program!B2:B")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("დაფინანსების პროგრამების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' გავასუფთაოთ კომბობოქსი
            CBDaf.Items.Clear()

            ' დავამატოთ პირველი ელემენტი - მინიშნება
            CBDaf.Items.Add("- აირჩიეთ დაფინანსების პროგრამა -")

            ' დავამატოთ დაფინანსების პროგრამები
            For Each row In rows
                If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) Then
                    CBDaf.Items.Add(row(0).ToString().Trim())
                End If
            Next

            ' ავირჩიოთ პირველი ელემენტი (მინიშნება)
            CBDaf.SelectedIndex = 0

        Catch ex As Exception
            MessageBox.Show($"შეცდომა დაფინანსების პროგრამების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' TCost-ში კლავიშზე დაჭერის ივენთი - მხოლოდ ციფრების და წერტილის/მძიმის დაშვება
    ''' </summary>
    Private Sub TCost_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TCost.KeyPress
        ' აქცეპტირებული სიმბოლოები: ციფრები, წაშლის კლავიში, წერტილი, მძიმე
        If Not (Char.IsDigit(e.KeyChar) OrElse e.KeyChar = ControlChars.Back OrElse e.KeyChar = "." OrElse e.KeyChar = ",") Then
            ' დანარჩენი სიმბოლოების ბლოკირება
            e.Handled = True
            Return
        End If

        ' შევამოწმოთ თუ ეს წერტილი ან მძიმეა
        If e.KeyChar = "." OrElse e.KeyChar = "," Then
            ' წერტილი/მძიმე არ უნდა იყოს პირველი სიმბოლო 
            If TCost.Text.Length = 0 Then
                e.Handled = True
                Return
            End If

            ' წერტილი/მძიმე არ უნდა გამეორდეს
            If TCost.Text.Contains(".") OrElse TCost.Text.Contains(",") Then
                e.Handled = True
                Return
            End If
        End If
    End Sub

    ''' <summary>
    ''' TCost-ში ტექსტის შეცვლის ივენთი - დამატებითი ვალიდაცია
    ''' </summary>
    Private Sub TCost_TextChanged(sender As Object, e As EventArgs) Handles TCost.TextChanged
        ' ტექსტის ვალიდაცია და კორექტირება
        Dim text As String = TCost.Text

        ' თუ ტექსტი ცარიელია, დავაყენოთ 0
        If String.IsNullOrEmpty(text) Then
            TCost.Text = "0"
            Return
        End If

        ' შევამოწმოთ რომ ტექსტი რიცხვის ფორმატშია
        Dim isValidNumber As Boolean = True

        ' ვეცადოთ დავაკონვერტიროთ Double-ად
        Dim value As Double
        If Not Double.TryParse(text.Replace(",", "."), value) Then
            isValidNumber = False
        End If

        ' თუ არ არის ვალიდური რიცხვი, დავაბრუნოთ 0 ან წინა ვალიდური მნიშვნელობა
        If Not isValidNumber Then
            TCost.Text = "0"
        End If
    End Sub

    ''' <summary>
    ''' TCost-დან გასვლის ივენთი - რიცხვის ფორმატირება
    ''' </summary>
    Private Sub TCost_Leave(sender As Object, e As EventArgs) Handles TCost.Leave
        Try
            ' ტექსტის ფორმატირება როგორც რიცხვის
            Dim text As String = TCost.Text.Replace(",", ".")

            ' თუ ტექსტი ცარიელია ან არ არის რიცხვი, დავაყენოთ 0
            Dim value As Double
            If String.IsNullOrEmpty(text) OrElse Not Double.TryParse(text, value) Then
                TCost.Text = "0"
                Return
            End If

            ' ფორმატირება ორი ციფრით მძიმის შემდეგ
            TCost.Text = String.Format("{0:0.00}", value)
        Catch ex As Exception
            ' შეცდომის შემთხვევაში დავაყენოთ 0
            TCost.Text = "0"
        End Try
    End Sub
    ''' <summary>
    ''' სივრცეების ღილაკების Click ივენთების მიბმა
    ''' </summary>
    Private Sub AttachSpaceButtonHandlers()
        ' ყველა BTNS ღილაკის მოძებნა და ივენთის მიბმა
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            If btn.Name.StartsWith("BTNS") Then
                AddHandler btn.Click, AddressOf SpaceButton_Click
            End If
        Next
    End Sub

    ''' <summary>
    ''' რადიობუტონების Click ივენთების მიბმა
    ''' </summary>
    Private Sub AttachRadioButtonHandlers()
        ' ყველა RB* რადიობუტონის მოძებნა და ივენთის მიბმა
        For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
            If rb.Name.StartsWith("RB") Then
                AddHandler rb.CheckedChanged, AddressOf StatusRadioButton_CheckedChanged
            End If
        Next
    End Sub

    ''' <summary>
    ''' სივრცეების ღილაკების საწყისი მდგომარეობის დაყენება
    ''' ამოწმებს რომელი სივრცეებია დაკავებული, თავისუფალი, ან ჯგუფური სესიით
    ''' </summary>
    Private Sub InitializeSpaceButtons()
        Try
            ' თარიღისა და დროის აღება ფორმიდან
            Dim selectedDate As DateTime = GetSelectedDateTime()

            ' სივრცეების დატვირთულობის ინფორმაციის წამოღება Google Sheets-დან
            Dim occupiedSpaces As Dictionary(Of String, Boolean) = GetOccupiedSpaces(selectedDate)
            Dim groupSpaces As Dictionary(Of String, Boolean) = GetGroupSpaces(selectedDate)

            ' ინფორმაცია დაკავებულ სივრცეებზე - ახალი ლექსიკონი
            Dim spaceDetailsDict As New Dictionary(Of String, IList(Of Object))

            ' მოვიპოვოთ ინფორმაცია ყველა სივრცეზე
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:K")

            ' დროის შუალედი
            Dim startTime As DateTime = selectedDate.AddMinutes(-30)
            Dim endTime As DateTime = selectedDate.AddMinutes(90)

            ' ავაგოთ სივრცის დეტალების ლექსიკონი
            If scheduleData IsNot Nothing Then
                For Each row In scheduleData
                    If row.Count >= 11 Then
                        Try
                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = If(row.Count > 5, row(5).ToString(), "")
                            Dim sessionDateTime As DateTime

                            ' სივრცის სახელი K სვეტიდან (ინდექსი 10)
                            Dim spaceName As String = row(10).ToString().Trim()

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None,
                                            sessionDateTime) OrElse
                            DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' შევამოწმოთ ემთხვევა თუ არა საჭირო დროის შუალედს
                                If sessionDateTime >= startTime AndAlso sessionDateTime <= endTime Then
                                    ' დავიმახსოვროთ დეტალები ამ სივრცისთვის
                                    spaceDetailsDict(spaceName) = row
                                End If
                            End If
                        Catch ex As Exception
                            ' გავაგრძელოთ შემდეგ ჩანაწერზე
                            Continue For
                        End Try
                    End If
                Next
            End If

            ' მიმდინარე არჩეული ღილაკი (თუ ასეთი არსებობს)
            Dim selectedButton As Button = Nothing

            ' ყველა BTNS ღილაკის მოძებნა
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") Then
                    Dim spaceName As String = btn.Text.Trim()

                    ' თუ ღილაკი უკვე ცისფერია (არჩეულია), დავიმახსოვროთ
                    If btn.BackColor = Color.FromArgb(173, 216, 230) Then
                        selectedButton = btn
                    End If

                    If occupiedSpaces.ContainsKey(spaceName) AndAlso occupiedSpaces(spaceName) Then
                        ' დაკავებული სივრცე - ღია წითელი
                        btn.BackColor = Color.FromArgb(255, 200, 200)
                    ElseIf groupSpaces.ContainsKey(spaceName) AndAlso groupSpaces(spaceName) Then
                        ' ჯგუფური სესიის სივრცე - ღია ყვითელი
                        btn.BackColor = Color.FromArgb(255, 255, 200)
                    Else
                        ' თავისუფალი სივრცე - ღია მწვანე
                        btn.BackColor = Color.FromArgb(200, 255, 200)
                    End If
                End If
            Next

            ' საწყისი მდგომარეობა ლეიბლისთვის - თუ ღილაკი არ არის არჩეული
            If selectedButton Is Nothing Then
                LMsgSpace.Text = "მოცემულ დროს სივრცე თავისუფალია"
                LMsgSpace.ForeColor = Color.DarkGreen
            Else
                ' აღვადგინოთ არჩეული ღილაკი
                selectedButton.BackColor = Color.FromArgb(173, 216, 230) ' ცისფერი

                ' შესაბამისი ინფორმაციის ჩვენება
                UpdateSpaceInfoLabel(selectedButton.Text, spaceDetailsDict)
            End If

        Catch ex As Exception
            Debug.WriteLine($"InitializeSpaceButtons: შეცდომა - {ex.Message}")
            MessageBox.Show($"სივრცეების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' დამხმარე მეთოდი სივრცის ინფორმაციის ლეიბლის განახლებისთვის
    ''' </summary>
    ''' <param name="spaceName"></param>
    ''' <param name="spaceDetails"></param>
    ''' <remarks>განახლება - 2023-10-01</remarks>
    Private Sub UpdateSpaceInfoLabel(spaceName As String, spaceDetails As Dictionary(Of String, IList(Of Object)))
        Try
            ' სივრცესთან დაკავშირებული ინფორმაცია
            If spaceDetails.ContainsKey(spaceName) Then
                ' დაკავებული სივრცე - დეტალური ინფორმაცია
                Dim sessionDetails As IList(Of Object) = spaceDetails(spaceName)

                ' ბენეფიციარი, თერაპევტი და თერაპიის ტიპი
                Dim beneficiaryName As String = If(sessionDetails.Count > 3, sessionDetails(3).ToString(), "")
                Dim beneficiarySurname As String = If(sessionDetails.Count > 4, sessionDetails(4).ToString(), "")
                Dim therapistName As String = If(sessionDetails.Count > 8, sessionDetails(8).ToString(), "")
                Dim therapyType As String = If(sessionDetails.Count > 9, sessionDetails(9).ToString(), "")

                ' ჯგუფური სეანსის ინფორმაცია
                Dim isGroup As Boolean = False
                If sessionDetails.Count > 6 AndAlso Not String.IsNullOrEmpty(sessionDetails(6).ToString()) Then
                    Boolean.TryParse(sessionDetails(6).ToString(), isGroup)
                End If

                ' ტექსტის მომზადება
                Dim infoText As New System.Text.StringBuilder()

                If isGroup Then
                    infoText.AppendLine($"სივრცე '{spaceName}' დაკავებულია ჯგუფური სეანსით:")
                    LMsgSpace.ForeColor = Color.DarkOrange
                Else
                    infoText.AppendLine($"სივრცე '{spaceName}' დაკავებულია:")
                    LMsgSpace.ForeColor = Color.DarkRed
                End If

                infoText.AppendLine($"ბენეფიციარი: {beneficiaryName} {beneficiarySurname}")
                infoText.AppendLine($"თერაპევტი: {therapistName}")
                infoText.Append($"თერაპია: {therapyType}")

                LMsgSpace.Text = infoText.ToString()
            Else
                ' თავისუფალი სივრცე
                LMsgSpace.Text = $"სივრცე '{spaceName}' მოცემულ დროს თავისუფალია"
                LMsgSpace.ForeColor = Color.DarkGreen
            End If
        Catch ex As Exception
            Debug.WriteLine($"UpdateSpaceInfoLabel: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში ვაჩვენოთ უბრალო ტექსტი
            LMsgSpace.Text = $"არჩეულია სივრცე: {spaceName}"
            LMsgSpace.ForeColor = Color.Blue
        End Try
    End Sub

    ''' <summary>
    ''' აბრუნებს არჩეულ თარიღსა და დროს DateTime ობიექტად
    ''' </summary>
    Private Function GetSelectedDateTime() As DateTime
        Dim selectedDate As DateTime = DTP1.Value.Date

        ' საათისა და წუთის აღება ტექსტბოქსებიდან
        Dim hour As Integer = 0
        Dim minute As Integer = 0

        If Not Integer.TryParse(THour.Text, hour) Then hour = 12
        If Not Integer.TryParse(TMin.Text, minute) Then minute = 0

        ' ვალიდაცია
        hour = Math.Max(0, Math.Min(23, hour))
        minute = Math.Max(0, Math.Min(59, minute))

        ' საბოლოო თარიღის და დროის ობიექტის შექმნა
        Return New DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, hour, minute, 0)
    End Function

    ''' <summary>
    ''' ამოწმებს Google Sheets-დან რომელი სივრცეებია დაკავებული მოცემული თარიღისა და დროისთვის
    ''' </summary>
    Private Function GetOccupiedSpaces(selectedDateTime As DateTime) As Dictionary(Of String, Boolean)
        Dim occupiedSpaces As New Dictionary(Of String, Boolean)

        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                Debug.WriteLine("GetOccupiedSpaces: dataService არ არის ინიციალიზებული")
                Return occupiedSpaces
            End If

            ' დროის შუალედის განსაზღვრა (30 წუთიანი შუალედი)
            Dim startTime As DateTime = selectedDateTime.AddMinutes(-30)
            Dim endTime As DateTime = selectedDateTime.AddMinutes(90) ' სესიებისთვის ვამოწმებთ დაგეგმილ დროს +/- 30 წთ

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:K")

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 11 Then ' K სვეტი არის ინდექსი 10
                        Try
                            ' სივრცის სახელი K სვეტიდან
                            Dim spaceName As String = row(10).ToString().Trim()

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = row(5).ToString()
                            Dim sessionDateTime As DateTime

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                     System.Globalization.CultureInfo.InvariantCulture,
                                                     System.Globalization.DateTimeStyles.None,
                                                     sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' შევამოწმოთ ემთხვევა თუ არა საჭირო დროის შუალედს
                                If sessionDateTime >= startTime AndAlso sessionDateTime <= endTime Then
                                    ' დაკავებული სივრცე
                                    occupiedSpaces(spaceName) = True
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Continue For
                        End Try
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetOccupiedSpaces: შეცდომა - {ex.Message}")
        End Try

        Return occupiedSpaces
    End Function

    ''' <summary>
    ''' ამოწმებს Google Sheets-დან რომელი სივრცეებშია ჯგუფური სესიები მოცემული თარიღისა და დროისთვის
    ''' </summary>
    Private Function GetGroupSpaces(selectedDateTime As DateTime) As Dictionary(Of String, Boolean)
        Dim groupSpaces As New Dictionary(Of String, Boolean)

        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                Debug.WriteLine("GetGroupSpaces: dataService არ არის ინიციალიზებული")
                Return groupSpaces
            End If

            ' დროის შუალედის განსაზღვრა (30 წუთიანი შუალედი)
            Dim startTime As DateTime = selectedDateTime.AddMinutes(-30)
            Dim endTime As DateTime = selectedDateTime.AddMinutes(90)

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება (სივრცე და IsGroup)
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:K")

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 11 Then
                        Try
                            ' სივრცის სახელი K სვეტიდან
                            Dim spaceName As String = row(10).ToString().Trim()

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = row(5).ToString()
                            Dim sessionDateTime As DateTime

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                     System.Globalization.CultureInfo.InvariantCulture,
                                                     System.Globalization.DateTimeStyles.None,
                                                     sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' შევამოწმოთ ემთხვევა თუ არა საჭირო დროის შუალედს 
                                ' და რომ ეს ჯგუფური სესიაა (G სვეტი - ინდექსი 6)
                                If sessionDateTime >= startTime AndAlso sessionDateTime <= endTime AndAlso
                                   row.Count > 6 AndAlso row(6).ToString().ToLower() = "true" Then
                                    ' ჯგუფური სესიის სივრცე
                                    groupSpaces(spaceName) = True
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Continue For
                        End Try
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetGroupSpaces: შეცდომა - {ex.Message}")
        End Try

        Return groupSpaces
    End Function

    ''' <summary>
    ''' სივრცის ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub SpaceButton_Click(sender As Object, e As EventArgs)
        ' აღვადგინოთ ყველა ღილაკის ფერი
        InitializeSpaceButtons()

        ' დავაყენოთ არჩეული ღილაკის ფერი ცისფერზე
        Dim clickedButton As Button = DirectCast(sender, Button)
        clickedButton.BackColor = Color.FromArgb(173, 216, 230) ' ცისფერი

        ' შევინახოთ არჩეული სივრცის სახელი
        Dim selectedSpace As String = clickedButton.Text

        ' განვაახლოთ სტატუსის ლეიბლი
        LMsgSpace.Text = $"არჩეულია სივრცე: {selectedSpace}"
        LMsgSpace.ForeColor = Color.Blue
    End Sub

    ''' <summary>
    ''' DTP1-ის CloseUp ივენთი - კალენდრის დახურვისას გააქტიურდება
    ''' </summary>
    Private Sub DTP1_CloseUp(sender As Object, e As EventArgs)
        ' დროის ცვლილების დამმუშავებლის გამოძახება
        DateTimeChanged(sender, e)
    End Sub

    ''' <summary>
    ''' დროის ან თარიღის ცვლილების ივენთი - გაფართოებული ბენეფიციარის შემოწმებით
    ''' დამატებულია რეკურსიის თავიდან აცილების მექანიზმი
    ''' </summary>
    Private Sub DateTimeChanged(sender As Object, e As EventArgs)
        ' რეკურსიისგან დაცვა
        If isUpdating Then
            Return
        End If

        isUpdating = True

        Try
            ' განვაახლოთ სივრცეების სტატუსები
            InitializeSpaceButtons()

            ' განვაახლოთ შესრულების სტატუსის ხილვადობა
            UpdateStatusVisibility()

            ' განვაახლოთ ბენეფიციარის მდგომარეობა
            CheckBeneficiaryAvailability()

            ' განვაახლოთ თერაპევტის მდგომარეობა
            If CBPer.SelectedIndex > 0 Then
                CheckTherapistAvailability()
            End If
        Finally
            isUpdating = False
        End Try
    End Sub

    ''' <summary>
    ''' განაახლებს შესრულების სტატუსის კონტროლების ხილვადობას თარიღის მიხედვით
    ''' თავიდან არცერთი რადიობუტონი არ არის მონიშნული
    ''' </summary>
    Private Sub UpdateStatusVisibility()
        Try
            ' არჩეული თარიღი და დრო
            Dim selectedDateTime As DateTime = GetSelectedDateTime()

            ' მიმდინარე დრო
            Dim currentDateTime As DateTime = DateTime.Now

            ' ყველა რადიობუტონი
            Dim radioButtons As New List(Of RadioButton)
            For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                If rb.Name.StartsWith("RB") Then
                    radioButtons.Add(rb)
                End If
            Next

            ' მოვხსნათ ყველა მონიშვნა
            For Each rb In radioButtons
                rb.Checked = False
            Next

            ' ვამოწმებთ არის თუ არა არჩეული დრო მიმდინარე მომენტზე მეტი (სამომავლო)
            If selectedDateTime > currentDateTime Then
                ' სამომავლო თარიღი - დავმალოთ რადიობუტონები და გამოვაჩინოთ LPlan
                For Each rb In radioButtons
                    rb.Visible = False
                Next

                LPlan.Visible = True

                ' შეტყობინების განახლება
                LWarning.Text = "სეანსი იგეგმება მომავალში, სტატუსი იქნება 'დაგეგმილი'"
                LWarning.ForeColor = Color.DarkGreen
            Else
                ' წარსული თარიღი - გამოვაჩინოთ რადიობუტონები და დავმალოთ LPlan
                For Each rb In radioButtons
                    rb.Visible = True
                Next

                LPlan.Visible = False

                ' შეტყობინების განახლება
                LWarning.Text = "სეანსის თარიღი არის წარსულში, აირჩიეთ შესრულების სტატუსი"
                LWarning.ForeColor = Color.DarkRed

                ' არ ვანიშნინებთ არცერთ რადიობუტონს ავტომატურად
            End If

        Catch ex As Exception
            Debug.WriteLine($"UpdateStatusVisibility: შეცდომა - {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' ბენეფიციარის მდგომარეობის შემოწმება არჩეული თარიღისა და დროისთვის
    ''' </summary>
    Private Sub CheckBeneficiaryAvailability()
        Try
            ' შევამოწმოთ გვაქვს თუ არა არჩეული ბენეფიციარი
            If CBBeneName.SelectedIndex <= 0 OrElse String.IsNullOrEmpty(CBBeneSurname.Text) Then
                ' თუ ბენეფიციარი არ არის არჩეული ან არჩეულია "-აირჩიეთ-", გავასუფთაოთ ლეიბლი და ბექგრაუნდი
                LMsgBene.Text = "აირჩიეთ ბენეფიციარი"
                LMsgBene.ForeColor = Color.Black
                CBBeneName.BackColor = SystemColors.Window
                CBBeneSurname.BackColor = SystemColors.Window
                Return
            End If

            ' არჩეული თარიღი და დრო
            Dim selectedDateTime As DateTime = GetSelectedDateTime()

            ' დროის შუალედის განსაზღვრა (30 წუთიანი შუალედი)
            Dim startTime As DateTime = selectedDateTime.AddMinutes(-30)
            Dim endTime As DateTime = selectedDateTime.AddMinutes(90)

            ' არჩეული ბენეფიციარი
            Dim beneficiaryName As String = CBBeneName.Text.Trim()
            Dim beneficiarySurname As String = CBBeneSurname.Text.Trim()

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:K")

            ' ბენეფიციარის არსებული სესიების მოძებნა
            Dim beneficiarySession As IList(Of Object) = Nothing

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 11 Then ' K სვეტამდე (სივრცემდე)
                        Try
                            ' ბენეფიციარის სახელი (D სვეტი - ინდექსი 3) და გვარი (E სვეტი - ინდექსი 4)
                            Dim rowBeneName As String = If(row.Count > 3, row(3).ToString().Trim(), "")
                            Dim rowBeneSurname As String = If(row.Count > 4, row(4).ToString().Trim(), "")

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = If(row.Count > 5, row(5).ToString(), "")
                            Dim sessionDateTime As DateTime

                            ' ვამოწმებთ თუ ეს იგივე ბენეფიციარია
                            If String.Equals(rowBeneName, beneficiaryName, StringComparison.OrdinalIgnoreCase) AndAlso
                           String.Equals(rowBeneSurname, beneficiarySurname, StringComparison.OrdinalIgnoreCase) Then

                                ' პარსინგის მცდელობა
                                If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                     System.Globalization.CultureInfo.InvariantCulture,
                                                     System.Globalization.DateTimeStyles.None,
                                                     sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                    ' შევამოწმოთ ემთხვევა თუ არა საჭირო დროის შუალედს
                                    If sessionDateTime >= startTime AndAlso sessionDateTime <= endTime Then
                                        ' ნაპოვნია ბენეფიციარის სესია მოცემულ დროის შუალედში
                                        beneficiarySession = row
                                        Exit For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Continue For
                        End Try
                    End If
                Next
            End If

            ' ბენეფიციარის მდგომარეობის ვიზუალური ასახვა
            If beneficiarySession IsNot Nothing Then
                ' ბენეფიციარი დაკავებულია - წითელი ინდიკატორი
                CBBeneName.BackColor = Color.FromArgb(255, 200, 200) ' ღია წითელი
                CBBeneSurname.BackColor = Color.FromArgb(255, 200, 200) ' ღია წითელი

                ' დეტალური ინფორმაცია თერაპევტზე, თერაპიაზე და სივრცეზე
                Dim therapistName As String = If(beneficiarySession.Count > 8, beneficiarySession(8).ToString(), "")
                Dim therapyType As String = If(beneficiarySession.Count > 9, beneficiarySession(9).ToString(), "")
                Dim spaceName As String = If(beneficiarySession.Count > 10, beneficiarySession(10).ToString(), "")

                ' მრავალხაზიანი ტექსტი ლეიბლში
                Dim infoText As New System.Text.StringBuilder()
                infoText.AppendLine("მოცემულ დროს ბენეფიციარი დაკავებულია:")
                infoText.AppendLine($"თერაპევტი: {therapistName}")
                infoText.AppendLine($"თერაპია: {therapyType}")
                infoText.Append($"სივრცე: {spaceName}")

                LMsgBene.Text = infoText.ToString()
                LMsgBene.ForeColor = Color.DarkRed
            Else
                ' ბენეფიციარი თავისუფალია - მწვანე ინდიკატორი
                CBBeneName.BackColor = Color.FromArgb(200, 255, 200) ' ღია მწვანე
                CBBeneSurname.BackColor = Color.FromArgb(200, 255, 200) ' ღია მწვანე

                LMsgBene.Text = "მოცემულ დროს ბენეფიციარი თავისუფალია"
                LMsgBene.ForeColor = Color.DarkGreen
            End If

        Catch ex As Exception
            Debug.WriteLine($"CheckBeneficiaryAvailability: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში ვაჩვენოთ ნეიტრალური სტატუსი
            LMsgBene.Text = "ბენეფიციარის მდგომარეობის შემოწმება ვერ მოხერხდა"
            LMsgBene.ForeColor = Color.Gray
            CBBeneName.BackColor = SystemColors.Window
            CBBeneSurname.BackColor = SystemColors.Window
        End Try
    End Sub

    ''' <summary>
    ''' თერაპევტის მდგომარეობის შემოწმება არჩეული თარიღისა და დროისთვის
    ''' </summary>
    Private Sub CheckTherapistAvailability()
        Try
            ' შევამოწმოთ გვაქვს თუ არა არჩეული თერაპევტი
            If CBPer.SelectedIndex <= 0 Then
                ' თუ თერაპევტი არ არის არჩეული ან არჩეულია "-აირჩიეთ-", გავასუფთაოთ ლეიბლი და ბექგრაუნდი
                LMsgPer.Text = "აირჩიეთ თერაპევტი"
                LMsgPer.ForeColor = Color.Black
                CBPer.BackColor = SystemColors.Window
                Return
            End If

            ' არჩეული თარიღი და დრო
            Dim selectedDateTime As DateTime = GetSelectedDateTime() ' იგივე მეთოდი რასაც ბენეფიციარების შემთხვევაში ვიყენებთ

            ' დროის შუალედის განსაზღვრა (30 წუთიანი შუალედი)
            Dim startTime As DateTime = selectedDateTime.AddMinutes(-30)
            Dim endTime As DateTime = selectedDateTime.AddMinutes(90)

            ' არჩეული თერაპევტი
            Dim therapistName As String = CBPer.Text.Trim()

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:K")

            ' თერაპევტის არსებული სესიების მოძებნა
            Dim therapistSession As IList(Of Object) = Nothing
            Dim isGroupSession As Boolean = False

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 11 Then ' K სვეტამდე (სივრცემდე)
                        Try
                            ' თერაპევტის სახელი (I სვეტი - ინდექსი 8)
                            Dim rowTherapistName As String = If(row.Count > 8, row(8).ToString().Trim(), "")

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = If(row.Count > 5, row(5).ToString(), "")
                            Dim sessionDateTime As DateTime

                            ' ვამოწმებთ თუ ეს იგივე თერაპევტია
                            If String.Equals(rowTherapistName, therapistName, StringComparison.OrdinalIgnoreCase) Then

                                ' პარსინგის მცდელობა
                                If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                 System.Globalization.CultureInfo.InvariantCulture,
                                                 System.Globalization.DateTimeStyles.None,
                                                 sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                    ' შევამოწმოთ ემთხვევა თუ არა საჭირო დროის შუალედს
                                    If sessionDateTime >= startTime AndAlso sessionDateTime <= endTime Then
                                        ' ნაპოვნია თერაპევტის სესია მოცემულ დროის შუალედში
                                        therapistSession = row

                                        ' შევამოწმოთ არის თუ არა ჯგუფური სესია (G სვეტი - ინდექსი 6)
                                        If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6).ToString()) Then
                                            isGroupSession = Boolean.Parse(row(6).ToString())
                                        End If

                                        Exit For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Continue For
                        End Try
                    End If
                Next
            End If

            ' თერაპევტის მდგომარეობის ვიზუალური ასახვა
            If therapistSession IsNot Nothing Then
                If isGroupSession Then
                    ' ჯგუფური სეანსი - ყვითელი ინდიკატორი
                    CBPer.BackColor = Color.FromArgb(255, 255, 200) ' ღია ყვითელი

                    ' დეტალური ინფორმაცია ბენეფიციარზე, თერაპიაზე და სივრცეზე
                    Dim beneficiaryName As String = If(therapistSession.Count > 3, therapistSession(3).ToString(), "")
                    Dim beneficiarySurname As String = If(therapistSession.Count > 4, therapistSession(4).ToString(), "")
                    Dim therapyType As String = If(therapistSession.Count > 9, therapistSession(9).ToString(), "")
                    Dim spaceName As String = If(therapistSession.Count > 10, therapistSession(10).ToString(), "")

                    ' მრავალხაზიანი ტექსტი ლეიბლში
                    Dim infoText As New System.Text.StringBuilder()
                    infoText.AppendLine("მოცემულ დროს თერაპევტს აქვს ჯგუფური სეანსი:")
                    infoText.AppendLine($"{beneficiaryName} {beneficiarySurname}")
                    infoText.AppendLine($"თერაპია: {therapyType}")
                    infoText.Append($"სივრცე: {spaceName}")

                    LMsgPer.Text = infoText.ToString()
                    LMsgPer.ForeColor = Color.DarkOrange
                Else
                    ' ჩვეულებრივი სეანსი - წითელი ინდიკატორი
                    CBPer.BackColor = Color.FromArgb(255, 200, 200) ' ღია წითელი

                    ' დეტალური ინფორმაცია ბენეფიციარზე, თერაპიაზე და სივრცეზე
                    Dim beneficiaryName As String = If(therapistSession.Count > 3, therapistSession(3).ToString(), "")
                    Dim beneficiarySurname As String = If(therapistSession.Count > 4, therapistSession(4).ToString(), "")
                    Dim therapyType As String = If(therapistSession.Count > 9, therapistSession(9).ToString(), "")
                    Dim spaceName As String = If(therapistSession.Count > 10, therapistSession(10).ToString(), "")

                    ' მრავალხაზიანი ტექსტი ლეიბლში
                    Dim infoText As New System.Text.StringBuilder()
                    infoText.AppendLine("მოცემულ დროს თერაპევტი დაკავებულია:")
                    infoText.AppendLine($"{beneficiaryName} {beneficiarySurname}")
                    infoText.AppendLine($"თერაპია: {therapyType}")
                    infoText.Append($"სივრცე: {spaceName}")

                    LMsgPer.Text = infoText.ToString()
                    LMsgPer.ForeColor = Color.DarkRed
                End If
            Else
                ' თერაპევტი თავისუფალია - მწვანე ინდიკატორი
                CBPer.BackColor = Color.FromArgb(200, 255, 200) ' ღია მწვანე

                LMsgPer.Text = "მოცემულ დროს თერაპევტი თავისუფალია"
                LMsgPer.ForeColor = Color.DarkGreen
            End If

        Catch ex As Exception
            Debug.WriteLine($"CheckTherapistAvailability: შეცდომა - {ex.Message}")

            ' შეცდომის შემთხვევაში ვაჩვენოთ ნეიტრალური სტატუსი
            LMsgPer.Text = "თერაპევტის მდგომარეობის შემოწმება ვერ მოხერხდა"
            LMsgPer.ForeColor = Color.Gray
            CBPer.BackColor = SystemColors.Window
        End Try
    End Sub

    ''' <summary>
    ''' რადიობუტონის არჩევის ივენთი
    ''' </summary>
    Private Sub StatusRadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = DirectCast(sender, RadioButton)

        ' მხოლოდ არჩეულ რადიობუტონს ვამუშავებთ
        If rb.Checked Then
            ' შეტყობინების განახლება
            LWarning.Text = $"არჩეული სტატუსი: {rb.Text}"
            LWarning.ForeColor = Color.Blue
        End If
    End Sub

    ''' <summary>
    ''' ValueChanged ივენთი აღარ გამოიყენება, მაგრამ რჩება დიზაინერში მიბმული
    ''' შევქმნით ცარიელ მეთოდს, რომელიც არაფერს აკეთებს
    ''' </summary>
    Private Sub DTP1_ValueChanged(sender As Object, e As EventArgs) Handles DTP1.ValueChanged
        ' არაფერს აკეთებს, რეალური ფუნქცონალი CloseUp-შია
    End Sub
End Class