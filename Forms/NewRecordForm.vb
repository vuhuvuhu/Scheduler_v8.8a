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

            ' ბენეფიციარის სახელების ჩატვირთვა
            LoadBeneNames()

            ' CBBeneSurname თავიდან დაბლოკილია
            CBBeneSurname.Enabled = False

            ' DateTimePicker (DTP1) კონფიგურაცია ქართული ფორმატით
            ConfigureDateTimePicker()

            ' საათის, წუთის და ხანგძლივობის ლეიბლების ინიციალიზაცია
            InitializeTimeAndDurationLabels()

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ფორმის ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

        ' ვიზუალური სტილი - სწორი მნიშვნელობები Label კონტროლებისთვის
        THour.TextAlign = HorizontalAlignment.Center
        TMin.TextAlign = HorizontalAlignment.Center
        TDur.TextAlign = HorizontalAlignment.Center

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
    ''' ბენეფიციარის სახელების ჩატვირთვა DB-Bene ფურცლიდან
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

            ' უნიკალური სახელების სორტირება ანბანის მიხედვით და დამატება CBBeneName-ში
            Dim sortedNamesList As List(Of String) = uniqueNames.OrderBy(Function(name) name).ToList()

            For i As Integer = 0 To sortedNamesList.Count - 1
                CBBeneName.Items.Add(sortedNamesList(i))
            Next

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ბენეფიციარის სახელების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneName-ში სახელის არჩევისას ივენთი
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            ' გავასუფთაოთ CBBeneSurname-ის Items კოლექცია
            CBBeneSurname.Items.Clear()

            ' თუ ინდექსი ვალიდურია, ჩავტვირთოთ შესაბამისი გვარები
            If CBBeneName.SelectedIndex >= 0 Then
                Dim selectedName As String = CBBeneName.SelectedItem.ToString()
                LoadBeneSurnames(selectedName)

                ' გავააქტიუროთ გვარების ჩამონათვალი
                CBBeneSurname.Enabled = True
            Else
                ' ინდექსი არავალიდურია, გამოვრთოთ გვარების ჩამონათვალი
                CBBeneSurname.Enabled = False
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
    ''' ბენეფიციარის გვარების ჩატვირთვა არჩეული სახელის შესაბამისად
    ''' </summary>
    ''' <param name="selectedName">არჩეული სახელი</param>
    Private Sub LoadBeneSurnames(selectedName As String)
        Try
            ' გავასუფთაოთ გვარების ჩამონათვალი
            CBBeneSurname.Items.Clear()

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

            ' თუ მხოლოდ ერთი გვარია, ავტომატურად ავირჩიოთ
            If sortedSurnamesList.Count = 1 Then
                CBBeneSurname.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ბენეფიციარის გვარების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
End Class