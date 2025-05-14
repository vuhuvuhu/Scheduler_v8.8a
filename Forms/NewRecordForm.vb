' ===========================================
' 📄 Forms/NewRecordForm.vb
' -------------------------------------------
' ახალი ჩანაწერის ფორმა - გამოიყენება ახალი სესიის, დაბადების დღის ან დავალების დასამატებლად
' ასევე არსებული ჩანაწერების რედაქტირებისთვის
' ===========================================
Imports System.Globalization
Imports System.Text
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
    ''' ფორმის ჩატვირთვის ივენთი - დამატებულია საათის, წუთის და ხანგძლივობის ელემენტების ინიციალიზაცია
    ''' </summary>
    Private Sub NewRecordForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' ფორმის გაფორმება რეჟიმის მიხედვით
            If _isAddMode Then
                ' ახალი ჩანაწერის დამატების რეჟიმი - ღია მწვანე ფონი
                Me.BackColor = Color.FromArgb(200, 255, 200) ' ღია მწვანე
                Me.Text = $"ახალი {recordType} - დამატება"

                ' ჩანაწერის ნომრის (ID) მოპოვება
                Dim maxId As Integer = GetMaxRecordId()
                LN.Text = (maxId + 1).ToString()

                ' BtnAdd ტექსტის დაყენება
                BtnAdd.Text = "დამატება"
            Else
                ' რედაქტირების რეჟიმი - ღია ყვითელი ფონი
                Me.BackColor = Color.FromArgb(255, 255, 200) ' ღია ყვითელი
                Me.Text = $"{recordType} - რედაქტირება (ID: {_editRecordId})"

                ' ID-ის დაყენება
                LN.Text = _editRecordId.ToString()

                ' BtnAdd ტექსტის დაყენება
                BtnAdd.Text = "შენახვა"
            End If

            ' მიმდინარე თარიღის და დროის დაყენება
            LNow.Text = DateTime.Now.ToString("dd.MM.yyyy HH:mm")

            ' ავტორიზირებული მომხმარებლის ელ.ფოსტის გამოყენება
            LAutor.Text = _userEmail

            ' TCost ცარიელი ნაცვლად "0"-ისა
            TCost.Text = ""

            ' LWarning-ის ინიციალიზაცია
            LWarning.Text = "გთხოვთ შეავსოთ ყველა აუცილებელი ველი"
            LWarning.ForeColor = Color.Black

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

            ' რედაქტირების რეჟიმში ჩავტვირთოთ არსებული ჩანაწერის მონაცემები
            If Not _isAddMode Then
                ' ცოტა დავიცადოთ, რომ კომბობოქსები მოასწრონ ჩატვირთვა
                Application.DoEvents()
                System.Threading.Thread.Sleep(500)

                ' ჩავტვირთოთ მონაცემები
                LoadRecordData()
            End If

        Catch ex As Exception
            MessageBox.Show($"შეცდომა ფორმის ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"NewRecordForm_Load შეცდომა: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' არსებული ჩანაწერის მონაცემების ჩატვირთვა ID-ის მიხედვით
    ''' </summary>
    Private Sub LoadRecordData()
        Try
            Debug.WriteLine($"LoadRecordData: ვტვირთავ ჩანაწერს ID={_editRecordId}, ტიპი={recordType}")

            ' ჩანაწერის ტიპის მიხედვით განვსაზღვროთ ფურცლის სახელი და დიაპაზონი
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

            ' ჩავტვირთოთ ყველა ჩანაწერი ფურცლიდან
            Dim data = dataService.GetData($"{sheetName}!A2:O")

            If data Is Nothing OrElse data.Count = 0 Then
                MessageBox.Show($"ჩანაწერი ID={_editRecordId} ვერ მოიძებნა", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ვეძებთ ჩანაწერს ID-ის მიხედვით
            Dim recordRow As IList(Of Object) = Nothing
            For Each row In data
                If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString()) Then
                    Dim id As Integer
                    If Integer.TryParse(row(0).ToString(), id) AndAlso id = _editRecordId Then
                        recordRow = row
                        Debug.WriteLine($"LoadRecordData: ნაპოვნია ჩანაწერი ID={_editRecordId}")
                        Exit For
                    End If
                End If
            Next

            ' თუ ჩანაწერი ვერ ვიპოვეთ
            If recordRow Is Nothing Then
                MessageBox.Show($"ჩანაწერი ID={_editRecordId} ვერ მოიძებნა", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ჩანაწერის ტიპის მიხედვით შევავსოთ შესაბამისი ველები
            If recordType.ToLower() = "სესია" Then
                LoadSessionData(recordRow)
            ElseIf recordType.ToLower() = "დაბადების დღე" Then
                ' აქ შეგვიძლია დავამატოთ დაბადების დღის მონაცემების ჩატვირთვის ლოგიკა
            ElseIf recordType.ToLower() = "დავალება" Then
                ' აქ შეგვიძლია დავამატოთ დავალების მონაცემების ჩატვირთვის ლოგიკა
            End If

        Catch ex As Exception
            Debug.WriteLine($"LoadRecordData: შეცდომა - {ex.Message}")
            MessageBox.Show($"ჩანაწერის მონაცემების ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' სესიის მონაცემების ჩატვირთვა და ფორმის ველების შევსება
    ''' </summary>
    Private Sub LoadSessionData(sessionRow As IList(Of Object))
        Try
            Debug.WriteLine("LoadSessionData: ვტვირთავ სესიის მონაცემებს")

            ' A: ID - უკვე გვაქვს _editRecordId და LN
            LN.Text = _editRecordId.ToString()

            ' ცვლადები თარიღისა და დროისთვის
            Dim sessionDateTime As DateTime = DateTime.Now
            Dim dateTimeLoaded As Boolean = False

            ' ვინახავთ ძველ სტატუსს ინფორმაციისთვის
            Dim oldStatus As String = ""

            ' F: თარიღი და დრო - ამას პირველად ვამუშავებთ, რადგან მასზეა დამოკიდებული რადიობუტონების ჩვენება
            If sessionRow.Count > 5 AndAlso sessionRow(5) IsNot Nothing Then
                Dim dateTimeStr = sessionRow(5).ToString().Trim()

                ' პარსინგის მცდელობა
                If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                  System.Globalization.CultureInfo.InvariantCulture,
                                  System.Globalization.DateTimeStyles.None,
                                  sessionDateTime) OrElse
               DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                    ' დავაყენოთ თარიღი
                    DTP1.Value = sessionDateTime.Date

                    ' დავაყენოთ საათი და წუთები
                    THour.Text = sessionDateTime.Hour.ToString("00")
                    TMin.Text = sessionDateTime.Minute.ToString("00")

                    dateTimeLoaded = True
                    Debug.WriteLine($"LoadSessionData: თარიღი და დრო დაყენებულია - {sessionDateTime}")
                End If
            End If

            ' D: ბენეფიციარის სახელი
            If sessionRow.Count > 3 AndAlso sessionRow(3) IsNot Nothing Then
                Dim beneName = sessionRow(3).ToString().Trim()

                ' ვეძებთ და ვაყენებთ სახელს კომბობოქსში
                For i As Integer = 0 To CBBeneName.Items.Count - 1
                    If String.Equals(CBBeneName.Items(i).ToString(), beneName, StringComparison.OrdinalIgnoreCase) Then
                        CBBeneName.SelectedIndex = i
                        Debug.WriteLine($"LoadSessionData: ბენეფიციარის სახელი დაყენებულია - {beneName}")
                        ' ვიცდით გვარების ჩატვირთვას
                        Application.DoEvents()
                        System.Threading.Thread.Sleep(300)
                        Exit For
                    End If
                Next
            End If

            ' E: ბენეფიციარის გვარი
            If sessionRow.Count > 4 AndAlso sessionRow(4) IsNot Nothing Then
                Dim beneSurname = sessionRow(4).ToString().Trim()

                ' ვეძებთ და ვაყენებთ გვარს კომბობოქსში
                For i As Integer = 0 To CBBeneSurname.Items.Count - 1
                    If String.Equals(CBBeneSurname.Items(i).ToString(), beneSurname, StringComparison.OrdinalIgnoreCase) Then
                        CBBeneSurname.SelectedIndex = i
                        Debug.WriteLine($"LoadSessionData: ბენეფიციარის გვარი დაყენებულია - {beneSurname}")
                        Exit For
                    End If
                Next
            End If

            ' G: ხანგრძლივობა
            If sessionRow.Count > 6 AndAlso sessionRow(6) IsNot Nothing Then
                Dim duration As Integer = 60 ' ნაგულისხმევი
                If Integer.TryParse(sessionRow(6).ToString(), duration) Then
                    TDur.Text = duration.ToString()
                    Debug.WriteLine($"LoadSessionData: ხანგრძლივობა დაყენებულია - {duration}")
                End If
            End If

            ' H: ჯგუფური
            If sessionRow.Count > 7 AndAlso sessionRow(7) IsNot Nothing Then
                Dim isGroup As Boolean = False
                If Boolean.TryParse(sessionRow(7).ToString(), isGroup) Then
                    CBGroup.Checked = isGroup
                    Debug.WriteLine($"LoadSessionData: ჯგუფური ჩეკბოქსი დაყენებულია - {isGroup}")
                End If
            End If

            ' I: თერაპევტი
            If sessionRow.Count > 8 AndAlso sessionRow(8) IsNot Nothing Then
                Dim therapistName = sessionRow(8).ToString().Trim()

                ' ვეძებთ და ვაყენებთ თერაპევტს კომბობოქსში
                For i As Integer = 0 To CBPer.Items.Count - 1
                    If String.Equals(CBPer.Items(i).ToString(), therapistName, StringComparison.OrdinalIgnoreCase) Then
                        CBPer.SelectedIndex = i
                        Debug.WriteLine($"LoadSessionData: თერაპევტი დაყენებულია - {therapistName}")
                        Exit For
                    End If
                Next
            End If

            ' J: თერაპიის ტიპი
            If sessionRow.Count > 9 AndAlso sessionRow(9) IsNot Nothing Then
                Dim therapyType = sessionRow(9).ToString().Trim()

                ' ვეძებთ და ვაყენებთ თერაპიის ტიპს კომბობოქსში
                For i As Integer = 0 To CBTer.Items.Count - 1
                    If String.Equals(CBTer.Items(i).ToString(), therapyType, StringComparison.OrdinalIgnoreCase) Then
                        CBTer.SelectedIndex = i
                        Debug.WriteLine($"LoadSessionData: თერაპიის ტიპი დაყენებულია - {therapyType}")
                        Exit For
                    End If
                Next
            End If

            ' K: სივრცე
            If sessionRow.Count > 10 AndAlso sessionRow(10) IsNot Nothing Then
                Dim spaceName = sessionRow(10).ToString().Trim()

                ' ვპოულობთ შესაბამის სივრცის ღილაკს და ვაჭერთ
                For Each btn As Button In Me.Controls.OfType(Of Button)()
                    If btn.Name.StartsWith("BTNS") AndAlso String.Equals(btn.Text, spaceName, StringComparison.OrdinalIgnoreCase) Then
                        ' ხელოვნურად გამოვიწვიოთ ღილაკზე დაჭერის ივენთი
                        Dim btnClickMethod = btn.GetType().GetMethod("OnClick", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                        If btnClickMethod IsNot Nothing Then
                            btnClickMethod.Invoke(btn, New Object() {EventArgs.Empty})
                            Debug.WriteLine($"LoadSessionData: სივრცე დაყენებულია - {spaceName}")
                        End If
                        Exit For
                    End If
                Next
            End If

            ' L: ფასი
            If sessionRow.Count > 11 AndAlso sessionRow(11) IsNot Nothing Then
                Dim price As Decimal = 0
                If Decimal.TryParse(sessionRow(11).ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, price) Then
                    TCost.Text = price.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                    Debug.WriteLine($"LoadSessionData: ფასი დაყენებულია - {price}")
                End If
            End If

            ' M: სტატუსი
            If sessionRow.Count > 12 AndAlso sessionRow(12) IsNot Nothing Then
                oldStatus = sessionRow(12).ToString().Trim()
                Debug.WriteLine($"LoadSessionData: სესიის სტატუსი არის '{oldStatus}'")
            End If

            ' N: დაფინანსების პროგრამა
            If sessionRow.Count > 13 AndAlso sessionRow(13) IsNot Nothing Then
                Dim fundingProgram = sessionRow(13).ToString().Trim()

                ' ვეძებთ და ვაყენებთ დაფინანსების პროგრამას კომბობოქსში
                For i As Integer = 0 To CBDaf.Items.Count - 1
                    If String.Equals(CBDaf.Items(i).ToString(), fundingProgram, StringComparison.OrdinalIgnoreCase) Then
                        CBDaf.SelectedIndex = i
                        Debug.WriteLine($"LoadSessionData: დაფინანსების პროგრამა დაყენებულია - {fundingProgram}")
                        Exit For
                    End If
                Next
            End If

            ' O: კომენტარი
            If sessionRow.Count > 14 AndAlso sessionRow(14) IsNot Nothing Then
                TCom.Text = sessionRow(14).ToString()
                Debug.WriteLine("LoadSessionData: კომენტარი დაყენებულია")
            End If

            ' ახლა გავაკეთოთ რადიობუტონების ხილვადობის კონტროლი
            If dateTimeLoaded Then
                ' შევამოწმოთ არის თუ არა თარიღი წარსულში
                Dim isPastDate As Boolean = sessionDateTime < DateTime.Now
                Debug.WriteLine($"LoadSessionData: თარიღი არის წარსულში: {isPastDate}")

                If isPastDate Then
                    ' წარსული თარიღისთვის - გამოვაჩინოთ რადიობუტონები, მაგრამ არ მოვნიშნოთ
                    LPlan.Visible = False

                    ' გავხადოთ ყველა რადიობუტონი ხილული მაგრამ არ მოვნიშნოთ არცერთი
                    For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                        If rb.Name.StartsWith("RB") Then
                            rb.Visible = True
                            rb.Checked = False
                        End If
                    Next

                    Debug.WriteLine("LoadSessionData: რადიობუტონები გამოჩნდა, ძველი სტატუსი იყო: " + oldStatus)

                    ' შეტყობინება რომ აირჩიოს სტატუსი
                    LWarning.Text = $"ძველი სტატუსი: '{oldStatus}'. გთხოვთ აირჩიოთ ახალი სტატუსი."
                    LWarning.ForeColor = Color.DarkOrange
                Else
                    ' მომავალი თარიღისთვის
                    LPlan.Visible = True

                    ' დავმალოთ ყველა რადიობუტონი
                    For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                        rb.Visible = False
                        rb.Checked = False
                    Next

                    Debug.WriteLine("LoadSessionData: მომავალი თარიღი, რადიობუტონები დამალულია")
                End If
            Else
                ' თუ თარიღი ვერ ჩაიტვირთა, მაშინ სტატუსის მიხედვით ვიმოქმედოთ
                Debug.WriteLine("LoadSessionData: თარიღი ვერ ჩაიტვირთა, ვმოქმედებთ სტატუსის მიხედვით")

                If oldStatus.ToLower() = "დაგეგმილი" Then
                    ' დავმალოთ რადიობუტონები და გამოვაჩინოთ LPlan
                    LPlan.Visible = True
                    For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                        rb.Visible = False
                        rb.Checked = False
                    Next
                Else
                    ' გამოვაჩინოთ რადიობუტონები და ვთხოვოთ მომხმარებელს არჩევა
                    LPlan.Visible = False
                    For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                        If rb.Name.StartsWith("RB") Then
                            rb.Visible = True
                            rb.Checked = False
                        End If
                    Next

                    ' შეტყობინება რომ აირჩიოს სტატუსი
                    LWarning.Text = $"ძველი სტატუსი: '{oldStatus}'. გთხოვთ აირჩიოთ ახალი სტატუსი."
                    LWarning.ForeColor = Color.DarkOrange
                End If
            End If

            ' BtnAdd ტექსტის შეცვლა
            BtnAdd.Text = "შენახვა"

            ' ვალიდაცია - შეგვიძლია გამოვიყენოთ ValidateFormInputs, თუ ის არ ბლოკავს შენახვას რადიობუტონების არ მონიშვნის გამო
            ValidateFormInputs()

            Debug.WriteLine("LoadSessionData: სესიის მონაცემების ჩატვირთვა დასრულებულია")

        Catch ex As Exception
            Debug.WriteLine($"LoadSessionData: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის მონაცემების ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
    ''' <summary>
    ''' დამატებითი მეთოდი კომბობოქსის ელემენტების დასახატად
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
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

                ' ფორმის ვალიდაცია
                ValidateFormInputs()

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
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
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
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
    End Sub

    ''' <summary>
    ''' CBPer-ს SelectedIndexChanged ივენთი - თერაპევტის არჩევისას
    ''' </summary>
    Private Sub CBPer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPer.SelectedIndexChanged
        Try
            ' შევამოწმოთ თერაპევტის მდგომარეობა
            CheckTherapistAvailability()
            ' ფორმის ვალიდაცია
            ValidateFormInputs()

        Catch ex As Exception
            MessageBox.Show($"შეცდომა თერაპევტის არჩევისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' CBTer-ს SelectedIndexChanged ივენთი - თერაპევტის არჩევისას
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CBTer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTer.SelectedIndexChanged
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
    End Sub
    ''' <summary>
    ''' CBDaf-ს SelectedIndexChanged ივენთი - თერაპევტის არჩევისას
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CBDaf_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBDaf.SelectedIndexChanged
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
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
    ''' გამართული ვერსია TCost_KeyPress ივენთისთვის
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
    ''' გამართული ვერსია TCost_TextChanged ივენთისთვის, რომელიც შეინარჩუნებს შეყვანილ მნიშვნელობას
    ''' </summary>
    Private Sub TCost_TextChanged(sender As Object, e As EventArgs) Handles TCost.TextChanged
        ' ტექსტი რჩება უცვლელი შეყვანის დროს
        ' ვალიდაცია მხოლოდ ფორმის დასრულებისას ხდება

        ' ფორმის ვალიდაცია
        ValidateFormInputs()
    End Sub

    ''' <summary>
    ''' გამართული ვერსია TCost_Leave ივენთისთვის, რომელიც არ გაანულებს მნიშვნელობას
    ''' და სწორად დააფორმატებს ფასს
    ''' </summary>
    Private Sub TCost_Leave(sender As Object, e As EventArgs) Handles TCost.Leave
        Try
            ' ტექსტის ფორმატირება როგორც რიცხვის
            Dim text As String = TCost.Text.Replace(",", ".")

            ' თუ ტექსტი ცარიელია, დავაყენოთ 0
            If String.IsNullOrWhiteSpace(text) Then
                TCost.Text = "0"
                Return
            End If

            ' მოვსინჯოთ რიცხვის პარსინგი
            Dim value As Decimal
            If Decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, value) Then
                ' მხოლოდ მაშინ ვაფორმატირებთ, თუ ეს ნამდვილად რიცხვია
                ' ვიყენებთ უცვლელ კულტურას ფორმატირებისთვის
                TCost.Text = value.ToString("0.00", CultureInfo.InvariantCulture)

                ' ფორმის ვალიდაცია ფორმატირების შემდეგ
                ValidateFormInputs()
            End If
        Catch ex As Exception
            Debug.WriteLine($"TCost_Leave: შეცდომა - {ex.Message}")
            ' ტექსტს არ ვცვლით შეცდომის შემთხვევაში
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
            Dim selectedDateTime As DateTime = GetSelectedDateTime()

            ' სივრცეების დატვირთულობის ინფორმაციის წამოღება Google Sheets-დან
            Dim occupiedSpaces As Dictionary(Of String, IList(Of Object)) = GetOccupiedSpacesWithDetails(selectedDateTime)
            Dim groupSpaces As Dictionary(Of String, Boolean) = GetGroupSpaces(selectedDateTime)

            ' ყველა BTNS ღილაკის მოძებნა
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") Then
                    Dim spaceName As String = btn.Text.Trim()

                    If occupiedSpaces.ContainsKey(spaceName) Then
                        ' დაკავებული სივრცე - ღია წითელი
                        btn.BackColor = Color.FromArgb(255, 200, 200)
                    ElseIf groupSpaces.ContainsKey(spaceName) Then
                        ' ჯგუფური სესიის სივრცე - ღია ყვითელი
                        btn.BackColor = Color.FromArgb(255, 255, 200)
                    Else
                        ' თავისუფალი სივრცე - ღია მწვანე
                        btn.BackColor = Color.FromArgb(200, 255, 200)
                    End If
                End If
            Next

            ' სტატუსის ტექსტის განახლება LMsgSpace ლეიბლში
            If occupiedSpaces.Count > 0 Then
                ' თუ დაკავებული სივრცეები არსებობს, გამოვიტანოთ დეტალურად ერთი მაგალითი
                Dim firstOccupiedSpace As String = occupiedSpaces.Keys.First()
                Dim sessionDetails As IList(Of Object) = occupiedSpaces(firstOccupiedSpace)

                ' დეტალური ინფორმაცია ბენეფიციარზე, თერაპევტზე, თერაპიის ტიპზე
                Dim beneName As String = If(sessionDetails.Count > 3, sessionDetails(3).ToString(), "")
                Dim beneSurname As String = If(sessionDetails.Count > 4, sessionDetails(4).ToString(), "")
                Dim therapistName As String = If(sessionDetails.Count > 8, sessionDetails(8).ToString(), "")
                Dim therapyType As String = If(sessionDetails.Count > 9, sessionDetails(9).ToString(), "")
                Dim dateTimeStr As String = If(sessionDetails.Count > 5, sessionDetails(5).ToString(), "")

                ' მრავალხაზიანი ტექსტი ლეიბლში
                Dim infoText As New System.Text.StringBuilder()
                infoText.AppendLine($"სივრცე '{firstOccupiedSpace}' დაკავებულია:")
                infoText.AppendLine($"ბენეფიციარი: {beneName} {beneSurname}")
                infoText.AppendLine($"თერაპევტი: {therapistName}")
                infoText.AppendLine($"თერაპია: {therapyType}")

                ' თუ არის სხვა დაკავებული სივრცეებიც
                If occupiedSpaces.Count > 1 Then
                    infoText.Append($"და კიდევ {occupiedSpaces.Count - 1} დაკავებული სივრცე")
                End If

                LMsgSpace.Text = infoText.ToString()
                LMsgSpace.ForeColor = Color.DarkRed
            ElseIf groupSpaces.Count > 0 Then
                ' თუ მხოლოდ ჯგუფური სივრცეები არსებობს
                Dim firstGroupSpace As String = groupSpaces.Keys.First()

                LMsgSpace.Text = $"სივრცე '{firstGroupSpace}' გამოიყენება ჯგუფური სესიისთვის" &
                             If(groupSpaces.Count > 1, $"{Environment.NewLine}და კიდევ {groupSpaces.Count - 1} ჯგუფური სივრცე", "")
                LMsgSpace.ForeColor = Color.DarkOrange
            Else
                ' ყველა სივრცე თავისუფალია
                LMsgSpace.Text = "მოცემულ დროს ყველა სივრცე თავისუფალია"
                LMsgSpace.ForeColor = Color.DarkGreen
            End If

        Catch ex As Exception
            Debug.WriteLine($"InitializeSpaceButtons: შეცდომა - {ex.Message}")
            MessageBox.Show($"სივრცეების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ამოწმებს Google Sheets-დან რომელი სივრცეებია დაკავებული და აბრუნებს სესიის დეტალებს
    ''' </summary>
    Private Function GetOccupiedSpacesWithDetails(selectedDateTime As DateTime) As Dictionary(Of String, IList(Of Object))
        Dim occupiedSpaces As New Dictionary(Of String, IList(Of Object))

        Try
            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                Debug.WriteLine("GetOccupiedSpacesWithDetails: dataService არ არის ინიციალიზებული")
                Return occupiedSpaces
            End If

            ' დროის შუალედის განსაზღვრა (უფრო ზუსტი შუალედი)
            Dim selectedDuration As Integer = 0
            Integer.TryParse(TDur.Text, selectedDuration)
            If selectedDuration <= 0 Then selectedDuration = 60 ' ნაგულისხმევი ხანგრძლივობა

            Dim startTime As DateTime = selectedDateTime
            Dim endTime As DateTime = selectedDateTime.AddMinutes(selectedDuration)

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:M") ' გაფართოებული დიაპაზონი სტატუსის ჩათვლით

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 11 Then ' სულ მცირე K სვეტამდე
                        Try
                            ' სივრცის სახელი K სვეტიდან (ინდექსი 10)
                            Dim spaceName As String = row(10).ToString().Trim()

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = row(5).ToString()
                            Dim sessionDateTime As DateTime

                            ' სესიის ხანგრძლივობა G სვეტიდან (ინდექსი 6)
                            Dim sessionDuration As Integer = 60 ' ნაგულისხმევი ხანგრძლივობა
                            If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6)?.ToString()) Then
                                Integer.TryParse(row(6).ToString(), sessionDuration)
                            End If

                            ' სესიის სტატუსის შემოწმება M სვეტიდან (ინდექსი 12)
                            Dim sessionStatus As String = ""
                            If row.Count > 12 AndAlso Not String.IsNullOrEmpty(row(12)?.ToString()) Then
                                sessionStatus = row(12).ToString().Trim().ToLower()
                            End If

                            ' თუ სესია გაუქმებულია, არ ჩავთვალოთ დაკავებულად
                            If sessionStatus = "გაუქმებული" OrElse sessionStatus = "გაუქმება" Then
                                Continue For
                            End If

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                System.Globalization.CultureInfo.InvariantCulture,
                                                System.Globalization.DateTimeStyles.None,
                                                sessionDateTime) OrElse
                           DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' სესიის დასრულების დრო
                                Dim sessionEndTime As DateTime = sessionDateTime.AddMinutes(sessionDuration)

                                ' შევამოწმოთ გადაკვეთა ორ დროის შუალედს შორის
                                ' [startTime, endTime] და [sessionDateTime, sessionEndTime]
                                If (startTime < sessionEndTime) AndAlso (endTime > sessionDateTime) Then
                                    ' დაკავებული სივრცე (დროითი გადაფარვით) - შევინახოთ მთელი მწკრივი დეტალებისთვის
                                    occupiedSpaces(spaceName) = row
                                    Debug.WriteLine($"GetOccupiedSpacesWithDetails: ნაპოვნია დაკავებული სივრცე: {spaceName}, დრო: {sessionDateTime}")
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
            Debug.WriteLine($"GetOccupiedSpacesWithDetails: შეცდომა - {ex.Message}")
        End Try

        Return occupiedSpaces
    End Function

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

            ' დროის შუალედის განსაზღვრა (უფრო ზუსტი შუალედი)
            Dim selectedDuration As Integer = 0
            Integer.TryParse(TDur.Text, selectedDuration)
            If selectedDuration <= 0 Then selectedDuration = 60 ' ნაგულისხმევი ხანგრძლივობა

            ' ახალი შუალედი - ფაქტიური სესიების გადაფარვის შემოწმებით
            Dim startTime As DateTime = selectedDateTime
            Dim endTime As DateTime = selectedDateTime.AddMinutes(selectedDuration)

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

                            ' სესიის ხანგრძლივობა G სვეტიდან (ინდექსი 6)
                            Dim sessionDuration As Integer = 60 ' ნაგულისხმევი ხანგრძლივობა
                            If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6)?.ToString()) Then
                                Integer.TryParse(row(6).ToString(), sessionDuration)
                            End If

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                System.Globalization.CultureInfo.InvariantCulture,
                                                System.Globalization.DateTimeStyles.None,
                                                sessionDateTime) OrElse
                           DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' სესიის დასრულების დრო
                                Dim sessionEndTime As DateTime = sessionDateTime.AddMinutes(sessionDuration)

                                ' შევამოწმოთ გადაკვეთა ორ დროის შუალედს შორის
                                ' [startTime, endTime] და [sessionDateTime, sessionEndTime]
                                If (startTime < sessionEndTime) AndAlso (endTime > sessionDateTime) Then
                                    ' დაკავებული სივრცე (დროითი გადაფარვით)
                                    occupiedSpaces(spaceName) = True
                                    Debug.WriteLine($"GetOccupiedSpaces: ნაპოვნია დაკავებული სივრცე: {spaceName}, დრო: {sessionDateTime}")
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

            ' დროის შუალედის განსაზღვრა (უფრო ზუსტი შუალედი)
            Dim selectedDuration As Integer = 0
            Integer.TryParse(TDur.Text, selectedDuration)
            If selectedDuration <= 0 Then selectedDuration = 60 ' ნაგულისხმევი ხანგრძლივობა

            ' ახალი შუალედი - ფაქტიური სესიების გადაფარვის შემოწმებით
            Dim startTime As DateTime = selectedDateTime
            Dim endTime As DateTime = selectedDateTime.AddMinutes(selectedDuration)

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
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

                            ' სესიის ხანგრძლივობა G სვეტიდან (ინდექსი 6)
                            Dim sessionDuration As Integer = 60 ' ნაგულისხმევი ხანგრძლივობა
                            If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6)?.ToString()) Then
                                Integer.TryParse(row(6).ToString(), sessionDuration)
                            End If

                            ' პარსინგის მცდელობა
                            If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                System.Globalization.CultureInfo.InvariantCulture,
                                                System.Globalization.DateTimeStyles.None,
                                                sessionDateTime) OrElse
                           DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                ' სესიის დასრულების დრო
                                Dim sessionEndTime As DateTime = sessionDateTime.AddMinutes(sessionDuration)

                                ' შევამოწმოთ გადაკვეთა ორ დროის შუალედს შორის და რომ ეს ჯგუფური სესიაა 
                                If (startTime < sessionEndTime) AndAlso (endTime > sessionDateTime) AndAlso
                                row.Count > 7 AndAlso row(7).ToString().ToLower() = "true" Then
                                    ' ჯგუფური სესიის სივრცე (დროითი გადაფარვით)
                                    groupSpaces(spaceName) = True
                                    Debug.WriteLine($"GetGroupSpaces: ნაპოვნია ჯგუფური სივრცე: {spaceName}, დრო: {sessionDateTime}")
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
    ''' სივრცის ღილაკზე დაჭერის ივენთი - გასწორებული ვერსია
    ''' </summary>

    Private Sub SpaceButton_Click(sender As Object, e As EventArgs)
        Try
            ' სტატიკური ცვლადი წინა ღილაკის დასამახსოვრებლად
            Static lastSelectedButton As Button = Nothing

            ' ახალი არჩეული ღილაკი
            Dim clickedButton As Button = DirectCast(sender, Button)
            Dim selectedSpace As String = clickedButton.Text

            ' თარიღის და დაკავებული სივრცეების ინფორმაციის მიღება - ერთხელ
            Dim selectedDateTime As DateTime = GetSelectedDateTime()
            Dim occupiedSpaces As Dictionary(Of String, IList(Of Object)) = GetOccupiedSpacesWithDetails(selectedDateTime)
            Dim groupSpaces As Dictionary(Of String, Boolean) = GetGroupSpaces(selectedDateTime)

            ' თუ არის წინა არჩეული ღილაკი და ის განსხვავდება ახლანდელისგან
            If lastSelectedButton IsNot Nothing AndAlso lastSelectedButton IsNot clickedButton Then
                ' წინა ღილაკის ტექსტი
                Dim lastSpaceName As String = lastSelectedButton.Text.Trim()

                ' წინა ღილაკისთვის შესაბამისი ფერის დაბრუნება
                If occupiedSpaces.ContainsKey(lastSpaceName) Then
                    lastSelectedButton.BackColor = Color.FromArgb(255, 200, 200) ' ღია წითელი
                ElseIf groupSpaces.ContainsKey(lastSpaceName) Then
                    lastSelectedButton.BackColor = Color.FromArgb(255, 255, 200) ' ღია ყვითელი
                Else
                    lastSelectedButton.BackColor = Color.FromArgb(200, 255, 200) ' ღია მწვანე
                End If
            End If

            ' მიმდინარე ღილაკის მონიშვნა
            clickedButton.BackColor = Color.FromArgb(173, 216, 230) ' ცისფერი
            lastSelectedButton = clickedButton

            ' ვნახოთ, დაკავებულია თუ არა ეს სივრცე და ვაჩვენოთ დეტალური ინფორმაცია
            If occupiedSpaces.ContainsKey(selectedSpace) Then
                ' დაკავებულია - ვაჩვენოთ დეტალები
                Dim sessionDetails As IList(Of Object) = occupiedSpaces(selectedSpace)

                ' დეტალური ინფორმაცია ბენეფიციარზე, თერაპევტზე, თერაპიის ტიპზე
                Dim beneName As String = If(sessionDetails.Count > 3, sessionDetails(3).ToString(), "")
                Dim beneSurname As String = If(sessionDetails.Count > 4, sessionDetails(4).ToString(), "")
                Dim therapistName As String = If(sessionDetails.Count > 8, sessionDetails(8).ToString(), "")
                Dim therapyType As String = If(sessionDetails.Count > 9, sessionDetails(9).ToString(), "")
                Dim dateTimeStr As String = If(sessionDetails.Count > 5, sessionDetails(5).ToString(), "")

                ' მრავალხაზიანი ტექსტი ლეიბლში
                Dim infoText As New System.Text.StringBuilder()
                infoText.AppendLine($"არჩეული სივრცე '{selectedSpace}' დაკავებულია:")
                infoText.AppendLine($"ბენეფიციარი: {beneName} {beneSurname}")
                infoText.AppendLine($"თერაპევტი: {therapistName}")
                infoText.AppendLine($"თერაპია: {therapyType}")
                'infoText.Append($"თარიღი: {dateTimeStr}")

                LMsgSpace.Text = infoText.ToString()
                LMsgSpace.ForeColor = Color.DarkRed
            ElseIf groupSpaces.ContainsKey(selectedSpace) Then
                ' ჯგუფური სესია
                LMsgSpace.Text = $"არჩეული სივრცე '{selectedSpace}' გამოიყენება ჯგუფური სესიისთვის"
                LMsgSpace.ForeColor = Color.DarkOrange
            Else
                ' თავისუფალი სივრცე
                LMsgSpace.Text = $"არჩეული სივრცე: {selectedSpace} (თავისუფალია)"
                LMsgSpace.ForeColor = Color.DarkGreen
            End If
            ' ფორმის ვალიდაცია
            ValidateFormInputs()
            Debug.WriteLine($"SpaceButton_Click: არჩეულია სივრცე '{selectedSpace}'")
        Catch ex As Exception
            Debug.WriteLine($"SpaceButton_Click: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DTP1-ის CloseUp ივენთი - კალენდრის დახურვისას გააქტიურდება
    ''' </summary>
    Private Sub DTP1_CloseUp(sender As Object, e As EventArgs)
        ' დროის ცვლილების დამმუშავებლის გამოძახება
        DateTimeChanged(sender, e)
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
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

            ' სესიის ხანგრძლივობის მიღება
            Dim selectedDuration As Integer = 60 ' ნაგულისხმევი ხანგრძლივობა
            Integer.TryParse(TDur.Text, selectedDuration)

            ' ზუსტი დროის შუალედის განსაზღვრა - აქტუალური სესიის დაწყება და დასრულება
            Dim startTime As DateTime = selectedDateTime
            Dim endTime As DateTime = selectedDateTime.AddMinutes(selectedDuration)

            ' არჩეული ბენეფიციარი
            Dim beneficiaryName As String = CBBeneName.Text.Trim()
            Dim beneficiarySurname As String = CBBeneSurname.Text.Trim()

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:M") ' M-მდე, რომ სტატუსიც მოვიცვათ

            ' ბენეფიციარის არსებული სესიების მოძებნა
            Dim beneficiarySession As IList(Of Object) = Nothing

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 13 Then ' M სვეტამდე
                        Try
                            ' ბენეფიციარის სახელი (D სვეტი - ინდექსი 3) და გვარი (E სვეტი - ინდექსი 4)
                            Dim rowBeneName As String = If(row.Count > 3, row(3).ToString().Trim(), "")
                            Dim rowBeneSurname As String = If(row.Count > 4, row(4).ToString().Trim(), "")

                            ' სესიის სტატუსი (M სვეტი - ინდექსი 12)
                            Dim sessionStatus As String = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")

                            ' გაუქმებული ან შესრულებული სესიების გამოტოვება
                            If sessionStatus = "გაუქმებული" OrElse sessionStatus = "გაუქმება" Then
                                Continue For
                            End If

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = If(row.Count > 5, row(5).ToString(), "")
                            Dim sessionDateTime As DateTime

                            ' სესიის ხანგრძლივობა G სვეტიდან (ინდექსი 6)
                            Dim sessionDuration As Integer = 60 ' ნაგულისხმევი
                            If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6)?.ToString()) Then
                                Integer.TryParse(row(6).ToString(), sessionDuration)
                            End If

                            ' ვამოწმებთ თუ ეს იგივე ბენეფიციარია
                            If String.Equals(rowBeneName, beneficiaryName, StringComparison.OrdinalIgnoreCase) AndAlso
                           String.Equals(rowBeneSurname, beneficiarySurname, StringComparison.OrdinalIgnoreCase) Then

                                ' პარსინგის მცდელობა
                                If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                     System.Globalization.CultureInfo.InvariantCulture,
                                                     System.Globalization.DateTimeStyles.None,
                                                     sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                    ' სესიის დასრულების დრო
                                    Dim sessionEndTime As DateTime = sessionDateTime.AddMinutes(sessionDuration)

                                    ' შევამოწმოთ ორი დროის შუალედის გადაკვეთა
                                    If (startTime < sessionEndTime) AndAlso (endTime > sessionDateTime) Then
                                        ' ნაპოვნია ბენეფიციარის სესია მოცემულ დროის შუალედში
                                        beneficiarySession = row
                                        Exit For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Debug.WriteLine($"CheckBeneficiaryAvailability: მწკრივის შემოწმების შეცდომა - {ex.Message}")
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
                ' თუ თერაპევტი არ არის არჩეული, გავასუფთაოთ ლეიბლი
                LMsgPer.Text = "აირჩიეთ თერაპევტი"
                LMsgPer.ForeColor = Color.Black
                CBPer.BackColor = SystemColors.Window
                Return
            End If

            ' არჩეული თარიღი და დრო
            Dim selectedDateTime As DateTime = GetSelectedDateTime()

            ' სესიის ხანგრძლივობის მიღება
            Dim selectedDuration As Integer = 60 ' ნაგულისხმევი ხანგრძლივობა
            Integer.TryParse(TDur.Text, selectedDuration)

            ' ზუსტი დროის შუალედის განსაზღვრა - აქტუალური სესიის დაწყება და დასრულება
            Dim startTime As DateTime = selectedDateTime
            Dim endTime As DateTime = selectedDateTime.AddMinutes(selectedDuration)

            ' არჩეული თერაპევტი
            Dim therapistName As String = CBPer.Text.Trim()

            ' DB-Schedule ფურცლიდან მონაცემების წამოღება
            Dim scheduleData = dataService.GetData("DB-Schedule!A2:M") ' M-მდე, რომ სტატუსიც მოვიცვათ

            ' თერაპევტის არსებული სესიების მოძებნა
            Dim therapistSession As IList(Of Object) = Nothing

            If scheduleData IsNot Nothing AndAlso scheduleData.Count > 0 Then
                For Each row In scheduleData
                    If row.Count >= 13 Then ' M სვეტამდე
                        Try
                            ' თერაპევტის სახელი (I სვეტი - ინდექსი 8)
                            Dim rowTherapistName As String = If(row.Count > 8, row(8).ToString().Trim(), "")

                            ' სესიის სტატუსი (M სვეტი - ინდექსი 12)
                            Dim sessionStatus As String = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")

                            ' გაუქმებული ან შესრულებული სესიების გამოტოვება
                            If sessionStatus = "გაუქმებული" OrElse sessionStatus = "გაუქმება" Then
                                Continue For
                            End If

                            ' თარიღი და დრო F სვეტიდან (ინდექსი 5)
                            Dim dateTimeStr As String = If(row.Count > 5, row(5).ToString(), "")
                            Dim sessionDateTime As DateTime

                            ' სესიის ხანგრძლივობა G სვეტიდან (ინდექსი 6)
                            Dim sessionDuration As Integer = 60 ' ნაგულისხმევი
                            If row.Count > 6 AndAlso Not String.IsNullOrEmpty(row(6)?.ToString()) Then
                                Integer.TryParse(row(6).ToString(), sessionDuration)
                            End If

                            ' ვამოწმებთ თუ ეს იგივე თერაპევტია
                            If String.Equals(rowTherapistName, therapistName, StringComparison.OrdinalIgnoreCase) Then
                                ' პარსინგის მცდელობა
                                If DateTime.TryParseExact(dateTimeStr, "dd.MM.yyyy HH:mm",
                                                     System.Globalization.CultureInfo.InvariantCulture,
                                                     System.Globalization.DateTimeStyles.None,
                                                     sessionDateTime) OrElse
                                DateTime.TryParse(dateTimeStr, sessionDateTime) Then

                                    ' სესიის დასრულების დრო
                                    Dim sessionEndTime As DateTime = sessionDateTime.AddMinutes(sessionDuration)

                                    ' შევამოწმოთ ორი დროის შუალედის გადაკვეთა
                                    If (startTime < sessionEndTime) AndAlso (endTime > sessionDateTime) Then
                                        ' ნაპოვნია თერაპევტის სესია მოცემულ დროის შუალედში
                                        therapistSession = row
                                        Exit For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            ' ვაგრძელებთ შემდეგ მწკრივზე შეცდომის შემთხვევაში
                            Debug.WriteLine($"CheckTherapistAvailability: მწკრივის შემოწმების შეცდომა - {ex.Message}")
                            Continue For
                        End Try
                    End If
                Next
            End If

            ' თერაპევტის მდგომარეობის ვიზუალური ასახვა
            If therapistSession IsNot Nothing Then
                ' თერაპევტი დაკავებულია - წითელი ინდიკატორი
                CBPer.BackColor = Color.FromArgb(255, 200, 200) ' ღია წითელი

                ' დეტალური ინფორმაცია ბენეფიციარზე, თერაპიაზე და სივრცეზე
                Dim beneName As String = If(therapistSession.Count > 3, therapistSession(3).ToString(), "")
                Dim beneSurname As String = If(therapistSession.Count > 4, therapistSession(4).ToString(), "")
                Dim therapyType As String = If(therapistSession.Count > 9, therapistSession(9).ToString(), "")
                Dim spaceName As String = If(therapistSession.Count > 10, therapistSession(10).ToString(), "")

                ' მრავალხაზიანი ტექსტი ლეიბლში
                Dim infoText As New System.Text.StringBuilder()
                infoText.AppendLine("მოცემულ დროს თერაპევტი დაკავებულია:")
                infoText.AppendLine($"ბენეფიციარი: {beneName} {beneSurname}")
                infoText.AppendLine($"თერაპია: {therapyType}")
                infoText.Append($"სივრცე: {spaceName}")

                LMsgPer.Text = infoText.ToString()
                LMsgPer.ForeColor = Color.DarkRed
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
        ' ფორმის ვალიდაცია
        ValidateFormInputs()
    End Sub

    ''' <summary>
    ''' ფორმის ველების შემოწმება და შესაბამისი გაფრთხილებების ჩვენება
    ''' </summary>
    Private Sub ValidateFormInputs()
        Try
            ' ტექსტისთვის ბილდერი
            Dim warningText As New System.Text.StringBuilder()
            Dim warningColor As Color = Color.DarkRed
            Dim isFormValid As Boolean = True

            ' სივრცის შერჩევის სტატუსის შენახვა
            Static selectedSpaceButton As Button = Nothing

            ' 1. ბენეფიციარის შემოწმება
            If CBBeneName.SelectedIndex <= 0 OrElse String.IsNullOrEmpty(CBBeneSurname.Text) Then
                warningText.AppendLine("• გთხოვთ შეიყვანოთ ბენეფიციარი")
                CBBeneName.BackColor = Color.MistyRose
                CBBeneSurname.BackColor = Color.MistyRose
                isFormValid = False
            Else
                CBBeneName.BackColor = SystemColors.Window
                CBBeneSurname.BackColor = SystemColors.Window
            End If

            ' 2. თერაპევტის შემოწმება
            If CBPer.SelectedIndex <= 0 Then
                warningText.AppendLine("• გთხოვთ აირჩიოთ თერაპევტი")
                CBPer.BackColor = Color.MistyRose
                isFormValid = False
            Else
                CBPer.BackColor = SystemColors.Window
            End If

            ' 3. თერაპიის შემოწმება
            If CBTer.SelectedIndex <= 0 Then
                warningText.AppendLine("• გთხოვთ აირჩიოთ თერაპიის ტიპი")
                CBTer.BackColor = Color.MistyRose
                isFormValid = False
            Else
                CBTer.BackColor = SystemColors.Window
            End If

            ' 4. დაფინანსების შემოწმება
            If CBDaf.SelectedIndex <= 0 Then
                warningText.AppendLine("• გთხოვთ აირჩიოთ დაფინანსების პროგრამა")
                CBDaf.BackColor = Color.MistyRose
                isFormValid = False
            Else
                CBDaf.BackColor = SystemColors.Window
            End If

            ' 5. ფასის შემოწმება - შესწორებული ვერსია სწორი კულტურით
            Dim price As Decimal = 0
            If String.IsNullOrWhiteSpace(TCost.Text) Then
                warningText.AppendLine("• გთხოვთ შეიყვანოთ ფასი")
                TCost.BackColor = Color.MistyRose
                isFormValid = False
            ElseIf Not Decimal.TryParse(TCost.Text.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, price) OrElse price < 0 Then
                warningText.AppendLine("• გთხოვთ შეიყვანოთ ვალიდური ფასი")
                TCost.BackColor = Color.MistyRose
                isFormValid = False
            Else
                TCost.BackColor = SystemColors.Window
            End If

            ' 6. სივრცის შემოწმება
            Dim spaceSelected As Boolean = False
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") AndAlso btn.BackColor = Color.FromArgb(173, 216, 230) Then
                    spaceSelected = True
                    selectedSpaceButton = btn
                    Exit For
                End If
            Next

            If Not spaceSelected Then
                warningText.AppendLine("• გთხოვთ აირჩიოთ სივრცე")
                isFormValid = False
            End If

            ' 7. დასრულების სტატუსის შემოწმება თუ საჭიროა - შეცვლილი ლოგიკით
            Dim selectedDateTime As DateTime = GetSelectedDateTime()
            Dim isFutureSession As Boolean = selectedDateTime > DateTime.Now

            If Not isFutureSession Then
                ' წარსული თარიღისთვის სტატუსი აუცილებელია
                Dim statusSelected As Boolean = False
                For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                    If rb.Name.StartsWith("RB") AndAlso rb.Checked Then
                        statusSelected = True
                        Exit For
                    End If
                Next

                If Not statusSelected Then
                    warningText.AppendLine("• გთხოვთ აირჩიოთ სესიის შესრულების სტატუსი")
                    isFormValid = False
                End If
            End If

            ' ღილაკის მართვა და ტექსტის დაყენება
            If isFormValid Then
                BtnAdd.Visible = True

                ' ყველა ველი სწორია - შეტყობინება დამოკიდებულია სესიის დროზე
                warningText.Clear()
                If isFutureSession Then
                    warningText.AppendLine("ყველა ველი შევსებულია!")
                    warningText.AppendLine("სესია დაემატება სტატუსით 'დაგეგმილი'.")
                Else
                    warningText.AppendLine("ყველა ველი შევსებულია!")
                    warningText.AppendLine("შეგიძლიათ დააჭიროთ დამატების ღილაკს.")
                End If
                warningColor = Color.DarkGreen
            Else
                BtnAdd.Visible = False
            End If

            ' LWarning ლეიბლის განახლება
            LWarning.Text = warningText.ToString()
            LWarning.ForeColor = warningColor

        Catch ex As Exception
            Debug.WriteLine($"ValidateFormInputs: შეცდომა - {ex.Message}")
            LWarning.Text = "ფორმის შემოწმებისას დაფიქსირდა შეცდომა"
            LWarning.ForeColor = Color.DarkRed
            BtnAdd.Visible = False
        End Try
    End Sub
    ''' <summary>
    ''' BtnClear ღილაკზე დაჭერის ჰენდლერი - ყველა ველის გასუფთავება
    ''' </summary>
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        Try
            ' დებაგ ინფორმაცია
            Debug.WriteLine("BtnClear_Click: ფორმის ველების გასუფთავება დაიწყო")

            ' 1. ComboBox-ების გასუფთავება - პირველ ელემენტზე (მინიშნებაზე) დაბრუნება
            CBBeneName.SelectedIndex = 0
            CBBeneSurname.Items.Clear()
            CBBeneSurname.Enabled = False
            CBPer.SelectedIndex = 0
            CBTer.SelectedIndex = 0
            CBDaf.SelectedIndex = 0

            ' 2. DateTimePicker-ის დაბრუნება მიმდინარე დღეზე
            DTP1.Value = DateTime.Today

            ' 3. დროის ველების დაბრუნება საწყის მნიშვნელობებზე
            THour.Text = "12"
            TMin.Text = "00"
            TDur.Text = "60"

            ' 4. ფასის ველის გასუფთავება
            TCost.Text = "0"

            ' 5. ჯგუფური ჩეკბოქსის გაუქმება
            CBGroup.Checked = False

            ' 6. კომენტარის ველის გასუფთავება
            TCom.Text = ""

            ' 7. სივრცის არჩევანის გაუქმება - ყველა ღილაკის საწყის ფერზე დაბრუნება
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") Then
                    ' დავაბრუნოთ ღილაკები საწყის ფერებზე
                    btn.BackColor = Color.FromArgb(240, 240, 240) ' სტანდარტული ღილაკის ფერი
                End If
            Next

            ' 8. რადიობუტონების მონიშვნის გაუქმება
            For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                If rb.Name.StartsWith("RB") Then
                    rb.Checked = False
                End If
            Next

            ' 9. შეტყობინების ლეიბლების გასუფთავება/განახლება
            LMsgBene.Text = "აირჩიეთ ბენეფიციარი"
            LMsgBene.ForeColor = Color.Black

            LMsgPer.Text = "აირჩიეთ თერაპევტი"
            LMsgPer.ForeColor = Color.Black

            LMsgSpace.Text = "აირჩიეთ სივრცე"
            LMsgSpace.ForeColor = Color.Black

            LWarning.Text = "გთხოვთ შეავსოთ ყველა აუცილებელი ველი"
            LWarning.ForeColor = Color.Black

            ' 10. ველების ფონის ფერების დაბრუნება ნორმალურ მდგომარეობაში
            CBBeneName.BackColor = SystemColors.Window
            CBBeneSurname.BackColor = SystemColors.Window
            CBPer.BackColor = SystemColors.Window
            CBTer.BackColor = SystemColors.Window
            CBDaf.BackColor = SystemColors.Window
            TCost.BackColor = SystemColors.Window

            ' 11. დამატების ღილაკის დამალვა, რადგან ველები ცარიელია
            BtnAdd.Visible = False

            ' 12. სივრცეების ინიციალიზაცია - თავიდან დავაინიციალიზიროთ ღილაკების ფერები
            InitializeSpaceButtons()

            ' 13. ფორმის ვალიდაციის განახლება
            ValidateFormInputs()

            Debug.WriteLine("BtnClear_Click: ფორმის ველების გასუფთავება დასრულდა")

            ' 14. მომხმარებლის ინფორმირება
            MessageBox.Show("ფორმის ველები გასუფთავდა", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"BtnClear_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ფორმის გასუფთავებისას დაფიქსირდა შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' ValueChanged ივენთი აღარ გამოიყენება, მაგრამ რჩება დიზაინერში მიბმული
    ''' შევქმნით ცარიელ მეთოდს, რომელიც არაფერს აკეთებს
    ''' </summary>
    Private Sub DTP1_ValueChanged(sender As Object, e As EventArgs) Handles DTP1.ValueChanged
        ' არაფერს აკეთებს, რეალური ფუნქცონალი CloseUp-შია
    End Sub
    ''' <summary>
    ''' BtnAdd ღილაკზე დაჭერის ჰენდლერი - ახალი სესიის დამატება ან არსებული სესიის რედაქტირება
    ''' შესწორებული ვერსია ფასის სწორი გადაცემისთვის
    ''' </summary>
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Try
            ' დებაგ ინფორმაცია
            Debug.WriteLine($"BtnAdd_Click: დაიწყო ოპერაცია - რეჟიმი: {If(_isAddMode, "დამატება", "რედაქტირება")}")

            ' 1. შემოწმება ფორმის ვალიდურობის - ValidateFormInputs გამოძახება
            ValidateFormInputs()

            ' BtnAdd-ის ხილვადობა განისაზღვრება ValidateFormInputs მეთოდში
            ' თუ BtnAdd არ არის ხილული, ფორმა არავალიდურია
            If Not BtnAdd.Visible Then
                MessageBox.Show("გთხოვთ შეავსოთ ყველა სავალდებულო ველი", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 2. მოვამზადოთ მონაცემები Google Sheets-ისთვის
            Dim rowData As New List(Of Object)()

            ' ა) ID (A სვეტი)
            rowData.Add(LN.Text) ' სესიის ID

            ' ბ) შევსების თარიღი (B სვეტი)
            rowData.Add(LNow.Text) ' მიმდინარე თარიღი და დრო

            ' გ) ავტორი (C სვეტი)
            rowData.Add(LAutor.Text) ' მომხმარებლის ელფოსტა

            ' დ) ბენეფიციარის სახელი (D სვეტი)
            rowData.Add(CBBeneName.Text)

            ' ე) ბენეფიციარის გვარი (E სვეტი)
            rowData.Add(CBBeneSurname.Text)

            ' ვ) თარიღი და დრო (F სვეტი)
            ' თარიღის და დროის კომბინაცია შესაბამის ფორმატში
            Dim selectedDate As String = DTP1.Value.ToString("dd.MM.yyyy")
            Dim hour As String = THour.Text.PadLeft(2, "0"c)
            Dim minute As String = TMin.Text.PadLeft(2, "0"c)
            rowData.Add($"{selectedDate} {hour}:{minute}")

            ' ზ) ხანგრძლივობა (G სვეტი)
            Dim duration As Integer = 60 ' ნაგულისხმევი მნიშვნელობა
            Integer.TryParse(TDur.Text, duration)
            rowData.Add(duration)

            ' თ) ჯგუფური (H სვეტი)
            rowData.Add(CBGroup.Checked)

            ' ი) თერაპევტი (I სვეტი)
            rowData.Add(CBPer.Text)

            ' კ) თერაპიის ტიპი (J სვეტი)
            rowData.Add(CBTer.Text)

            ' ლ) სივრცე (K სვეტი)
            Dim selectedSpace As String = GetSelectedSpace()
            rowData.Add(selectedSpace)

            ' მ) ფასი (L სვეტი) - გასწორებული ვერსია
            Dim priceText As String = TCost.Text.Trim().Replace(",", ".")
            Dim price As Decimal = 0

            ' დებაგისთვის ვბეჭდავთ საწყის მნიშვნელობას
            Debug.WriteLine($"BtnAdd_Click: ფასის საწყისი მნიშვნელობა TCost.Text = '{TCost.Text}'")
            Debug.WriteLine($"BtnAdd_Click: ფასის დამუშავებული ტექსტი priceText = '{priceText}'")

            ' მკაცრი პარსინგი უცვლელი კულტურით
            If Decimal.TryParse(priceText, NumberStyles.Any, CultureInfo.InvariantCulture, price) Then
                Debug.WriteLine($"BtnAdd_Click: ფასის პარსინგი წარმატებულია: {price}")
            Else
                ' თუ პარსინგი ვერ მოხერხდა, ჩავატაროთ მეტი დებაგინგი
                Debug.WriteLine($"BtnAdd_Click: ფასის პარსინგი ვერ მოხერხდა!")

                ' მეტად გავმართოთ ტექსტი და კიდევ ერთხელ ვცადოთ
                priceText = priceText.Trim().Replace(" ", "")
                If Decimal.TryParse(priceText, NumberStyles.Any, CultureInfo.InvariantCulture, price) Then
                    Debug.WriteLine($"BtnAdd_Click: ფასის მეორადი პარსინგი წარმატებულია: {price}")
                Else
                    ' თუ მაინც ვერ მოხერხდა, ვიყენებთ 0-ს და ვბეჭდავთ გაფრთხილებას
                    Debug.WriteLine($"BtnAdd_Click: ფასის მეორადი პარსინგიც ვერ მოხერხდა! ვიყენებთ 0-ს")
                    price = 0
                End If
            End If

            ' ფასის მნიშვნელობა აქ უნდა დაემატოს როგორც რიცხვითი მნიშვნელობა
            rowData.Add(price)

            ' ნ) შესრულების სტატუსი (M სვეტი)
            Dim status As String = GetSelectedStatus()
            rowData.Add(status)

            ' ო) დაფინანსების პროგრამა (N სვეტი)
            rowData.Add(CBDaf.Text)

            ' პ) კომენტარი (O სვეტი)
            rowData.Add(TCom.Text)

            ' მთლიანი მწკრივის მნიშვნელობების ლოგირება დებაგისთვის
            Dim rowDataDebug As New StringBuilder("BtnAdd_Click: მწკრივის მნიშვნელობები: ")
            For i As Integer = 0 To rowData.Count - 1
                rowDataDebug.Append($"[{i}]={rowData(i)}, ")
            Next
            Debug.WriteLine(rowDataDebug.ToString())

            ' 3. ოპერაციის შესრულება - დამატება ან რედაქტირება
            If _isAddMode Then
                ' ახალი სესიის დამატება
                dataService.AppendData("DB-Schedule!A:O", rowData)
                Debug.WriteLine($"BtnAdd_Click: დაემატა ახალი სესია ID={LN.Text}")
                MessageBox.Show("სესია წარმატებით დაემატა", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' არსებული სესიის რედაქტირება

                ' ჯერ მოვძებნოთ შესაბამისი მწკრივი DB-Schedule ფურცელში
                Dim scheduleData = dataService.GetData("DB-Schedule!A2:A")
                Dim rowIndex As Integer = -1

                If scheduleData IsNot Nothing Then
                    For i As Integer = 0 To scheduleData.Count - 1
                        Try
                            If scheduleData(i).Count > 0 AndAlso
                           scheduleData(i)(0).ToString() = LN.Text Then
                                ' ვიპოვეთ შესაბამისი ID
                                rowIndex = i + 2 ' +2 იმიტომ, რომ (1) ინდექსები 0-დან იწყება და (2) პირველი რიგი სათაურებია
                                Exit For
                            End If
                        Catch ex As Exception
                            ' გავაგრძელოთ შემდეგი მწკრივით
                            Continue For
                        End Try
                    Next
                End If

                If rowIndex > 0 Then
                    ' განვაახლოთ მწკრივი
                    dataService.UpdateData($"DB-Schedule!A{rowIndex}:O{rowIndex}", rowData)
                    Debug.WriteLine($"BtnAdd_Click: განახლდა სესია ID={LN.Text}, რიგი={rowIndex}")
                    MessageBox.Show("სესია წარმატებით განახლდა", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' ვერ ვიპოვეთ შესაბამისი მწკრივი
                    Debug.WriteLine($"BtnAdd_Click: ვერ მოიძებნა სესია ID={LN.Text}")
                    MessageBox.Show("სესიის განახლება ვერ მოხერხდა - შესაბამისი ID ვერ მოიძებნა", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            ' 4. ფორმის დახურვა წარმატებული ოპერაციის შემდეგ
            Me.Close()

        Catch ex As Exception
            Debug.WriteLine($"BtnAdd_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"სესიის დამატება/რედაქტირება ვერ მოხერხდა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' აბრუნებს არჩეულ სივრცეს ფორმიდან
    ''' </summary>
    Private Function GetSelectedSpace() As String
        ' ვეძებთ არჩეულ ღილაკს (ცისფერი ფონით)
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            If btn.Name.StartsWith("BTNS") AndAlso btn.BackColor = Color.FromArgb(173, 216, 230) Then
                Return btn.Text
            End If
        Next

        ' არცერთი სივრცე არ არის არჩეული
        Return ""
    End Function

    ''' <summary>
    ''' აბრუნებს არჩეულ სტატუსს ფორმიდან
    ''' </summary>
    Private Function GetSelectedStatus() As String
        ' შევამოწმოთ არის თუ არა სამომავლო სესია
        Dim selectedDateTime As DateTime = GetSelectedDateTime()

        If selectedDateTime > DateTime.Now Then
            ' სამომავლო სესია ავტომატურად "დაგეგმილი" სტატუსით
            Return "დაგეგმილი"
        End If

        ' წარსული თარიღისთვის ვამოწმებთ არჩეულ რადიობუტონს
        For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
            If rb.Name.StartsWith("RB") AndAlso rb.Checked Then
                Return rb.Text
            End If
        Next

        ' ნაგულისხმევი სტატუსი, თუ არცერთი არ არის არჩეული
        Return "დაგეგმილი"
    End Function

    ''' <summary>
    ''' BtnAddBene ღილაკზე დაჭერის ჰენდლერი - ახალი ბენეფიციარის დამატება
    ''' </summary>
    Private Sub BtnAddBene_Click(sender As Object, e As EventArgs) Handles BtnAddBene.Click
        Try
            Debug.WriteLine("BtnAddBene_Click: ბენეფიციარის დამატების დაწყება")

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ბენეფიციარის დამატების ფორმის გახსნა
            Using addBeneForm As New AddBeneForm(dataService)
                Dim result = addBeneForm.ShowDialog()

                ' შევამოწმოთ შედეგი
                If result = DialogResult.OK AndAlso addBeneForm.IsSuccess Then
                    Debug.WriteLine($"BtnAddBene_Click: ბენეფიციარი დაემატა - {addBeneForm.AddedName} {addBeneForm.AddedSurname}")

                    ' განვაახლოთ ბენეფიციარების ჩამონათვალი
                    LoadBeneNames()

                    ' დავაყენოთ ახლად დამატებული ბენეფიციარი
                    If CBBeneName.Items.Count > 0 Then
                        ' მოვძებნოთ დამატებული ბენეფიციარის სახელი ჩამონათვალში
                        Dim nameIndex As Integer = -1
                        For i As Integer = 0 To CBBeneName.Items.Count - 1
                            If String.Equals(CBBeneName.Items(i).ToString(), addBeneForm.AddedName, StringComparison.OrdinalIgnoreCase) Then
                                nameIndex = i
                                Exit For
                            End If
                        Next

                        ' თუ ვიპოვეთ, დავაყენოთ ეს სახელი
                        If nameIndex > 0 Then
                            CBBeneName.SelectedIndex = nameIndex

                            ' ცოტა დავიცადოთ გვარების ჩატვირთვას
                            Application.DoEvents()

                            ' ახლა ვეძებთ გვარს
                            For i As Integer = 0 To CBBeneSurname.Items.Count - 1
                                If String.Equals(CBBeneSurname.Items(i).ToString(), addBeneForm.AddedSurname, StringComparison.OrdinalIgnoreCase) Then
                                    CBBeneSurname.SelectedIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    ' შევამოწმოთ ბენეფიციარის მდგომარეობა
                    CheckBeneficiaryAvailability()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"BtnAddBene_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბენეფიციარის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BtnAddPer ღილაკზე დაჭერის ივენთი - თერაპევტის დამატება
    ''' ბენეფიციარის ლოგიკის გამოყენებით კომბობოქსის განახლებისთვის
    ''' </summary>
    Private Sub BtnAddPer_Click(sender As Object, e As EventArgs) Handles BtnAddPer.Click
        Try
            Debug.WriteLine("BtnAddPer_Click: თერაპევტის დამატების დაწყება")

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' თერაპევტის დამატების ფორმის გახსნა
            Using addPerForm As New AddPerForm(dataService)
                Dim result = addPerForm.ShowDialog()

                ' შევამოწმოთ შედეგი
                If result = DialogResult.OK AndAlso addPerForm.IsSuccess Then
                    Debug.WriteLine($"BtnAddPer_Click: თერაპევტი დაემატა - {addPerForm.AddedName} {addPerForm.AddedSurname}")

                    ' განვაახლოთ თერაპევტების ჩამონათვალი
                    LoadTherapists()

                    ' დავაყენოთ ახლად დამატებული თერაპევტი
                    If CBPer.Items.Count > 0 Then
                        ' ვეძებთ ახლად დამატებული თერაპევტის სრულ სახელს ჩამონათვალში
                        Dim fullName As String = $"{addPerForm.AddedName} {addPerForm.AddedSurname}"
                        Dim therapistIndex As Integer = -1

                        For i As Integer = 0 To CBPer.Items.Count - 1
                            If String.Equals(CBPer.Items(i).ToString(), fullName, StringComparison.OrdinalIgnoreCase) Then
                                therapistIndex = i
                                Exit For
                            End If
                        Next

                        ' თუ ვიპოვეთ, დავაყენოთ ეს თერაპევტი არჩეულად
                        If therapistIndex > 0 Then
                            CBPer.SelectedIndex = therapistIndex
                            Debug.WriteLine($"BtnAddPer_Click: ახლად დამატებული თერაპევტი '{fullName}' წარმატებით აირჩა, ინდექსი {therapistIndex}")
                        Else
                            Debug.WriteLine($"BtnAddPer_Click: ახლად დამატებული თერაპევტი '{fullName}' ვერ მოიძებნა ჩამონათვალში")
                        End If
                    End If

                    ' განვაახლოთ UI
                    Application.DoEvents()

                    ' შევამოწმოთ თერაპევტის მდგომარეობა
                    CheckTherapistAvailability()

                    ' შეტყობინება მომხმარებლისთვის
                    MessageBox.Show($"თერაპევტი '{addPerForm.AddedName} {addPerForm.AddedSurname}' წარმატებით დაემატა",
                                 "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"BtnAddPer_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"თერაპევტის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BtnAddTer ღილაკზე დაჭერის ივენთი - თერაპიის ტიპის დამატება
    ''' </summary>
    Private Sub BtnAddTer_Click(sender As Object, e As EventArgs) Handles BtnAddTer.Click
        Try
            Debug.WriteLine("BtnAddTer_Click: თერაპიის ტიპის დამატების დაწყება")

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' თერაპიის ტიპის დამატების ფორმის გახსნა
            Using addTerForm As New AddTerForm(dataService)
                Dim result = addTerForm.ShowDialog()

                ' შევამოწმოთ შედეგი
                If result = DialogResult.OK AndAlso addTerForm.IsSuccess Then
                    Debug.WriteLine($"BtnAddTer_Click: თერაპიის ტიპი დაემატა - {addTerForm.AddedTherapy}")

                    ' მცირე დაყოვნება, რათა მონაცემები დარეგისტრირდეს Google Sheets-ში
                    System.Threading.Thread.Sleep(500)

                    ' განვაახლოთ თერაპიის ტიპების ჩამონათვალი ძალდატანებით და ახალი ტიპის მითითებით
                    'LoadTherapyTypes(True, addTerForm.AddedTherapy)

                    ' UI-ის განახლების უზრუნველყოფა
                    Application.DoEvents()

                    ' შეტყობინება მომხმარებლისთვის
                    MessageBox.Show($"თერაპიის ტიპი '{addTerForm.AddedTherapy}' წარმატებით დაემატა და აირჩა",
                              "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' ფორმის ვალიდაცია
                    ValidateFormInputs()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"BtnAddTer_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"თერაპიის ტიპის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' დაფინანსების პროგრამების ჩატვირთვა DB-Program ფურცლიდან - გაუმჯობესებული
    ''' </summary>
    ''' <param name="forceRefresh">ძალდატანებითი განახლება</param>
    ''' <param name="newProgram">ახლად დამატებული დაფინანსების პროგრამა (არასავალდებულო)</param>
    Private Sub LoadFundingPrograms(Optional forceRefresh As Boolean = False, Optional newProgram As String = "")
        Try
            Debug.WriteLine($"LoadFundingPrograms: დაიწყო ჩატვირთვა, forceRefresh={forceRefresh}, newProgram={newProgram}")

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' დაყოვნება ძალდატანებითი განახლების შემთხვევაში
            If forceRefresh Then
                ' მოიცადეთ ცოტა ხანი, რომ Google Sheets-მა მოასწროს განახლება
                System.Threading.Thread.Sleep(1000)
                Debug.WriteLine("LoadFundingPrograms: დაყოვნება 1 წამი ძალდატანებითი განახლებისთვის")
            End If

            ' B სვეტის წაკითხვა DB-Program ფურცლიდან
            Dim rows = dataService.GetData("DB-Program!B2:B")

            ' თუ მონაცემები არ არის, გამოვიდეთ ფუნქციიდან
            If rows Is Nothing OrElse rows.Count = 0 Then
                MessageBox.Show("დაფინანსების პროგრამების მონაცემები ვერ მოიძებნა", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' მიმდინარე არჩეული დაფინანსების პროგრამა (თუ არის)
            Dim currentSelectedProgram As String = ""
            If CBDaf.SelectedIndex > 0 Then
                currentSelectedProgram = CBDaf.SelectedItem.ToString()
                Debug.WriteLine($"LoadFundingPrograms: მიმდინარე არჩეული დაფინანსების პროგრამა - {currentSelectedProgram}")
            End If

            ' გავასუფთაოთ კომბობოქსი
            CBDaf.Items.Clear()

            ' დავამატოთ პირველი ელემენტი - მინიშნება
            CBDaf.Items.Add("- აირჩიეთ დაფინანსების პროგრამა -")

            ' დაფინანსების პროგრამების სია
            Dim programs As New List(Of String)
            Dim newProgramFound As Boolean = False

            For Each row In rows
                If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString().Trim()) Then
                    Dim programName = row(0).ToString().Trim()
                    programs.Add(programName)

                    ' შევამოწმოთ არის თუ არა ეს ახალი დაფინანსების პროგრამა
                    If Not String.IsNullOrEmpty(newProgram) AndAlso
                       String.Equals(programName, newProgram, StringComparison.OrdinalIgnoreCase) Then
                        newProgramFound = True
                        Debug.WriteLine($"LoadFundingPrograms: ნაპოვნია ახალი დაფინანსების პროგრამა - {programName}")
                    End If
                End If
            Next

            ' დავალაგოთ ანბანის მიხედვით
            programs.Sort()
            Debug.WriteLine($"LoadFundingPrograms: ნაპოვნია {programs.Count} დაფინანსების პროგრამა")

            ' დავამატოთ კომბობოქსში
            For Each program In programs
                CBDaf.Items.Add(program)
            Next

            ' ავირჩიოთ შესაბამისი ელემენტი
            If Not String.IsNullOrEmpty(newProgram) AndAlso newProgramFound Then
                ' ვეძებთ ახალი დაფინანსების პროგრამის ინდექსს
                For i As Integer = 0 To CBDaf.Items.Count - 1
                    If String.Equals(CBDaf.Items(i).ToString(), newProgram, StringComparison.OrdinalIgnoreCase) Then
                        CBDaf.SelectedIndex = i
                        Debug.WriteLine($"LoadFundingPrograms: ახალი დაფინანსების პროგრამა არჩეულია - ინდექსი {i}")
                        Exit For
                    End If
                Next
            ElseIf Not String.IsNullOrEmpty(currentSelectedProgram) Then
                ' ვცდილობთ წინა არჩეული დაფინანსების პროგრამის ხელახლა არჩევას
                For i As Integer = 0 To CBDaf.Items.Count - 1
                    If String.Equals(CBDaf.Items(i).ToString(), currentSelectedProgram, StringComparison.OrdinalIgnoreCase) Then
                        CBDaf.SelectedIndex = i
                        Debug.WriteLine($"LoadFundingPrograms: წინა დაფინანსების პროგრამა კვლავ არჩეულია - ინდექსი {i}")
                        Exit For
                    End If
                Next
            Else
                ' ავირჩიოთ პირველი ელემენტი (მინიშნება)
                CBDaf.SelectedIndex = 0
                Debug.WriteLine("LoadFundingPrograms: არჩეულია პირველი ელემენტი (მინიშნება)")
            End If

            ' ძალდატანებითი განახლება UI-სთვის
            CBDaf.Refresh()
            Application.DoEvents()

            Debug.WriteLine("LoadFundingPrograms: დაფინანსების პროგრამების ჩატვირთვა დასრულდა წარმატებით")

        Catch ex As Exception
            Debug.WriteLine($"LoadFundingPrograms: შეცდომა - {ex.Message}")
            MessageBox.Show($"შეცდომა დაფინანსების პროგრამების ჩატვირთვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' BtnAddDaf ღილაკზე დაჭერის ივენთი - დაფინანსების პროგრამის დამატება
    ''' </summary>
    Private Sub BtnAddDaf_Click(sender As Object, e As EventArgs) Handles BtnAddDaf.Click
        Try
            Debug.WriteLine("BtnAddDaf_Click: დაფინანსების პროგრამის დამატების დაწყება")

            ' შევამოწმოთ dataService
            If dataService Is Nothing Then
                MessageBox.Show("მონაცემთა სერვისი არ არის ინიციალიზებული", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' დაფინანსების პროგრამის დამატების ფორმის გახსნა
            Using addDafForm As New AddDafForm(dataService)
                Dim result = addDafForm.ShowDialog()

                ' შევამოწმოთ შედეგი
                If result = DialogResult.OK AndAlso addDafForm.IsSuccess Then
                    Debug.WriteLine($"BtnAddDaf_Click: დაფინანსების პროგრამა დაემატა - {addDafForm.AddedProgram}")

                    ' მცირე დაყოვნება, რათა მონაცემები დარეგისტრირდეს Google Sheets-ში
                    System.Threading.Thread.Sleep(500)

                    ' განვაახლოთ დაფინანსების პროგრამების ჩამონათვალი ძალდატანებით და ახალი პროგრამის მითითებით
                    LoadFundingPrograms(True, addDafForm.AddedProgram)

                    ' UI-ის განახლების უზრუნველყოფა
                    Application.DoEvents()

                    ' შეტყობინება მომხმარებლისთვის
                    MessageBox.Show($"დაფინანსების პროგრამა '{addDafForm.AddedProgram}' წარმატებით დაემატა და აირჩა",
                              "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' ფორმის ვალიდაცია
                    ValidateFormInputs()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"BtnAddDaf_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"დაფინანსების პროგრამის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class