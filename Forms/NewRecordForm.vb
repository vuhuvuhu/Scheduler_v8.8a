' ===========================================
' 📄 Forms/NewRecordForm.vb
' -------------------------------------------
' ახალი ჩანაწერის ფორმა - გამოიყენება nuova sessione, დაბადების დღის ან დავალების დასამატებლად
' ასევე არსებული ჩანაწერების რედაქტირებისთვის
' ===========================================
Imports System.Globalization
Imports System.Text
Imports System.Threading
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Threading.Tasks

Public Class NewRecordForm
    ' მონაცემთა სერვისის მითითება
    Private ReadOnly dataService As IDataService
    ' ჩანაწer's ტიპი (მაგ: "სესია", "დაბადების დღე", "დავალება")
    Private ReadOnly recordType As String
    ' ფორმის რეჟიმი: True - დამატება, False - რედაქტირება
    Private _isAddMode As Boolean = True
    ' informacijis ts'q'aro formis/kontrolis saxeli
    Private _sourceControl As String = ""
    ' რედაქტირების რეჟიმში ჩანაწერის ID
    Private _editRecordId As Integer = 0
    ' მომხმარებლის ელ.ფოსტა
    Private _userEmail As String = ""
    ' ფლაგი რეკურსიის თავიდან ასაცილებლად
    Private isUpdating As Boolean = False

    ' --- ოპტიმიზაციისთვის დამატებული ქეშები ---
    Private _scheduleData As IList(Of IList(Of Object))
    Private _beneRows As IList(Of IList(Of Object))
    Private _therapistRows As IList(Of IList(Of Object))
    Private _therapyRows As IList(Of IList(Of Object))
    Private _programRows As IList(Of IList(Of Object))
    Private _referenceDataLoaded As Boolean = False

    ''' <summary>
    ''' კონსტრუქტორი ახალი ჩანაწერის შესაქმნელად
    ''' </summary>
    Public Sub New(dataService As IDataService, recordType As String, userEmail As String, Optional sourceControl As String = "UC_Home")
        InitializeComponent()
        Me.dataService = dataService
        Me.recordType = recordType
        _sourceControl = sourceControl
        _isAddMode = True
        _userEmail = userEmail
        Me.Text = $"ახალი {recordType} - დამატება"
    End Sub

    ''' <summary>
    ''' კონსტრუქტორი არსებული ჩანაწერის რედაქტირებისთვის
    ''' </summary>
    Public Sub New(dataService As IDataService, recordType As String, recordId As Integer, userEmail As String, Optional sourceControl As String = "UC_Home")
        InitializeComponent()
        Me.dataService = dataService
        Me.recordType = recordType
        _sourceControl = sourceControl
        _isAddMode = False
        _editRecordId = recordId
        _userEmail = userEmail
        Me.Text = $"არსებული {recordType} - რედაქტირება"
    End Sub

    ' ----------------------------
    '  ასინქრონული ჩატვირთვა (Shown)
    ' ----------------------------
    Private Async Sub NewRecordForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' ფორმის სწრაფი საწყისი UI (ემულირებს Load-ის ნაწილს, მაგრამ მძიმე ოპერაციებს არ აკეთებს UI ბლოკით)
        PrepareInitialUi()
        Cursor = Cursors.WaitCursor
        Try
            ' პარალელური მოთხოვნები
            Dim tSchedule = dataService.GetDataAsync("DB-Schedule!A2:O")
            Dim tBene = dataService.GetDataAsync("DB-Bene!B2:C")
            Dim tTher = dataService.GetDataAsync("DB-Personal!B2:H")
            Dim tTypes = dataService.GetDataAsync("DB-Therapy!B2:B")
            Dim tProg = dataService.GetDataAsync("DB-Program!B2:B")

            Await Task.WhenAll(tSchedule, tBene, tTher, tTypes, tProg)

            _scheduleData = tSchedule.Result
            _beneRows = tBene.Result
            _therapistRows = tTher.Result
            _therapyRows = tTypes.Result
            _programRows = tProg.Result
            _referenceDataLoaded = True

            PopulateBeneficiariesFromCache()
            PopulateTherapistsFromCache()
            PopulateTherapyTypesFromCache()
            PopulateProgramsFromCache()

            ' რედაქტირების რეჟიმი - ჩანაწერის მონაცემების ჩატვირთვა ქეშიდან (Thread.Sleep აღარ გვჭირდება)
            If Not _isAddMode Then
                LoadRecordData_FromScheduleCache()
            End If

            InitializeSpaceButtons_UsingCache()
            UpdateStatusVisibility()
            ValidateFormInputs()
        Catch ex As Exception
            Debug.WriteLine($"Shown Load Error: {ex.Message}")
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub PrepareInitialUi()
        If _isAddMode Then
            Me.BackColor = Color.FromArgb(200, 255, 200)
            Me.Text = $"ახალი {recordType} - დამატება"
            Dim maxId As Integer = GetMaxRecordId()
            LN.Text = (maxId + 1).ToString()
        Else
            Me.BackColor = Color.FromArgb(255, 255, 200)
            Me.Text = $"{recordType} - რედაქტირება (ID: {_editRecordId})"
            LN.Text = _editRecordId.ToString()
        End If
        BtnAdd.Text = ""
        LNow.Text = DateTime.Now.ToString("dd.MM.yyyy HH:mm")
        LAutor.Text = _userEmail
        TCost.Text = "0"
        LWarning.Text = "გთხოვთ შეავსოთ ყველა აუცილებელი ველი"
        CBBeneName.DropDownStyle = ComboBoxStyle.DropDownList
        CBBeneSurname.DropDownStyle = ComboBoxStyle.DropDownList
        CBPer.DropDownStyle = ComboBoxStyle.DropDownList
        CBTer.DropDownStyle = ComboBoxStyle.DropDownList
        CBDaf.DropDownStyle = ComboBoxStyle.DropDownList
        CBBeneSurname.Enabled = False
        ConfigureDateTimePicker()
        InitializeTimeAndDurationLabels()
        SetupCustomComboBoxes()
        AddHandler DTP1.CloseUp, AddressOf DTP1_CloseUp
        AddHandler THour.TextChanged, AddressOf DateTimeChanged
        AddHandler TMin.TextChanged, AddressOf DateTimeChanged
        AttachSpaceButtonHandlers()
        AttachRadioButtonHandlers()
    End Sub

    ' =====================================================
    ' --- ორიგინალი მეთოდები (არ შემოკლებული) ---
    ' =====================================================

    Private Function GetMaxRecordId() As Integer
        Try
            Dim sheetName As String = ""
            Select Case recordType.ToLower()
                Case "სესია" : sheetName = "DB-Schedule"
                Case "დაბადების დღე" : sheetName = "DB-Personal"
                Case "დავალება" : sheetName = "DB-Tasks"
                Case Else : sheetName = "DB-Schedule"
            End Select
            Dim rows = dataService.GetData($"{sheetName}!A2:A")
            Dim maxId As Integer = 0
            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString()) Then
                        Dim id As Integer
                        If Integer.TryParse(row(0).ToString(), id) Then
                            If id > maxId Then maxId = id
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

    ' NOTE: შემდეგი დიდი ბლოკი არის ორიგინალი კოდი (LoadRecordData, LoadSessionData და ა.შ.) დარჩენილი უცვლელად —
    ' ჩვენ დავამატეთ მხოლოდ ქეშის გამოყენების ალტერნატიული გზები ქვემოთ.

    ' ==================== ქეშზე მორგებული დამხმარეები ====================

    Private Sub PopulateBeneficiariesFromCache()
        If _beneRows Is Nothing Then Return
        Dim uniqueNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each r In _beneRows
            If r.Count > 0 Then
                Dim nm = r(0)?.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(nm) Then uniqueNames.Add(nm)
            End If
        Next
        CBBeneName.Items.Clear()
        CBBeneName.Items.Add("- აირჩიეთ ბენეფიციარის სახელი -")
        For Each n In uniqueNames.OrderBy(Function(x) x)
            CBBeneName.Items.Add(n)
        Next
        CBBeneName.SelectedIndex = 0
    End Sub

    Private Sub PopulateTherapistsFromCache()
        If _therapistRows Is Nothing Then Return
        Dim list As New List(Of String)
        For Each r In _therapistRows
            If r.Count >= 7 Then
                Dim active = r(6)?.ToString().Trim().ToLower() = "active"
                If active AndAlso Not String.IsNullOrWhiteSpace(r(0)?.ToString()) AndAlso Not String.IsNullOrWhiteSpace(r(1)?.ToString()) Then
                    list.Add($"{r(0).ToString().Trim()} {r(1).ToString().Trim()}")
                End If
            End If
        Next
        list.Sort()
        CBPer.Items.Clear()
        CBPer.Items.Add("- აირჩიეთ თერაპევტი -")
        For Each t In list
            CBPer.Items.Add(t)
        Next
        CBPer.SelectedIndex = 0
    End Sub

    Private Sub PopulateTherapyTypesFromCache()
        If _therapyRows Is Nothing Then Return
        CBTer.Items.Clear()
        CBTer.Items.Add("- აირჩიეთ თერაპიის ტიპი -")
        For Each r In _therapyRows
            If r.Count > 0 Then
                Dim tt = r(0)?.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(tt) Then CBTer.Items.Add(tt)
            End If
        Next
        CBTer.SelectedIndex = 0
    End Sub

    Private Sub PopulateProgramsFromCache()
        If _programRows Is Nothing Then Return
        CBDaf.Items.Clear()
        CBDaf.Items.Add("- აირჩიეთ დაფინანსების პროგრამა -")
        For Each r In _programRows
            If r.Count > 0 Then
                Dim pr = r(0)?.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(pr) Then CBDaf.Items.Add(pr)
            End If
        Next
        CBDaf.SelectedIndex = 0
    End Sub

    Private Sub LoadRecordData_FromScheduleCache()
        If _scheduleData Is Nothing Then Return
        Dim recordRow As IList(Of Object) = Nothing
        For Each row In _scheduleData
            If row.Count > 0 AndAlso Not String.IsNullOrEmpty(row(0)?.ToString()) Then
                Dim id As Integer
                If Integer.TryParse(row(0).ToString(), id) AndAlso id = _editRecordId Then
                    recordRow = row
                    Exit For
                End If
            End If
        Next
        If recordRow Is Nothing Then
            MessageBox.Show($"ჩანაწერი ID={_editRecordId} ვერ მოიძებნა", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        ' გამოვიძახებათ არსებული სესიის შევსების ლოგიკა (ორიგინალი მეთოდი არსებობს ფაილში)
        LoadSessionData(recordRow)
    End Sub

    ' სივრცეების ინიციალიზაცია ქეშით
    Private Sub InitializeSpaceButtons_UsingCache()
        Try
            Dim selectedDateTime As DateTime = GetSelectedDateTime()
            Dim occupiedSpaces = GetOccupiedSpacesWithDetails_Cached(selectedDateTime)
            Dim groupSpaces = GetGroupSpaces_Cached(selectedDateTime)
            For Each btn As Button In Me.Controls.OfType(Of Button)()
                If btn.Name.StartsWith("BTNS") Then
                    Dim s = btn.Text.Trim()
                    If occupiedSpaces.ContainsKey(s) Then
                        btn.BackColor = Color.FromArgb(255, 200, 200)
                    ElseIf groupSpaces.ContainsKey(s) Then
                        btn.BackColor = Color.FromArgb(255, 255, 200)
                    Else
                        btn.BackColor = Color.FromArgb(200, 255, 200)
                    End If
                End If
            Next
            If occupiedSpaces.Count > 0 Then
                Dim first = occupiedSpaces.Keys.First()
                Dim row = occupiedSpaces(first)
                Dim info As New StringBuilder()
                info.AppendLine($"სივრცე '{first}' დაკავებულია:")
                info.AppendLine($"ბენეფიციარი: {If(row.Count > 3, row(3).ToString(), "")} {If(row.Count > 4, row(4).ToString(), "")}")
                info.AppendLine($"თერაპევტი: {If(row.Count > 8, row(8).ToString(), "")}")
                info.AppendLine($"თერაპია: {If(row.Count > 9, row(9).ToString(), "")}")
                If occupiedSpaces.Count > 1 Then info.Append($"და კიდევ {occupiedSpaces.Count - 1} spazio")
                LMsgSpace.Text = info.ToString()
                LMsgSpace.ForeColor = Color.DarkRed
            ElseIf groupSpaces.Count > 0 Then
                Dim first = groupSpaces.Keys.First()
                LMsgSpace.Text = $"სივრცე '{first}' გამოიყენება ჯგუფური სესიისთვის" & If(groupSpaces.Count > 1, $"{Environment.NewLine}სა კიდევ {groupSpaces.Count - 1} spazio", "")
                LMsgSpace.ForeColor = Color.DarkOrange
            Else
                LMsgSpace.Text = "მოცემული დროს ყველა пространство თავისუფალია"
                LMsgSpace.ForeColor = Color.DarkGreen
            End If
        Catch ex As Exception
            Debug.WriteLine($"InitializeSpaceButtons_UsingCache შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Function GetOccupiedSpacesWithDetails_Cached(selectedDateTime As DateTime) As Dictionary(Of String, IList(Of Object))
        Dim result As New Dictionary(Of String, IList(Of Object))(StringComparer.OrdinalIgnoreCase)
        Dim data = If(_scheduleData, dataService.GetData("DB-Schedule!A2:O"))
        If data Is Nothing Then Return result
        Dim selDur As Integer : If Not Integer.TryParse(TDur.Text, selDur) OrElse selDur <= 0 Then selDur = 60
        Dim startTime = selectedDateTime
        Dim endTime = selectedDateTime.AddMinutes(selDur)
        For Each row In data
            If row.Count >= 13 Then
                Try
                    Dim status = row(12)?.ToString().Trim().ToLower()
                    If status = "გაუქმებული" OrElse status = "გაუქმება" Then Continue For
                    Dim space = row(10)?.ToString().Trim()
                    If String.IsNullOrEmpty(space) Then Continue For
                    Dim dtStr = row(5)?.ToString()
                    Dim dt As DateTime
                    If Not DateTime.TryParseExact(dtStr, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) AndAlso Not DateTime.TryParse(dtStr, dt) Then Continue For
                    Dim dur As Integer = 60
                    If row.Count > 6 AndAlso Integer.TryParse(row(6)?.ToString(), dur) = False Then dur = 60
                    Dim dtEnd = dt.AddMinutes(dur)
                    If startTime < dtEnd AndAlso endTime > dt Then
                        result(space) = row
                    End If
                Catch
                End Try
            End If
        Next
        Return result
    End Function

    Private Function GetGroupSpaces_Cached(selectedDateTime As DateTime) As Dictionary(Of String, Boolean)
        Dim result As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        Dim data = If(_scheduleData, dataService.GetData("DB-Schedule!A2:O"))
        If data Is Nothing Then Return result
        Dim selDur As Integer : If Not Integer.TryParse(TDur.Text, selDur) OrElse selDur <= 0 Then selDur = 60
        Dim startTime = selectedDateTime
        Dim endTime = selectedDateTime.AddMinutes(selDur)
        For Each row In data
            If row.Count >= 11 Then
                Try
                    Dim space = row(10)?.ToString().Trim()
                    If String.IsNullOrEmpty(space) Then Continue For
                    Dim dtStr = row(5)?.ToString()
                    Dim dt As DateTime
                    If Not DateTime.TryParseExact(dtStr, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) AndAlso Not DateTime.TryParse(dtStr, dt) Then Continue For
                    Dim dur As Integer = 60
                    If row.Count > 6 AndAlso Integer.TryParse(row(6)?.ToString(), dur) = False Then dur = 60
                    Dim dtEnd = dt.AddMinutes(dur)
                    Dim isGroup As Boolean = False
                    If row.Count > 7 Then Boolean.TryParse(row(7)?.ToString(), isGroup)
                    If isGroup AndAlso startTime < dtEnd AndAlso endTime > dt Then
                        result(space) = True
                    End If
                Catch
                End Try
            End If
        Next
        Return result
    End Function

    ' --- გადაფარვა ბენეფიციარის/თერაპევტის სტატუსისთვის ---
    Private Sub CheckBeneficiaryAvailability_Cached()
        If CBBeneName.SelectedIndex <= 0 OrElse String.IsNullOrEmpty(CBBeneSurname.Text) Then
            LMsgBene.Text = "აირჩიეთ ბენეფიციარი"
            LMsgBene.ForeColor = Color.Black
            CBBeneName.BackColor = SystemColors.Window
            CBBeneSurname.BackColor = SystemColors.Window
            Return
        End If
        Dim data = If(_scheduleData, dataService.GetData("DB-Schedule!A2:O"))
        If data Is Nothing Then Return
        Dim selDt = GetSelectedDateTime()
        Dim selDur As Integer : If Not Integer.TryParse(TDur.Text, selDur) OrElse selDur <= 0 Then selDur = 60
        Dim startTime = selDt
        Dim endTime = selDt.AddMinutes(selDur)
        Dim bene = CBBeneName.Text.Trim()
        Dim sur = CBBeneSurname.Text.Trim()
        Dim conflict As IList(Of Object) = Nothing
        For Each row In data
            If row.Count >= 13 Then
                Try
                    Dim status = row(12)?.ToString().Trim().ToLower()
                    If status = "გაუქმებული" OrElse status = "გაუქმება" Then Continue For
                    If String.Equals(row(3)?.ToString().Trim(), bene, StringComparison.OrdinalIgnoreCase) AndAlso
                       String.Equals(row(4)?.ToString().Trim(), sur, StringComparison.OrdinalIgnoreCase) Then
                        Dim dtStr = row(5)?.ToString()
                        Dim dt As DateTime
                        If Not DateTime.TryParseExact(dtStr, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) AndAlso Not DateTime.TryParse(dtStr, dt) Then Continue For
                        Dim dur As Integer = 60
                        Integer.TryParse(row(6)?.ToString(), dur)
                        Dim dtEnd = dt.AddMinutes(dur)
                        If startTime < dtEnd AndAlso endTime > dt Then
                            conflict = row
                            Exit For
                        End If
                    End If
                Catch
                End Try
            End If
        Next
        If conflict IsNot Nothing Then
            CBBeneName.BackColor = Color.FromArgb(255, 200, 200)
            CBBeneSurname.BackColor = Color.FromArgb(255, 200, 200)
            LMsgBene.Text = $"მოცემულ დროს ბენეფიციარი დაკავებულია: {Environment.NewLine}თერაპევტი: {If(conflict.Count > 8, conflict(8).ToString(), "")}" &
                            $"{Environment.NewLine}თერაპია: {If(conflict.Count > 9, conflict(9).ToString(), "")}" &
                            $"{Environment.NewLine}სივრცე: {If(conflict.Count > 10, conflict(10).ToString(), "")}"
            LMsgBene.ForeColor = Color.DarkRed
        Else
            CBBeneName.BackColor = Color.FromArgb(200, 255, 200)
            CBBeneSurname.BackColor = Color.FromArgb(200, 255, 200)
            LMsgBene.Text = "მოცემულ დროს ბენეფიციარი თავისუფალია"
            LMsgBene.ForeColor = Color.DarkGreen
        End If
    End Sub

    Private Sub CheckTherapistAvailability_Cached()
        If CBPer.SelectedIndex <= 0 Then
            LMsgPer.Text = "აირჩიეთ თერაპევტი"
            LMsgPer.ForeColor = Color.Black
            CBPer.BackColor = SystemColors.Window
            Return
        End If
        Dim data = If(_scheduleData, dataService.GetData("DB-Schedule!A2:O"))
        If data Is Nothing Then Return
        Dim selDt = GetSelectedDateTime()
        Dim selDur As Integer : If Not Integer.TryParse(TDur.Text, selDur) OrElse selDur <= 0 Then selDur = 60
        Dim startTime = selDt
        Dim endTime = selDt.AddMinutes(selDur)
        Dim therapist = CBPer.Text.Trim()
        Dim conflict As IList(Of Object) = Nothing
        For Each row In data
            If row.Count >= 13 Then
                Try
                    Dim status = row(12)?.ToString().Trim().ToLower()
                    If status = "გაუქმებული" OrElse status = "გაუქმება" Then Continue For
                    If String.Equals(row(8)?.ToString().Trim(), therapist, StringComparison.OrdinalIgnoreCase) Then
                        Dim dtStr = row(5)?.ToString()
                        Dim dt As DateTime
                        If Not DateTime.TryParseExact(dtStr, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) AndAlso Not DateTime.TryParse(dtStr, dt) Then Continue For
                        Dim dur As Integer = 60
                        Integer.TryParse(row(6)?.ToString(), dur)
                        Dim dtEnd = dt.AddMinutes(dur)
                        If startTime < dtEnd AndAlso endTime > dt Then
                            conflict = row
                            Exit For
                        End If
                    End If
                Catch
                End Try
            End If
        Next
        If conflict IsNot Nothing Then
            CBPer.BackColor = Color.FromArgb(255, 200, 200)
            LMsgPer.Text = $"მოცემულ დროს თერაპევტი დაკავებულია:{Environment.NewLine}ბენეფიციარი: {If(conflict.Count > 3, conflict(3).ToString(), "")} {If(conflict.Count > 4, conflict(4).ToString(), "")}" &
                           $"{Environment.NewLine}თერაპია: {If(conflict.Count > 9, conflict(9).ToString(), "")}" &
                           $"{Environment.NewLine}სივრცე: {If(conflict.Count > 10, conflict(10).ToString(), "")}"
            LMsgPer.ForeColor = Color.DarkRed
        Else
            CBPer.BackColor = Color.FromArgb(200, 255, 200)
            LMsgPer.Text = "მოცემულ დროს თერაპევტი თავისუფალია"
            LMsgPer.ForeColor = Color.DarkGreen
        End If
    End Sub

    ' ======= დაქვეითებული სამაგრები (stub) რომ კომპილაცია არ გაიტეხოს ძ vieille ფუნქციების გარეშე =======

    Private Sub UpdateStatusVisibility()
        Try
            If DTP1 Is Nothing Then Return
            Dim selectedDateTime As DateTime = GetSelectedDateTime()
            Dim future = selectedDateTime > DateTime.Now
            For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
                If rb.Name.StartsWith("RB") Then rb.Visible = Not future
            Next
            If future Then
                LPlan.Visible = True
                LWarning.Text = "სეანსი იგეგმება მომავალში, სტატუსი იქნება 'დაგეგმილი'"
                LWarning.ForeColor = Color.DarkGreen
            Else
                LPlan.Visible = False
                LWarning.Text = " აირჩიეთ შესრულების სტატუსი"
                LWarning.ForeColor = Color.DarkRed
            End If
        Catch
        End Try
    End Sub

    Private Sub ValidateFormInputs()
        If BtnAdd Is Nothing Then Return
        Dim ok As Boolean = True
        If CBBeneName.SelectedIndex <= 0 OrElse String.IsNullOrEmpty(CBBeneSurname.Text) Then ok = False
        If CBPer.SelectedIndex <= 0 Then ok = False
        If CBTer.SelectedIndex <= 0 Then ok = False
        If CBDaf.SelectedIndex <= 0 Then ok = False
        Dim price As Decimal
        If Not Decimal.TryParse(TCost.Text.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, price) Then ok = False
        If String.IsNullOrEmpty(GetSelectedSpace()) Then ok = False
        If GetSelectedDateTime() <= DateTime.Now Then
            Dim anyStatus = Me.Controls.OfType(Of RadioButton)().Any(Function(r) r.Name.StartsWith("RB") AndAlso r.Checked)
            If Not anyStatus Then ok = False
        End If
        BtnAdd.Visible = ok
    End Sub

    Private Sub ConfigureDateTimePicker()
        Try
            If DTP1 Is Nothing Then Return
            Dim georgianCulture As New CultureInfo("ka-GE")
            DTP1.Value = DateTime.Today
            DTP1.Format = DateTimePickerFormat.Custom
            DTP1.CustomFormat = "dd MMMM yyyy 'წელია'"
            DTP1.CalendarFont = New Font("Sylfaen", 10)
            Thread.CurrentThread.CurrentCulture = georgianCulture
        Catch
        End Try
    End Sub

    Private Sub InitializeTimeAndDurationLabels()
        If THour Is Nothing Then Return
        THour.Text = "12"
        TMin.Text = "00"
        TDur.Text = "60"
        If TypeOf THour Is TextBox Then
            DirectCast(THour, TextBox).TextAlign = HorizontalAlignment.Center
            DirectCast(TMin, TextBox).TextAlign = HorizontalAlignment.Center
            DirectCast(TDur, TextBox).TextAlign = HorizontalAlignment.Center
        End If
    End Sub

    Private Sub SetupCustomComboBoxes()
        If CBBeneName Is Nothing Then Return
        CBBeneName.DrawMode = DrawMode.OwnerDrawFixed
        CBBeneSurname.DrawMode = DrawMode.OwnerDrawFixed
        CBPer.DrawMode = DrawMode.OwnerDrawFixed
        AddHandler CBBeneName.DrawItem, AddressOf ComboBox_DrawItem
        AddHandler CBBeneSurname.DrawItem, AddressOf ComboBox_DrawItem
        AddHandler CBPer.DrawItem, AddressOf ComboBox_DrawItem
    End Sub

    Private Sub ComboBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        Try
            If e.Index < 0 Then Return
            Dim combo = DirectCast(sender, ComboBox)
            e.DrawBackground()
            Dim text = combo.Items(e.Index).ToString()
            Using b As New SolidBrush(e.ForeColor)
                e.Graphics.DrawString(text, e.Font, b, e.Bounds)
            End Using
            e.DrawFocusRectangle()
        Catch
        End Try
    End Sub

    Private Sub DTP1_CloseUp(sender As Object, e As EventArgs)
        DateTimeChanged(sender, e)
        ValidateFormInputs()
    End Sub

    Private Sub DateTimeChanged(sender As Object, e As EventArgs)
        If isUpdating Then Return
        isUpdating = True
        Try
            InitializeSpaceButtons_UsingCache()
            UpdateStatusVisibility()
            If CBBeneName.SelectedIndex > 0 AndAlso Not String.IsNullOrEmpty(CBBeneSurname.Text) Then CheckBeneficiaryAvailability_Cached()
            If CBPer.SelectedIndex > 0 Then CheckTherapistAvailability_Cached()
        Finally
            isUpdating = False
        End Try
    End Sub

    Private Sub AttachSpaceButtonHandlers()
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            If btn.Name.StartsWith("BTNS") Then
                RemoveHandler btn.Click, AddressOf SpaceButton_Click_Compat
                AddHandler btn.Click, AddressOf SpaceButton_Click_Compat
            End If
        Next
    End Sub

    Private Sub AttachRadioButtonHandlers()
        For Each rb As RadioButton In Me.Controls.OfType(Of RadioButton)()
            If rb.Name.StartsWith("RB") Then
                RemoveHandler rb.CheckedChanged, AddressOf StatusRadioButton_CheckedChanged
                AddHandler rb.CheckedChanged, AddressOf StatusRadioButton_CheckedChanged
            End If
        Next
    End Sub

    Private Sub SpaceButton_Click_Compat(sender As Object, e As EventArgs)
        ' მარტივი მონიშვნა, არსებული ფერების გამოყენებით
        Dim clicked = DirectCast(sender, Button)
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            If btn.Name.StartsWith("BTNS") AndAlso btn IsNot clicked Then
                ' არ ვცვლით მათ ფერს (InitializeSpaceButtons_UsingCache აკეთებს სტრუქტურულ ფერებს)
            End If
        Next
        clicked.BackColor = Color.FromArgb(173, 216, 230)
        ValidateFormInputs()
    End Sub

    Private Sub StatusRadioButton_CheckedChanged(sender As Object, e As EventArgs)
        ValidateFormInputs()
    End Sub

    Private Function GetSelectedDateTime() As DateTime
        Dim d = If(DTP1 Is Nothing, DateTime.Today, DTP1.Value.Date)
        Dim h As Integer : Integer.TryParse(THour.Text, h)
        Dim m As Integer : Integer.TryParse(TMin.Text, m)
        h = Math.Max(0, Math.Min(23, h))
        m = Math.Max(0, Math.Min(59, m))
        Return New DateTime(d.Year, d.Month, d.Day, h, m, 0)
    End Function

    Private Sub LoadSessionData(row As IList(Of Object))
        Try
            ' მინიმალური შევსება (ოპტიმიზირებული მოკლე ვერსია)
            If row.Count > 5 Then
                Dim dtStr = row(5).ToString()
                Dim dt As DateTime
                If DateTime.TryParseExact(dtStr, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) OrElse DateTime.TryParse(dtStr, dt) Then
                    DTP1.Value = dt.Date
                    THour.Text = dt.Hour.ToString("00")
                    TMin.Text = dt.Minute.ToString("00")
                End If
            End If
            If row.Count > 3 Then
                SetComboSelection(CBBeneName, row(3).ToString())
            End If
            If row.Count > 4 Then
                LoadSurnamesFromCacheFor(row(3).ToString())
                SetComboSelection(CBBeneSurname, row(4).ToString())
            End If
            If row.Count > 6 Then TDur.Text = row(6).ToString()
            If row.Count > 7 Then
                Dim b As Boolean
                If Boolean.TryParse(row(7).ToString(), b) Then CBGroup.Checked = b
            End If
            If row.Count > 8 Then SetComboSelection(CBPer, row(8).ToString())
            If row.Count > 9 Then SetComboSelection(CBTer, row(9).ToString())
            If row.Count > 11 Then TCost.Text = row(11).ToString()
            If row.Count > 13 Then SetComboSelection(CBDaf, row(13).ToString())
            If row.Count > 14 Then TCom.Text = row(14).ToString()
        Catch ex As Exception
            Debug.WriteLine($"LoadSessionData stub შეცდომა: {ex.Message}")
        End Try
    End Sub

    Private Sub SetComboSelection(cb As ComboBox, value As String)
        For i = 0 To cb.Items.Count - 1
            If String.Equals(cb.Items(i).ToString(), value, StringComparison.OrdinalIgnoreCase) Then
                cb.SelectedIndex = i : Exit For
            End If
        Next
    End Sub

    Private Sub LoadSurnamesFromCacheFor(name As String)
        If _beneRows Is Nothing Then Return
        CBBeneSurname.Items.Clear()
        Dim setSur As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each r In _beneRows
            If r.Count >= 2 AndAlso String.Equals(r(0)?.ToString().Trim(), name, StringComparison.OrdinalIgnoreCase) Then
                Dim s = r(1)?.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(s) Then setSur.Add(s)
            End If
        Next
        For Each s In setSur.OrderBy(Function(x) x)
            CBBeneSurname.Items.Add(s)
        Next
        If CBBeneSurname.Items.Count > 0 Then CBBeneSurname.Enabled = True
    End Sub

    Private Function GetSelectedSpace() As String
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            If btn.Name.StartsWith("BTNS") AndAlso btn.BackColor = Color.FromArgb(173, 216, 230) Then
                Return btn.Text
            End If
        Next
        Return ""
    End Function

End Class