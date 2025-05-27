' ===========================================
' 📄 UserControls/UC_BeneficiaryReport.vb - გაუმჯობესებული ვერსია
' -------------------------------------------
' ბენეფიციარის ანგარიშის UserControl - სრული ფილტრაციის სისტემით
' UC_Schedule.vb-ის მსგავსი ფუნქციონალით
' 
' 🎯 ხელმისაწვდომია როლებისთვის: 1 (ადმინი), 2 (მენეჯერი), 3
' ===========================================
Imports System.ComponentModel
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports System.Text
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_BeneficiaryReport

    ' სერვისები და მონაცემები
    Private dataService As IDataService
    Private allSessions As List(Of SessionModel)
    Private filteredSessions As List(Of SessionModel)
    Private currentBeneficiarySessions As List(Of SessionModel)
    Private printService As AdvancedDataGridViewPrintService

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRole As String = ""

    ' 🔧 ფილტრაციის კონტროლი - ციკლური ივენთების თავიდან აცილება
    Private isUpdatingFilters As Boolean = False
    Private isLoadingData As Boolean = False

    ' სტატუსების მონაცემები
    Private statusCheckBoxes As New List(Of CheckBox)()

    ''' <summary>
    ''' ბენეფიციარის ინფორმაციის კლასი
    ''' </summary>
    Public Class BeneficiaryInfo
        Public Property FullName As String
        Public Property FirstName As String
        Public Property LastName As String
        Public Property TotalSessions As Integer
        Public Property CompletedSessions As Integer
        Public Property CancelledSessions As Integer
        Public Property PendingSessions As Integer
        Public Property TotalAmount As Decimal
        Public Property LastSessionDate As DateTime?
        Public Property FirstSessionDate As DateTime?

        ' 🔧 დამატებითი სტატისტიკა ფილტრაციისთვის
        Public Property MissedUnexcusedSessions As Integer
        Public Property RestoredSessions As Integer
        Public Property OtherStatusSessions As Integer
    End Class

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_BeneficiaryReport: მონაცემთა სერვისი მითითებულია")
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის მითითება
    ''' </summary>
    Public Sub SetUserInfo(email As String, role As String)
        userEmail = email
        userRole = role
        Debug.WriteLine($"UC_BeneficiaryReport: მომხმარებლის ინფორმაცია - Email: {email}, Role: {role}")
    End Sub

    ''' <summary>
    ''' UserControl-ის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_BeneficiaryReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Debug.WriteLine("UC_BeneficiaryReport: UserControl-ის ჩატვირთვა")
            InitializeControls()
            InitializeStatusCheckBoxes()
            LoadInitialData()
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport_Load: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონტროლების საწყისი ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeControls()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: კონტროლების ინიციალიზაცია")

            ' თარიღების საწყისი მნიშვნელობები - მიმდინარე თვე
            Dim currentDate As DateTime = DateTime.Today
            Dim firstDayOfMonth As DateTime = New DateTime(currentDate.Year, currentDate.Month, 1)
            Dim lastDayOfMonth As DateTime = firstDayOfMonth.AddMonths(1).AddDays(-1)

            DtpDan.Value = firstDayOfMonth
            DtpMde.Value = lastDayOfMonth

            ' ComboBox-ების საწყისი მდგომარეობა
            InitializeComboBoxes()

            ' სესიების DataGridView-ის საწყისი მდგომარეობა
            If DgvSessions IsNot Nothing Then
                DgvSessions.DataSource = Nothing
            End If

            ' ღილაკების საწყისი მდგომარეობა
            EnableActionButtons(False)

            Debug.WriteLine("UC_BeneficiaryReport: კონტროლები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"InitializeControls: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ComboBox-ების საწყისი ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeComboBoxes()
        Try
            ' ბენეფიციარის სახელი
            CBBeneName.Items.Clear()
            CBBeneName.Items.Add("ყველა")
            CBBeneName.SelectedIndex = 0
            CBBeneName.Enabled = False

            ' ბენეფიციარის გვარი
            CBBeneSurname.Items.Clear()
            CBBeneSurname.Items.Add("ყველა")
            CBBeneSurname.SelectedIndex = 0
            CBBeneSurname.Enabled = False

            ' 🔧 თერაპევტი
            If CBPer IsNot Nothing Then
                CBPer.Items.Clear()
                CBPer.Items.Add("ყველა")
                CBPer.SelectedIndex = 0
                CBPer.Enabled = False
            End If

            ' 🔧 თერაპიის ტიპი
            If CBTer IsNot Nothing Then
                CBTer.Items.Clear()
                CBTer.Items.Add("ყველა")
                CBTer.SelectedIndex = 0
                CBTer.Enabled = False
            End If

            Debug.WriteLine("UC_BeneficiaryReport: ComboBox-ები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"InitializeComboBoxes: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 სტატუსის CheckBox-ების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeStatusCheckBoxes()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: სტატუსის CheckBox-ების ინიციალიზაცია")

            statusCheckBoxes.Clear()

            ' CheckBox1-დან CheckBox7-მდე ძებნა
            For i As Integer = 1 To 7
                Dim checkBox As CheckBox = FindCheckBoxRecursive(Me, $"CheckBox{i}")
                If checkBox IsNot Nothing Then
                    statusCheckBoxes.Add(checkBox)
                    Debug.WriteLine($"UC_BeneficiaryReport: ნაპოვნია CheckBox{i} - '{checkBox.Text}'")
                End If
            Next

            ' ნაგულისხმევი სტატუსების მონიშვნა
            SetDefaultStatusSelection()

            ' ივენთების მიბმა
            For Each checkBox In statusCheckBoxes
                AddHandler checkBox.CheckedChanged, AddressOf StatusCheckBox_CheckedChanged
            Next

            Debug.WriteLine($"UC_BeneficiaryReport: ინიციალიზებულია {statusCheckBoxes.Count} სტატუსის CheckBox")

        Catch ex As Exception
            Debug.WriteLine($"InitializeStatusCheckBoxes: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ნაგულისხმევი სტატუსების მონიშვნა
    ''' </summary>
    Private Sub SetDefaultStatusSelection()
        Try
            ' ნაგულისხმევად მონიშნული სტატუსები
            Dim defaultStatuses As String() = {"შესრულებული", "გაცდენა არასაპატიო", "აღდგენა"}

            For Each checkBox In statusCheckBoxes
                If checkBox IsNot Nothing Then
                    Dim statusText = checkBox.Text.Trim()

                    ' შევამოწმოთ არის თუ არა ნაგულისხმევ სტატუსებში
                    Dim shouldBeChecked = defaultStatuses.Any(Function(status)
                                                                  String.Equals(statusText, status, StringComparison.OrdinalIgnoreCase))
                    
                    checkBox.Checked = shouldBeChecked

                                                                  Debug.WriteLine($"UC_BeneficiaryReport: '{statusText}' - მონიშნული: {shouldBeChecked}")
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"SetDefaultStatusSelection: შეცდომა - {ex.Message}")
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
    ''' საწყისი მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadInitialData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: საწყისი მონაცემების ჩატვირთვა")

            If dataService Is Nothing Then
                Debug.WriteLine("LoadInitialData: dataService არ არის ხელმისაწვდომი")
                Return
            End If

            ' ყველა სესიის წამოღება
            allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"LoadInitialData: წამოღებულია {allSessions.Count} სესია")

            ' საწყისი ფილტრაცია
            FilterSessionsByDateRange()
            LoadBeneficiaryNames()

        Catch ex As Exception
            Debug.WriteLine($"LoadInitialData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DtpDan თარიღის შეცვლისას
    ''' </summary>
    Private Sub DtpDan_ValueChanged(sender As Object, e As EventArgs) Handles DtpDan.ValueChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: DtpDan შეიცვალა - {DtpDan.Value:dd.MM.yyyy}")

            ' ვალიდაცია - საწყისი თარიღი არ უნდა იყოს უფრო გვიანინ ვიდრე დასასრული
            If DtpDan.Value > DtpMde.Value Then
                DtpMde.Value = DtpDan.Value
            End If

            FilterSessionsByDateRange()
            LoadBeneficiaryNames()
            ResetDependentFilters()

        Catch ex As Exception
            Debug.WriteLine($"DtpDan_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DtpMde თარიღის შეცვლისას
    ''' </summary>
    Private Sub DtpMde_ValueChanged(sender As Object, e As EventArgs) Handles DtpMde.ValueChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: DtpMde შეიცვალა - {DtpMde.Value:dd.MM.yyyy}")

            ' ვალიდაცია - დასასრული თარიღი არ უნდა იყოს უფრო ადრე ვიდრე საწყისი
            If DtpMde.Value < DtpDan.Value Then
                DtpDan.Value = DtpMde.Value
            End If

            FilterSessionsByDateRange()
            LoadBeneficiaryNames()
            ResetDependentFilters()

        Catch ex As Exception
            Debug.WriteLine($"DtpMde_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ფილტრაცია თარიღების მიხედვით
    ''' </summary>
    Private Sub FilterSessionsByDateRange()
        Try
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                filteredSessions = New List(Of SessionModel)()
                Return
            End If

            Dim startDate As DateTime = DtpDan.Value.Date
            Dim endDate As DateTime = DtpMde.Value.Date.AddDays(1).AddTicks(-1)

            filteredSessions = allSessions.Where(Function(s) s.DateTime >= startDate AndAlso s.DateTime <= endDate).ToList()

            Debug.WriteLine($"FilterSessionsByDateRange: ფილტრირებულია {filteredSessions.Count} სესია პერიოდისთვის {startDate:dd.MM.yyyy} - {DtpMde.Value.Date:dd.MM.yyyy}")

        Catch ex As Exception
            Debug.WriteLine($"FilterSessionsByDateRange: შეცდომა - {ex.Message}")
            filteredSessions = New List(Of SessionModel)()
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სახელების ჩატვირთვა CBBeneName-ში
    ''' </summary>
    Private Sub LoadBeneficiaryNames()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარების სახელების ჩატვირთვა")

            isUpdatingFilters = True

            Try
                CBBeneName.Items.Clear()
                CBBeneName.Items.Add("ყველა")
                CBBeneName.Text = "ყველა"

                If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                    CBBeneName.Enabled = False
                    Debug.WriteLine("LoadBeneficiaryNames: ფილტრირებული სესიები ცარიელია")
                    Return
                End If

                ' უნიკალური სახელების მიღება
                Dim uniqueNames = filteredSessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                               .Select(Function(s) s.BeneficiaryName.Trim()) _
                                               .Distinct() _
                                               .OrderBy(Function(name) name) _
                                               .ToList()

                ' სახელების დამატება
                For Each Name In uniqueNames
                    CBBeneName.Items.Add(Name)
                Next

                CBBeneName.Enabled = True
                CBBeneName.SelectedIndex = 0

                Debug.WriteLine($"LoadBeneficiaryNames: ჩატვირთულია {uniqueNames.Count} უნიკალური სახელი")

            Finally
                isUpdatingFilters = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiaryNames: შეცდომა - {ex.Message}")
            CBBeneName.Enabled = False
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneName-ის მნიშვნელობის შეცვლისას
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneName შეიცვალა - '{CBBeneName.SelectedItem?.ToString()}'")

            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) OrElse selectedName = "ყველა" Then
                ResetDependentFilters()
                ClearBeneficiaryData()
                Return
            End If

            LoadBeneficiarySurnames(selectedName)

        Catch ex As Exception
            Debug.WriteLine($"CBBeneName_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' არჩეული სახელის შესაბამისი გვარების ჩატვირთვა
    ''' </summary>
    Private Sub LoadBeneficiarySurnames(selectedName As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: გვარების ჩატვირთვა სახელისთვის '{selectedName}'")

            isUpdatingFilters = True

            Try
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Items.Add("ყველა")
                CBBeneSurname.Text = "ყველა"

                If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                    CBBeneSurname.Enabled = False
                    Return
                End If

                ' არჩეული სახელის შესაბამისი უნიკალური გვარები
                Dim uniqueSurnames = filteredSessions.Where(Function(s)
                                                                s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                    Not String.IsNullOrEmpty(s.BeneficiarySurname)) _
                    .Select(Function(s) s.BeneficiarySurname.Trim()) _
                    .Distinct() _
                    .OrderBy(Function(surname) surname) _
                    .ToList()

                ' გვარების დამატება
                For Each surname In uniqueSurnames
                                                                    CBBeneSurname.Items.Add(surname)
                                                                Next

                                                                CBBeneSurname.Enabled = True
                                                                CBBeneSurname.SelectedIndex = 0

                                                                ' თუ მხოლოდ ერთი გვარია, ავტომატურად აირჩიოს
                                                                If uniqueSurnames.Count = 1 Then
                                                                    CBBeneSurname.SelectedIndex = 1
                                                                End If

                                                                Debug.WriteLine($"LoadBeneficiarySurnames: ჩატვირთულია {uniqueSurnames.Count} უნიკალური გვარი")

            Finally
                isUpdatingFilters = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiarySurnames: შეცდომა - {ex.Message}")
            CBBeneSurname.Enabled = False
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneSurname-ის მნიშვნელობის შეცვლისას
    ''' </summary>
    Private Sub CBBeneSurname_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneSurname.SelectedIndexChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: CBBeneSurname შეიცვალა - '{CBBeneSurname.SelectedItem?.ToString()}'")

            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) OrElse selectedName = "ყველა" OrElse
               String.IsNullOrEmpty(selectedSurname) OrElse selectedSurname = "ყველა" Then

                ' თუ ერთ-ერთი ცარიელია, მხოლოდ თერაპევტებისა და თერაპიების ფილტრი დავანულოთ
                ResetTherapistAndTherapyFilters()
                ClearBeneficiaryData()
                Return
            End If

            ' კონკრეტული ბენეფიციარის არჩევისას
            LoadTherapistAndTherapyFilters(selectedName, selectedSurname)
            LoadBeneficiaryData(selectedName, selectedSurname)

        Catch ex As Exception
            Debug.WriteLine($"CBBeneSurname_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპევტების და თერაპიების ფილტრების ჩატვირთვა
    ''' </summary>
    Private Sub LoadTherapistAndTherapyFilters(selectedName As String, selectedSurname As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: თერაპევტების და თერაპიების ფილტრების ჩატვირთვა - '{selectedName} {selectedSurname}'")

            isUpdatingFilters = True

            Try
                ' ბენეფიციარის სესიების ფილტრაცია
                Dim beneficiarySessions = filteredSessions.Where(Function(s)
                                                                     s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                    s.BeneficiarySurname.Trim().Equals(selectedSurname, StringComparison.OrdinalIgnoreCase)
                                                                 End Function).ToList()

                ' 🔧 თერაპევტების ComboBox-ის შევსება
                If CBPer IsNot Nothing Then
                    CBPer.Items.Clear()
                    CBPer.Items.Add("ყველა")

                    Dim uniqueTherapists = beneficiarySessions.Where(Function(s) Not String.IsNullOrEmpty(s.TherapistName)) _
                                                             .Select(Function(s) s.TherapistName.Trim()) _
                                                             .Distinct() _
                                                             .OrderBy(Function(therapist) therapist) _
                                                             .ToList()

                    For Each therapist In uniqueTherapists
                        CBPer.Items.Add(therapist)
                    Next

                    CBPer.SelectedIndex = 0
                    CBPer.Enabled = True

                    Debug.WriteLine($"LoadTherapistAndTherapyFilters: ჩატვირთულია {uniqueTherapists.Count} თერაპევტი")
                End If

                ' 🔧 თერაპიების ComboBox-ის შევსება
                If CBTer IsNot Nothing Then
                    CBTer.Items.Clear()
                    CBTer.Items.Add("ყველა")

                    Dim uniqueTherapies = beneficiarySessions.Where(Function(s) Not String.IsNullOrEmpty(s.TherapyType)) _
                                                            .Select(Function(s) s.TherapyType.Trim()) _
                                                            .Distinct() _
                                                            .OrderBy(Function(therapy) therapy) _
                                                            .ToList()

                    For Each therapy In uniqueTherapies
                        CBTer.Items.Add(therapy)
                    Next

                    CBTer.SelectedIndex = 0
                    CBTer.Enabled = True

                    Debug.WriteLine($"LoadTherapistAndTherapyFilters: ჩატვირთულია {uniqueTherapies.Count} თერაპია")
                End If

            Finally
                isUpdatingFilters = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"LoadTherapistAndTherapyFilters: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 CBPer (თერაპევტი) ComboBox-ის შეცვლისას
    ''' </summary>
    Private Sub CBPer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPer.SelectedIndexChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: CBPer შეიცვალა - '{CBPer.SelectedItem?.ToString()}'")

            ' ბენეფიციარის მონაცემების ხელახალი ჩატვირთვა ახალი ფილტრით
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"CBPer_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 CBTer (თერაპია) ComboBox-ის შეცვლისას
    ''' </summary>
    Private Sub CBTer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTer.SelectedIndexChanged
        Try
            If isUpdatingFilters Then Return

            Debug.WriteLine($"UC_BeneficiaryReport: CBTer შეიცვალა - '{CBTer.SelectedItem?.ToString()}'")

            ' ბენეფიციარის მონაცემების ხელახალი ჩატვირთვა ახალი ფილტრით
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"CBTer_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 სტატუსის CheckBox-ის შეცვლისას
    ''' </summary>
    Private Sub StatusCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        Try
            If isUpdatingFilters Then Return

            Dim checkBox As CheckBox = DirectCast(sender, CheckBox)
            Debug.WriteLine($"UC_BeneficiaryReport: სტატუსის CheckBox შეიცვალა - '{checkBox.Text}' = {checkBox.Checked}")

            ' ყველა ფილტრის გადატარება
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"StatusCheckBox_CheckedChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ყველა ფილტრის გადატარება
    ''' </summary>
    Private Sub ApplyAllFilters()
        Try
            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) OrElse selectedName = "ყველა" OrElse
               String.IsNullOrEmpty(selectedSurname) OrElse selectedSurname = "ყველა" Then
                ClearBeneficiaryData()
                Return
            End If

            ' ყველა ფილტრის გათვალისწინებით მონაცემების ჩატვირთვა
            LoadBeneficiaryDataWithFilters(selectedName, selectedSurname)

        Catch ex As Exception
            Debug.WriteLine($"ApplyAllFilters: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ბენეფიციარის მონაცემების ჩატვირთვა ყველა ფილტრით
    ''' </summary>
    Private Sub LoadBeneficiaryDataWithFilters(name As String, surname As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: ბენეფიციარის მონაცემების ჩატვირთვა ფილტრებით - '{name} {surname}'")

            ' ძირითადი ბენეფიციარის სესიების ფილტრაცია
            Dim beneficiarySessions = filteredSessions.Where(Function(s)
                                                                 s.BeneficiaryName.Trim().Equals(name, StringComparison.OrdinalIgnoreCase) AndAlso
                s.BeneficiarySurname.Trim().Equals(surname, StringComparison.OrdinalIgnoreCase)
                                                             End Function).ToList()

            ' 🔧 თერაპევტის ფილტრი
            Dim selectedTherapist As String = CBPer?.SelectedItem?.ToString()
            If Not String.IsNullOrEmpty(selectedTherapist) AndAlso selectedTherapist <> "ყველა" Then
                beneficiarySessions = beneficiarySessions.Where(Function(s)
                                                                    s.TherapistName.Trim().Equals(selectedTherapist, StringComparison.OrdinalIgnoreCase)
                                                                End Function).ToList()
                Debug.WriteLine($"LoadBeneficiaryDataWithFilters: თერაპევტის ფილტრი '{selectedTherapist}' - დარჩა {beneficiarySessions.Count} სესია")
            End If

            ' 🔧 თერაპიის ტიპის ფილტრი
            Dim selectedTherapy As String = CBTer?.SelectedItem?.ToString()
            If Not String.IsNullOrEmpty(selectedTherapy) AndAlso selectedTherapy <> "ყველა" Then
                beneficiarySessions = beneficiarySessions.Where(Function(s)
                                                                    s.TherapyType.Trim().Equals(selectedTherapy, StringComparison.OrdinalIgnoreCase)
                                                                End Function).ToList()
                Debug.WriteLine($"LoadBeneficiaryDataWithFilters: თერაპიის ფილტრი '{selectedTherapy}' - დარჩა {beneficiarySessions.Count} სესია")
            End If

            ' 🔧 სტატუსის ფილტრი
            Dim selectedStatuses = GetSelectedStatuses()
            If selectedStatuses.Count > 0 Then
                beneficiarySessions = beneficiarySessions.Where(Function(s)
                                                                    selectedStatuses.Any(Function(status)
                                                                                             String.Equals(s.Status.Trim(), status.Trim(), StringComparison.OrdinalIgnoreCase))
                End Function).ToList()
                                                                    Debug.WriteLine($"LoadBeneficiaryDataWithFilters: სტატუსის ფილტრი ({selectedStatuses.Count} სტატუსი) - დარჩა {beneficiarySessions.Count} სესია")
            End If

            If beneficiarySessions.Count = 0 Then
                Debug.WriteLine($"LoadBeneficiaryDataWithFilters: ფილტრების შემდეგ სესიები ვერ მოიძებნა")
                ClearBeneficiaryData()
                Return
            End If

            ' დალაგება თარიღის მიხედვით (ახლიდან ძველისკენ)
            beneficiarySessions = beneficiarySessions.OrderByDescending(Function(s) s.DateTime).ToList()

            ' მიმდინარე ბენეფიციარის სესიების განახლება
            currentBeneficiarySessions = beneficiarySessions

            ' ბენეფიციარის ინფორმაციის შექმნა
            Dim beneficiaryInfo = CreateBeneficiaryInfo(name, surname, currentBeneficiarySessions)

            ' სტატისტიკის განახლება
            UpdateBeneficiaryStatistics(beneficiaryInfo)

            ' სესიების ჩვენება DataGridView-ში
            LoadSessionsToGrid(currentBeneficiarySessions)

            ' ღილაკების ააქტიურება
            EnableActionButtons(True)

            Debug.WriteLine($"LoadBeneficiaryDataWithFilters: საბოლოოდ ჩატვირთულია {currentBeneficiarySessions.Count} სესია")

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiaryDataWithFilters: შეცდომა - {ex.Message}")
            ClearBeneficiaryData()
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 მონიშნული სტატუსების მიღება
    ''' </summary>
    Private Function GetSelectedStatuses() As List(Of String)
        Dim selectedStatuses As New List(Of String)()

        Try
            For Each checkBox In statusCheckBoxes
                If checkBox IsNot Nothing AndAlso checkBox.Checked Then
                    Dim statusText = checkBox.Text.Trim()
                    If Not String.IsNullOrEmpty(statusText) Then
                        selectedStatuses.Add(statusText)
                    End If
                End If
            Next

            Debug.WriteLine($"GetSelectedStatuses: მონიშნულია {selectedStatuses.Count} სტატუსი: {String.Join(", ", selectedStatuses)}")

        Catch ex As Exception
            Debug.WriteLine($"GetSelectedStatuses: შეცდომა - {ex.Message}")
        End Try

        Return selectedStatuses
    End Function

    ''' <summary>
    ''' დამოკიდებული ფილტრების რესეტი
    ''' </summary>
    Private Sub ResetDependentFilters()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: დამოკიდებული ფილტრების რესეტი")

            isUpdatingFilters = True

            Try
                ' ბენეფიციარის გვარის რესეტი
                CBBeneSurname.Items.Clear()
                CBBeneSurname.Items.Add("ყველა")
                CBBeneSurname.SelectedIndex = 0
                CBBeneSurname.Enabled = False

                ' თერაპევტებისა და თერაპიების რესეტი
                ResetTherapistAndTherapyFilters()

            Finally
                isUpdatingFilters = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ResetDependentFilters: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 თერაპევტებისა და თერაპიების ფილტრების რესეტი
    ''' </summary>
    Private Sub ResetTherapistAndTherapyFilters()
        Try
            isUpdatingFilters = True

            Try
                ' თერაპევტის ComboBox-ის რესეტი
                If CBPer IsNot Nothing Then
                    CBPer.Items.Clear()
                    CBPer.Items.Add("ყველა")
                    CBPer.SelectedIndex = 0
                    CBPer.Enabled = False
                End If

                ' თერაპიის ComboBox-ის რესეტი
                If CBTer IsNot Nothing Then
                    CBTer.Items.Clear()
                    CBTer.Items.Add("ყველა")
                    CBTer.SelectedIndex = 0
                    CBTer.Enabled = False
                End If

            Finally
                isUpdatingFilters = False
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ResetTherapistAndTherapyFilters: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის მონაცემების ჩატვირთვა (ძველი მეთოდი - თავსებადობისთვის)
    ''' </summary>
    Private Sub LoadBeneficiaryData(name As String, surname As String)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: LoadBeneficiaryData (ძველი მეთოდი) - '{name} {surname}'")
            LoadBeneficiaryDataWithFilters(name, surname)
        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiaryData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის ინფორმაციის შექმნა - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Function CreateBeneficiaryInfo(name As String, surname As String, sessions As List(Of SessionModel)) As BeneficiaryInfo
        Try
            Dim info As New BeneficiaryInfo()
            info.FirstName = name
            info.LastName = surname
            info.FullName = $"{name} {surname}"

            ' ძირითადი სტატისტიკა
            info.TotalSessions = sessions.Count
            info.CompletedSessions = sessions.Where(Function(s) s.Status = "შესრულებული").Count()
            info.CancelledSessions = sessions.Where(Function(s) s.Status = "გაუქმებული").Count()
            info.PendingSessions = sessions.Where(Function(s) s.Status = "დაგეგმილი").Count()

            ' 🔧 დამატებითი სტატისტიკა
            info.MissedUnexcusedSessions = sessions.Where(Function(s) s.Status = "გაცდენა არასაპატიო").Count()
            info.RestoredSessions = sessions.Where(Function(s) s.Status = "აღდგენა").Count()
            info.OtherStatusSessions = sessions.Where(Function(s)
                                                          s.Status <> "შესრულებული" AndAlso 
                s.Status <> "გაუქმებული" AndAlso 
                s.Status <> "დაგეგმილი" AndAlso
                s.Status <> "გაცდენა არასაპატიო" AndAlso
                s.Status <> "აღდგენა"
            ).Count()

            ' ჯამური თანხა
            info.TotalAmount = sessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)

            ' თარიღების დადგენა
            If sessions.Count > 0 Then
                info.FirstSessionDate = sessions.Min(Function(s) s.DateTime)
                info.LastSessionDate = sessions.Max(Function(s) s.DateTime)
            End If

            Debug.WriteLine($"CreateBeneficiaryInfo: შეიქმნა ინფორმაცია - სულ: {info.TotalSessions}, შესრულებული: {info.CompletedSessions}")

            Return info

        Catch ex As Exception
            Debug.WriteLine($"CreateBeneficiaryInfo: შეცდომა - {ex.Message}")
            Return New BeneficiaryInfo()
        End Try
    End Function

    ''' <summary>
    ''' ბენეფიციარის სტატისტიკის განახლება UI-ზე - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub UpdateBeneficiaryStatistics(info As BeneficiaryInfo)
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: სტატისტიკის განახლება - {info.FullName}")

            ' ძირითადი ინფორმაცია
            If LblBeneficiaryName IsNot Nothing Then
                LblBeneficiaryName.Text = $"ბენეფიციარი: {info.FullName}"
            End If

            If LblTotalSessions IsNot Nothing Then
                LblTotalSessions.Text = $"სულ სესიები: {info.TotalSessions}"
            End If

            If LblCompletedSessions IsNot Nothing Then
                LblCompletedSessions.Text = $"შესრულებული: {info.CompletedSessions}"
            End If

            If LblCancelledSessions IsNot Nothing Then
                LblCancelledSessions.Text = $"გაუქმებული: {info.CancelledSessions}"
            End If

            If LblPendingSessions IsNot Nothing Then
                LblPendingSessions.Text = $"დაგეგმილი: {info.PendingSessions}"
            End If

            ' 🔧 დამატებითი სტატისტიკა (თუ Label-ები არსებობს)
            If Me.Controls.OfType(Of Label).Any(Function(l) l.Name = "LblMissedUnexcused") Then
                Dim lblMissed = Me.Controls.OfType(Of Label).First(Function(l) l.Name = "LblMissedUnexcused")
                lblMissed.Text = $"გაცდენა არასაპატიო: {info.MissedUnexcusedSessions}"
            End If

            If Me.Controls.OfType(Of Label).Any(Function(l) l.Name = "LblRestored") Then
                Dim lblRestored = Me.Controls.OfType(Of Label).First(Function(l) l.Name = "LblRestored")
                lblRestored.Text = $"აღდგენა: {info.RestoredSessions}"
            End If

            ' ფინანსური ინფორმაცია
            If LblTotalAmount IsNot Nothing Then
                LblTotalAmount.Text = $"ჯამური თანხა: {info.TotalAmount:N2} ₾"
            End If

            ' თარიღების ინფორმაცია
            If LblFirstSession IsNot Nothing Then
                LblFirstSession.Text = If(info.FirstSessionDate.HasValue,
                                       $"პირველი სესია: {info.FirstSessionDate.Value:dd.MM.yyyy}",
                                       "პირველი სესია: -")
            End If

            If LblLastSession IsNot Nothing Then
                LblLastSession.Text = If(info.LastSessionDate.HasValue,
                                      $"ბოლო სესია: {info.LastSessionDate.Value:dd.MM.yyyy}",
                                      "ბოლო სესია: -")
            End If

            Debug.WriteLine($"UpdateBeneficiaryStatistics: სტატისტიკა წარმატებით განახლდა")

        Catch ex As Exception
            Debug.WriteLine($"UpdateBeneficiaryStatistics: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ჩატვირთვა DataGridView-ში
    ''' </summary>
    Private Sub LoadSessionsToGrid(sessions As List(Of SessionModel))
        Try
            Debug.WriteLine($"UC_BeneficiaryReport: სესიების ჩატვირთვა DataGridView-ში - {sessions.Count} სესია")

            If DgvSessions Is Nothing Then
                Debug.WriteLine("LoadSessionsToGrid: DgvSessions არ არსებობს")
                Return
            End If

            ' DataGridView-ის განახლება
            DgvSessions.DataSource = Nothing
            DgvSessions.DataSource = sessions

            ' მწკრივების ფერების განახლება SessionStatusColors კლასის გამოყენებით
            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim session = TryCast(row.DataBoundItem, SessionModel)
                    If session IsNot Nothing Then
                        ' SessionStatusColors კლასის გამოყენება
                        Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.BackColor = statusColor

                        ' ჩარჩოს ფერიც დავაყენოთ
                        Dim borderColor = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.SelectionBackColor = borderColor
                    End If
                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"LoadSessionsToGrid: ჩატვირთულია {sessions.Count} სესია DataGridView-ში")

        Catch ex As Exception
            Debug.WriteLine($"LoadSessionsToGrid: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ღილაკების ააქტიურება/გათიშვა
    ''' </summary>
    Private Sub EnableActionButtons(enabled As Boolean)
        Try
            If BtnExportToExcel IsNot Nothing Then
                BtnExportToExcel.Enabled = enabled
            End If

            If BtnPrint IsNot Nothing Then
                BtnPrint.Enabled = enabled
            End If

            If BtnGenerateReport IsNot Nothing Then
                BtnGenerateReport.Enabled = enabled
            End If

            Debug.WriteLine($"EnableActionButtons: ღილაკები {'ააქტიურებულია' If enabled Else 'გათიშულია'}")

        Catch ex As Exception
            Debug.WriteLine($"EnableActionButtons: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის მონაცემების გასუფთავება
    ''' </summary>
    Private Sub ClearBeneficiaryData()
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ბენეფიციარის მონაცემების გასუფთავება")

            currentBeneficiarySessions = New List(Of SessionModel)()

            ' სტატისტიკის გასუფთავება
            If LblBeneficiaryName IsNot Nothing Then LblBeneficiaryName.Text = "ბენეფიციარი: -"
            If LblTotalSessions IsNot Nothing Then LblTotalSessions.Text = "სულ სესიები: 0"
            If LblCompletedSessions IsNot Nothing Then LblCompletedSessions.Text = "შესრულებული: 0"
            If LblCancelledSessions IsNot Nothing Then LblCancelledSessions.Text = "გაუქმებული: 0"
            If LblPendingSessions IsNot Nothing Then LblPendingSessions.Text = "დაგეგმილი: 0"
            If LblTotalAmount IsNot Nothing Then LblTotalAmount.Text = "ჯამური თანხა: 0.00 ₾"
            If LblFirstSession IsNot Nothing Then LblFirstSession.Text = "პირველი სესია: -"
            If LblLastSession IsNot Nothing Then LblLastSession.Text = "ბოლო სესია: -"

            ' DataGridView-ის გასუფთავება
            If DgvSessions IsNot Nothing Then
                DgvSessions.DataSource = Nothing
            End If

            ' ღილაკების გათიშვა
            EnableActionButtons(False)

            Debug.WriteLine("ClearBeneficiaryData: მონაცემები გაწმინდა")

        Catch ex As Exception
            Debug.WriteLine($"ClearBeneficiaryData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 მონაცემების განახლების ღილაკი (თუ არსებობს)
    ''' </summary>
    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: მონაცემების განახლება")

            ' ქეშის გაუქმება
            If TypeOf dataService Is SheetDataService Then
                DirectCast(dataService, SheetDataService).InvalidateAllCache()
            End If

            ' მონაცემების ხელახალი ჩატვირთვა
            LoadInitialData()

            MessageBox.Show("მონაცემები წარმატებით განახლდა", "ინფორმაცია",
                           MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"BtnRefresh_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"მონაცემების განახლების შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ===========================================
    ' 📄 არსებული ღილაკების მეთოდები (შეუცვლელი)
    ' ===========================================

    ''' <summary>
    ''' Excel-ში ექსპორტის ღილაკზე დაჭერა
    ''' </summary>
    Private Sub BtnExportToExcel_Click(sender As Object, e As EventArgs) Handles BtnExportToExcel.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("ექსპორტისთვის ჯერ აირჩიეთ ბენეფიციარი და გამოიყენეთ ფილტრები", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ExportToCSV()

        Catch ex As Exception
            Debug.WriteLine($"BtnExportToExcel_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ექსპორტის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' CSV ფორმატში ექსპორტი - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub ExportToCSV()
        Try
            Dim selectedName = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname = CBBeneSurname.SelectedItem?.ToString()
            Dim selectedTherapist = CBPer?.SelectedItem?.ToString()
            Dim selectedTherapy = CBTer?.SelectedItem?.ToString()

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "CSV Files|*.csv"
                saveDialog.Title = "ბენეფიციარის ანგარიშის ექსპორტი"
                saveDialog.FileName = $"ბენეფიციარის_ანგარიში_{selectedName}_{selectedSurname}_{DateTime.Now:yyyyMMdd}.csv"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    Using writer As New System.IO.StreamWriter(saveDialog.FileName, False, System.Text.Encoding.UTF8)
                        ' BOM-ის დამატება Excel-ისთვის
                        writer.Write(System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble()))

                        ' სათაური ინფორმაცია
                        writer.WriteLine($"ბენეფიციარის ანგარიში,{selectedName} {selectedSurname}")
                        writer.WriteLine($"პერიოდი,{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")
                        writer.WriteLine($"გენერირების თარიღი,{DateTime.Now:dd.MM.yyyy HH:mm}")
                        writer.WriteLine($"მომხმარებელი,{userEmail}")
                        
                        ' 🔧 ფილტრების ინფორმაცია
                        If Not String.IsNullOrEmpty(selectedTherapist) AndAlso selectedTherapist <> "ყველა" Then
                            writer.WriteLine($"თერაპევტი,{selectedTherapist}")
                        End If
                        If Not String.IsNullOrEmpty(selectedTherapy) AndAlso selectedTherapy <> "ყველა" Then
                            writer.WriteLine($"თერაპია,{selectedTherapy}")
                        End If
                        
                        Dim selectedStatuses = GetSelectedStatuses()
                        If selectedStatuses.Count > 0 Then
                            writer.WriteLine($"სტატუსები,{String.Join("; ", selectedStatuses)}")
                        End If
                        
                        writer.WriteLine()

                        ' სტატისტიკა
                        writer.WriteLine("სტატისტიკა")
                        writer.WriteLine($"ფილტრირებული სესიები,{currentBeneficiarySessions.Count}")
                        writer.WriteLine($"შესრულებული,{currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Count()}")
                        writer.WriteLine($"გაუქმებული,{currentBeneficiarySessions.Where(Function(s) s.Status = "გაუქმებული").Count()}")
                        writer.WriteLine($"დაგეგმილი,{currentBeneficiarySessions.Where(Function(s) s.Status = "დაგეგმილი").Count()}")
                        writer.WriteLine($"ჯამური თანხა,{currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price):N2} ₾")
                        writer.WriteLine()

                        ' სესიების სათაურები
                        writer.WriteLine("ID,თარიღი,თერაპევტი,თერაპიის ტიპი,ხანგრძლივობა (წთ),ფასი (₾),სტატუსი,დაფინანსება,კომენტარები")

                        ' სესიების მონაცემები
                        For Each session In currentBeneficiarySessions
                            Dim line As String = $"{session.Id}," &
                                               $"{session.DateTime:dd.MM.yyyy HH:mm}," &
                                               $"""{session.TherapistName}""," &
                                               $"""{session.TherapyType}""," &
                                               $"{session.Duration}," &
                                               $"{session.Price:N2}," &
                                               $"""{session.Status}""," &
                                               $"""{session.Funding}""," &
                                               $"""{session.Comments.Replace("""", """""").Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")}"""
                            writer.WriteLine(line)
                        Next
                    End Using

                    MessageBox.Show($"ბენეფიციარის ანგარიში წარმატებით ექსპორტირდა:{Environment.NewLine}{saveDialog.FileName}",
                                  "ექსპორტი დასრულდა", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"ExportToCSV: შეცდომა - {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ბეჭდვის ღილაკზე დაჭერა
    ''' </summary>
    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის ჯერ აირჩიეთ ბენეფიციარი და გამოიყენეთ ფილტრები", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If DgvSessions Is Nothing Then
                MessageBox.Show("სესიების ცხრილი არ არის ხელმისაწვდომი", "შეცდომა",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ბეჭდვის სერვისის ინიციალიზაცია
            If printService IsNot Nothing Then
                printService.Dispose()
            End If

            printService = New AdvancedDataGridViewPrintService(DgvSessions)
            printService.ShowFullPrintDialog()

        Catch ex As Exception
            Debug.WriteLine($"BtnPrint_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დეტალური რეპორტის გენერირების ღილაკზე დაჭერა - გაუმჯობესებული ვერსია
    ''' </summary>
    Private Sub BtnGenerateReport_Click(sender As Object, e As EventArgs) Handles BtnGenerateReport.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("რეპორტის გენერირებისთვის ჯერ აირჩიეთ ბენეფიციარი და გამოიყენეთ ფილტრები", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            GenerateDetailedReport()

        Catch ex As Exception
            Debug.WriteLine($"BtnGenerateReport_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"რეპორტის გენერირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დეტალური რეპორტის გენერირება - გაუმჯობესებული ვერსია ფილტრების ჩათვლით
    ''' </summary>
    Private Sub GenerateDetailedReport()
        Try
            Dim selectedName = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname = CBBeneSurname.SelectedItem?.ToString()
            Dim selectedTherapist = CBPer?.SelectedItem?.ToString()
            Dim selectedTherapy = CBTer?.SelectedItem?.ToString()
            Dim selectedStatuses = GetSelectedStatuses()

            ' დეტალური რეპორტის ტექსტის შექმნა
            Dim report As New StringBuilder()

            ' რეპორტის სათაური
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine($"           ბენეფიციარის დეტალური ანგარიში")
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine()

            ' ბენეფიციარის ინფორმაცია
            report.AppendLine("📋 ბენეფიციარის ინფორმაცია:")
            report.AppendLine($"   სახელი, გვარი: {selectedName} {selectedSurname}")
            report.AppendLine($"   პერიოდი: {DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")
            report.AppendLine($"   გენერირების თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}")
            report.AppendLine($"   მომხმარებელი: {userEmail}")
            report.AppendLine()

            ' 🔧 ფილტრების ინფორმაცია
            report.AppendLine("🔍 გამოყენებული ფილტრები:")
            If Not String.IsNullOrEmpty(selectedTherapist) AndAlso selectedTherapist <> "ყველა" Then
                report.AppendLine($"   თერაპევტი: {selectedTherapist}")
            Else
                report.AppendLine($"   თერაპევტი: ყველა")
            End If
            
            If Not String.IsNullOrEmpty(selectedTherapy) AndAlso selectedTherapy <> "ყველა" Then
                report.AppendLine($"   თერაპიის ტიპი: {selectedTherapy}")
            Else
                report.AppendLine($"   თერაპიის ტიპი: ყველა")
            End If
            
            If selectedStatuses.Count > 0 Then
                report.AppendLine($"   სტატუსები: {String.Join(", ", selectedStatuses)}")
            Else
                report.AppendLine($"   სტატუსები: ყველა")
            End If
            report.AppendLine()

            ' სტატისტიკა
            report.AppendLine("📊 სტატისტიკა:")
            report.AppendLine($"   ფილტრირებული სესიები: {currentBeneficiarySessions.Count}")
            report.AppendLine($"   შესრულებული: {currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Count()}")
            report.AppendLine($"   გაუქმებული: {currentBeneficiarySessions.Where(Function(s) s.Status = "გაუქმებული").Count()}")
            report.AppendLine($"   დაგეგმილი: {currentBeneficiarySessions.Where(Function(s) s.Status = "დაგეგმილი").Count()}")
            report.AppendLine($"   გაცდენა არასაპატიო: {currentBeneficiarySessions.Where(Function(s) s.Status = "გაცდენა არასაპატიო").Count()}")
            report.AppendLine($"   აღდგენა: {currentBeneficiarySessions.Where(Function(s) s.Status = "აღდგენა").Count()}")
            report.AppendLine($"   ჯამური თანხა: {currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price):N2} ₾")

            If currentBeneficiarySessions.Count > 0 Then
                report.AppendLine($"   პირველი სესია: {currentBeneficiarySessions.Min(Function(s) s.DateTime):dd.MM.yyyy}")
                report.AppendLine($"   ბოლო სესია: {currentBeneficiarySessions.Max(Function(s) s.DateTime):dd.MM.yyyy}")
            End If
            report.AppendLine()

            ' 🔧 თერაპევტების და თერაპიების ანალიზი
            If currentBeneficiarySessions.Count > 0 Then
                report.AppendLine("👨‍⚕️ თერაპევტების ანალიზი:")
                Dim therapistStats = currentBeneficiarySessions.GroupBy(Function(s) s.TherapistName) _
                    .Select(Function(g) New With {
                        .Name = g.Key,
                        .Count = g.Count(),
                        .CompletedCount = g.Where(Function(s) s.Status = "შესრულებული").Count(),
                        .TotalAmount = g.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)
                    }).OrderByDescending(Function(t) t.Count)

                For Each therapist In therapistStats
                    report.AppendLine($"   • {therapist.Name}: {therapist.Count} სესია ({therapist.CompletedCount} შესრულებული, {therapist.TotalAmount:N2} ₾)")
                Next
                report.AppendLine()

                report.AppendLine("💊 თერაპიების ანალიზი:")
                Dim therapyStats = currentBeneficiarySessions.GroupBy(Function(s) s.TherapyType) _
                    .Select(Function(g) New With {
                        .Type = g.Key,
                        .Count = g.Count(),
                        .CompletedCount = g.Where(Function(s) s.Status = "შესრულებული").Count(),
                        .TotalAmount = g.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)
                    }).OrderByDescending(Function(t) t.Count)

                For Each therapy In therapyStats
                    report.AppendLine($"   • {therapy.Type}: {therapy.Count} სესია ({therapy.CompletedCount} შესრულებული, {therapy.TotalAmount:N2} ₾)")
                Next
                report.AppendLine()
            End If

            ' სესიების დეტალები
            report.AppendLine("📅 სესიების დეტალები:")
            report.AppendLine("─────────────────────────────────────────────────────────────")

            For Each session In currentBeneficiarySessions
                report.AppendLine($"🔸 სესია #{session.Id}")
                report.AppendLine($"   თარიღი: {session.DateTime:dd.MM.yyyy HH:mm}")
                report.AppendLine($"   თერაპევტი: {session.TherapistName}")
                report.AppendLine($"   თერაპიის ტიპი: {session.TherapyType}")
                report.AppendLine($"   ხანგრძლივობა: {session.Duration} წუთი")
                report.AppendLine($"   ფასი: {session.Price:N2} ₾")
                report.AppendLine($"   სტატუსი: {session.Status}")
                report.AppendLine($"   დაფინანსება: {session.Funding}")
                If Not String.IsNullOrEmpty(session.Comments) Then
                    report.AppendLine($"   კომენტარები: {session.Comments}")
                End If
                report.AppendLine("─────────────────────────────────────────────────────────────")
            Next

            report.AppendLine()
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine($"რეპორტი გენერირებულია: {DateTime.Now:dd.MM.yyyy HH:mm:ss}")
            report.AppendLine("Scheduler v8.8a - ბენეფიციარის ანგარიშის მოდული")
            report.AppendLine("═══════════════════════════════════════════════════════════════")

            ' რეპორტის შენახვა ფაილად
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "Text Files|*.txt"
                saveDialog.Title = "რეპორტის შენახვა"
                saveDialog.FileName = $"ბენეფიციარის_ანგარიში_{selectedName}_{selectedSurname}_{DateTime.Now:yyyyMMdd}.txt"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    System.IO.File.WriteAllText(saveDialog.FileName, report.ToString(), System.Text.Encoding.UTF8)
                    MessageBox.Show($"რეპორტი შენახულია:{Environment.NewLine}{saveDialog.FileName}",
                                  "შენახვა დასრულდა", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' ფაილის გახსნის შეთავაზება
                    Dim openResult As DialogResult = MessageBox.Show(
                        "გსურთ რეპორტის გახსნა?",
                        "ფაილის გახსნა",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question)

                    If openResult = DialogResult.Yes Then
                        System.Diagnostics.Process.Start(saveDialog.FileName)
                    End If
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"GenerateDetailedReport: შეცდომა - {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ყველა სტატუსის CheckBox-ის მონიშვნა/განიშვნა (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnSelectAllStatuses_Click(sender As Object, e As EventArgs) Handles BtnSelectAllStatuses.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სტატუსის მონიშვნა")

            isUpdatingFilters = True

            Try
                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing Then
                        checkBox.Checked = True
                    End If
                Next

                Debug.WriteLine($"BtnSelectAllStatuses_Click: მონიშნულია {statusCheckBoxes.Count} სტატუსი")

            Finally
                isUpdatingFilters = False
            End Try

            ' ფილტრების გადატარება
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"BtnSelectAllStatuses_Click: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ყველა სტატუსის CheckBox-ის განიშვნა (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnDeselectAllStatuses_Click(sender As Object, e As EventArgs) Handles BtnDeselectAllStatuses.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ყველა სტატუსის განიშვნა")

            isUpdatingFilters = True

            Try
                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing Then
                        checkBox.Checked = False
                    End If
                Next

                Debug.WriteLine($"BtnDeselectAllStatuses_Click: განიშნულია {statusCheckBoxes.Count} სტატუსი")

            Finally
                isUpdatingFilters = False
            End Try

            ' ფილტრების გადატარება
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"BtnDeselectAllStatuses_Click: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' 🔧 ნაგულისხმევი სტატუსების მონიშვნის ღილაკი (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnDefaultStatuses_Click(sender As Object, e As EventArgs) Handles BtnDefaultStatuses.Click
        Try
            Debug.WriteLine("UC_BeneficiaryReport: ნაგულისხმევი სტატუსების მონიშვნა")

            isUpdatingFilters = True

            Try
                ' ჯერ ყველა განვიშნოთ
                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing Then
                        checkBox.Checked = False
                    End If
                Next

                ' შემდეგ ნაგულისხმევები მონიშნოთ
                SetDefaultStatusSelection()

            Finally
                isUpdatingFilters = False
            End Try

            ' ფილტრების გადატარება
            ApplyAllFilters()

        Catch ex As Exception
            Debug.WriteLine($"BtnDefaultStatuses_Click: შეცდომა - {ex.Message}")
            isUpdatingFilters = False
        End Try
    End Sub

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            printService?.Dispose()
            allSessions?.Clear()
            filteredSessions?.Clear()
            currentBeneficiarySessions?.Clear()
            statusCheckBoxes?.Clear()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class