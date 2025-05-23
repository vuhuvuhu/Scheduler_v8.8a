' ===========================================
' 📄 Services/ScheduleFilterManager.vb
' -------------------------------------------
' განრიგის ფილტრების მართვის სერვისი - გაუმჯობესებული ვერსია
' UC_Schedule-დან გატანილი ფილტრების ლოგიკა
' დინამიური ფილტრაცია: მხოლოდ შერჩეულ პერიოდში არსებული მნიშვნელობები
' ===========================================
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის ფილტრების მართვის სერვისი
    ''' მართავს CheckBox-ებს, ComboBox-ებს, DatePicker-ებს და რადიობუტონებს
    ''' ყველა ComboBox შეიცავს მხოლოდ შერჩეულ პერიოდში არსებულ მნიშვნელობებს
    ''' </summary>
    Public Class ScheduleFilterManager

        ' თარიღის კონტროლები
        Private ReadOnly dtpFrom As DateTimePicker
        Private ReadOnly dtpTo As DateTimePicker

        ' ფილტრების ComboBox-ები
        Private ReadOnly cbBeneName As ComboBox
        Private ReadOnly cbBeneSurname As ComboBox
        Private ReadOnly cbTherapist As ComboBox
        Private ReadOnly cbTherapyType As ComboBox
        Private ReadOnly cbSpace As ComboBox
        Private ReadOnly cbFunding As ComboBox

        ' რადიობუტონები გვერდის ზომისთვის
        Private ReadOnly rbPageSizes As List(Of RadioButton)

        ' სტატუსის CheckBox-ები
        Private ReadOnly statusCheckBoxes As List(Of CheckBox)

        ' მონაცემების დამუშავების სერვისი - ცუცავს ლინკი დინამიური განახლებისთვის
        Private dataProcessor As ScheduleDataProcessor = Nothing

        ' ღირებულების დელეგატები
        Public Delegate Sub FilterChangedEventHandler()
        Public Delegate Sub PageSizeChangedEventHandler()

        ''' <summary>ფილტრის შეცვლის ღონისძიება</summary>
        Public Event FilterChanged As FilterChangedEventHandler

        ''' <summary>გვერდის ზომის შეცვლის ღონისძიება</summary>
        Public Event PageSizeChanged As PageSizeChangedEventHandler

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dateFrom">საწყისი თარიღის DatePicker</param>
        ''' <param name="dateTo">საბოლოო თარიღის DatePicker</param>
        ''' <param name="beneficiaryName">ბენეფიციარის სახელის ComboBox</param>
        ''' <param name="beneficiarySurname">ბენეფიციარის გვარის ComboBox</param>
        ''' <param name="therapist">თერაპევტის ComboBox</param>
        ''' <param name="therapyType">თერაპიის ტიპის ComboBox</param>
        ''' <param name="space">სივრცის ComboBox</param>
        ''' <param name="funding">დაფინანსების ComboBox</param>
        ''' <param name="pageSizeButtons">გვერდის ზომის რადიობუტონები</param>
        Public Sub New(dateFrom As DateTimePicker,
                      dateTo As DateTimePicker,
                      beneficiaryName As ComboBox,
                      beneficiarySurname As ComboBox,
                      therapist As ComboBox,
                      therapyType As ComboBox,
                      space As ComboBox,
                      funding As ComboBox,
                      ParamArray pageSizeButtons As RadioButton())

            dtpFrom = dateFrom
            dtpTo = dateTo
            cbBeneName = beneficiaryName
            cbBeneSurname = beneficiarySurname
            cbTherapist = therapist
            cbTherapyType = therapyType
            cbSpace = space
            cbFunding = funding
            rbPageSizes = pageSizeButtons.ToList()
            statusCheckBoxes = New List(Of CheckBox)()
        End Sub

        ''' <summary>
        ''' ფილტრების საწყისი ინიციალიზაცია
        ''' </summary>
        Public Sub InitializeFilters()
            Try
                Debug.WriteLine("ScheduleFilterManager: ფილტრების ინიციალიზაცია")

                ' თარიღების ინიციალიზაცია
                InitializeDatePickers()

                ' ComboBox-ების საწყისი მომზადება
                InitializeComboBoxes()

                ' რადიობუტონების ინიციალიზაცია
                InitializeRadioButtons()

                ' ივენთების მიბმა
                BindEvents()

                Debug.WriteLine("ScheduleFilterManager: ფილტრების ინიციალიზაცია დასრულდა")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ების საწყისი ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeComboBoxes()
            Try
                ' ყველა ComboBox-ის გასუფთავება და "ყველა" ვარიანტის დამატება
                Dim comboBoxes As ComboBox() = {cbBeneName, cbBeneSurname, cbTherapist, cbTherapyType, cbSpace, cbFunding}

                For Each cb In comboBoxes
                    If cb IsNot Nothing Then
                        cb.Items.Clear()
                        cb.Items.Add("ყველა")
                        cb.SelectedIndex = 0
                    End If
                Next

                ' ბენეფიციარის გვარის ComboBox-ის დახურვა საწყისად
                If cbBeneSurname IsNot Nothing Then
                    cbBeneSurname.Enabled = False
                    Debug.WriteLine("ScheduleFilterManager: ბენეფიციარის გვარის ComboBox დახურულია საწყისად")
                End If

                Debug.WriteLine("ScheduleFilterManager: ComboBox-ები ინიციალიზებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ComboBox-ების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თარიღის DatePicker-ების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeDatePickers()
            Try
                ' მიმდინარე თარიღი
                Dim currentDate As DateTime = DateTime.Today

                ' მიმდინარე თვის პირველი და ბოლო დღე
                Dim firstDayOfMonth As DateTime = New DateTime(currentDate.Year, currentDate.Month, 1)
                Dim lastDayOfMonth As DateTime = firstDayOfMonth.AddMonths(1).AddDays(-1)

                ' DatePicker-ების დაყენება
                If dtpFrom IsNot Nothing Then
                    dtpFrom.Value = firstDayOfMonth
                End If

                If dtpTo IsNot Nothing Then
                    dtpTo.Value = lastDayOfMonth
                End If

                Debug.WriteLine($"ScheduleFilterManager: თარიღები დაყენებულია {firstDayOfMonth:dd.MM.yyyy} - {lastDayOfMonth:dd.MM.yyyy}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: თარიღების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ივენთების მიბმა
        ''' </summary>
        Private Sub BindEvents()
            Try
                ' თარიღის DatePicker-ების ივენთები
                If dtpFrom IsNot Nothing Then
                    AddHandler dtpFrom.ValueChanged, AddressOf DatePicker_ValueChanged
                End If

                If dtpTo IsNot Nothing Then
                    AddHandler dtpTo.ValueChanged, AddressOf DatePicker_ValueChanged
                End If

                ' ბენეფიციარის სახელის ComboBox - სპეციალური ხელი
                If cbBeneName IsNot Nothing Then
                    AddHandler cbBeneName.SelectedIndexChanged, AddressOf BeneficiaryName_SelectedIndexChanged
                End If

                ' დანარჩენი ComboBox-ები
                If cbBeneSurname IsNot Nothing Then AddHandler cbBeneSurname.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                If cbTherapist IsNot Nothing Then AddHandler cbTherapist.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                If cbTherapyType IsNot Nothing Then AddHandler cbTherapyType.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                If cbSpace IsNot Nothing Then AddHandler cbSpace.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                If cbFunding IsNot Nothing Then AddHandler cbFunding.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged

                ' რადიობუტონების ივენთები
                For Each rb In rbPageSizes
                    If rb IsNot Nothing Then
                        AddHandler rb.CheckedChanged, AddressOf RadioButton_CheckedChanged
                    End If
                Next

                Debug.WriteLine("ScheduleFilterManager: ივენთები მიბმულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ივენთების მიბმის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' რადიობუტონების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeRadioButtons()
            Try
                ' პირველი რადიობუტონის მონიშვნა (ნაგულისხმევად)
                If rbPageSizes.Count > 0 AndAlso rbPageSizes(0) IsNot Nothing Then
                    rbPageSizes(0).Checked = True
                End If

                Debug.WriteLine($"ScheduleFilterManager: {rbPageSizes.Count} რადიობუტონი ინიციალიზებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: რადიობუტონების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სტატუსის CheckBox-ების დამატება და ინიციალიზაცია
        ''' </summary>
        ''' <param name="checkBoxes">CheckBox-ების სია</param>
        Public Sub InitializeStatusCheckBoxes(ParamArray checkBoxes As CheckBox())
            Try
                statusCheckBoxes.Clear()
                statusCheckBoxes.AddRange(checkBoxes.Where(Function(cb) cb IsNot Nothing))

                ' ყველა CheckBox-ის მონიშვნა და ივენთების მიბმა
                For Each checkBox In statusCheckBoxes
                    checkBox.Checked = True ' ყველა მონიშნული ჩატვირთვისას
                    AddHandler checkBox.CheckedChanged, AddressOf StatusCheckBox_CheckedChanged
                    Debug.WriteLine($"ScheduleFilterManager: CheckBox '{checkBox.Text}' ინიციალიზებული")
                Next

                Debug.WriteLine($"ScheduleFilterManager: {statusCheckBoxes.Count} სტატუსის CheckBox ინიციალიზებული")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: სტატუსის CheckBox-ების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ების შევსება დინამიური ფილტრაციით - მთავარი მეთოდი
        ''' მხოლოდ მოცემულ პერიოდში მონაწილე მნიშვნელობები ComboBox-ებში
        ''' </summary>
        ''' <param name="processor">მონაცემების დამუშავების სერვისი</param>
        Public Sub PopulateFilterComboBoxes(processor As ScheduleDataProcessor)
            Try
                Debug.WriteLine("ScheduleFilterManager: დინამიური ComboBox-ების შევსება დაიწყო")

                ' შევინახოთ მონაცემების პროცესორი შემდგომი გამოყენებისთვის
                dataProcessor = processor

                If dataProcessor Is Nothing Then
                    Debug.WriteLine("ScheduleFilterManager: dataProcessor არის Nothing, ComboBox-ების შევსება შეუძლებელია")
                    Return
                End If

                ' მიმდინარე თარიღის პერიოდის მიღება
                Dim dateFrom = If(dtpFrom?.Value.Date, DateTime.Today.AddDays(-30))
                Dim dateTo = If(dtpTo?.Value.Date, DateTime.Today)

                Debug.WriteLine($"ScheduleFilterManager: ComboBox-ების შევსება პერიოდისთვის {dateFrom:dd.MM.yyyy} - {dateTo:dd.MM.yyyy}")

                ' ყველა ComboBox-ის შევსება (გვარის გარდა)
                PopulateBeneficiaryNamesComboBox(dateFrom, dateTo)
                PopulateTherapistComboBox(dateFrom, dateTo)
                PopulateTherapyTypeComboBox(dateFrom, dateTo)
                PopulateSpaceComboBox(dateFrom, dateTo)
                PopulateFundingComboBox(dateFrom, dateTo)

                ' ბენეფიციარის გვარის ComboBox-ის საწყისი მდგომარეობა
                ResetBeneficiarySurnameComboBox()

                Debug.WriteLine("ScheduleFilterManager: დინამიური ComboBox-ები შევსებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ComboBox-ების შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ბენეფიციარის სახელების ComboBox-ის შევსება
        ''' </summary>
        Private Sub PopulateBeneficiaryNamesComboBox(dateFrom As Date, dateTo As Date)
            Try
                If cbBeneName Is Nothing OrElse dataProcessor Is Nothing Then Return

                ' მიმდინარე არჩეული მნიშვნელობის შენახვა
                Dim selectedValue = GetComboBoxSelectedValue(cbBeneName)

                ' ComboBox-ის გასუფთავება
                cbBeneName.Items.Clear()

                ' უნიკალური სახელების მიღება მოცემული პერიოდისთვის
                Dim names = dataProcessor.GetUniqueValuesForPeriod("BeneficiaryName", dateFrom, dateTo)

                ' ComboBox-ის შევსება
                cbBeneName.Items.AddRange(names.ToArray())

                ' არჩეული მნიშვნელობის აღდგენა ან "ყველა"-ს დაყენება
                SetComboBoxSelectedValue(cbBeneName, selectedValue)

                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის სახელების ComboBox შევსებულია - {names.Count} ელემენტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის სახელების ComboBox-ის შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თერაპევტის ComboBox-ის შევსება
        ''' </summary>
        Private Sub PopulateTherapistComboBox(dateFrom As Date, dateTo As Date)
            Try
                If cbTherapist Is Nothing OrElse dataProcessor Is Nothing Then Return

                Dim selectedValue = GetComboBoxSelectedValue(cbTherapist)
                cbTherapist.Items.Clear()

                Dim therapists = dataProcessor.GetUniqueValuesForPeriod("Therapist", dateFrom, dateTo)
                cbTherapist.Items.AddRange(therapists.ToArray())

                SetComboBoxSelectedValue(cbTherapist, selectedValue)

                Debug.WriteLine($"ScheduleFilterManager: თერაპევტის ComboBox შევსებულია - {therapists.Count} ელემენტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: თერაპევტის ComboBox-ის შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თერაპიის ტიპის ComboBox-ის შევსება
        ''' </summary>
        Private Sub PopulateTherapyTypeComboBox(dateFrom As Date, dateTo As Date)
            Try
                If cbTherapyType Is Nothing OrElse dataProcessor Is Nothing Then Return

                Dim selectedValue = GetComboBoxSelectedValue(cbTherapyType)
                cbTherapyType.Items.Clear()

                Dim therapyTypes = dataProcessor.GetUniqueValuesForPeriod("TherapyType", dateFrom, dateTo)
                cbTherapyType.Items.AddRange(therapyTypes.ToArray())

                SetComboBoxSelectedValue(cbTherapyType, selectedValue)

                Debug.WriteLine($"ScheduleFilterManager: თერაპიის ტიპის ComboBox შევსებულია - {therapyTypes.Count} ელემენტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: თერაპიის ტიპის ComboBox-ის შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სივრცის ComboBox-ის შევსება
        ''' </summary>
        Private Sub PopulateSpaceComboBox(dateFrom As Date, dateTo As Date)
            Try
                If cbSpace Is Nothing OrElse dataProcessor Is Nothing Then Return

                Dim selectedValue = GetComboBoxSelectedValue(cbSpace)
                cbSpace.Items.Clear()

                Dim spaces = dataProcessor.GetUniqueValuesForPeriod("Space", dateFrom, dateTo)
                cbSpace.Items.AddRange(spaces.ToArray())

                SetComboBoxSelectedValue(cbSpace, selectedValue)

                Debug.WriteLine($"ScheduleFilterManager: სივრცის ComboBox შევსებულია - {spaces.Count} ელემენტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: სივრცის ComboBox-ის შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' დაფინანსების ComboBox-ის შევსება
        ''' </summary>
        Private Sub PopulateFundingComboBox(dateFrom As Date, dateTo As Date)
            Try
                If cbFunding Is Nothing OrElse dataProcessor Is Nothing Then Return

                Dim selectedValue = GetComboBoxSelectedValue(cbFunding)
                cbFunding.Items.Clear()

                Dim fundingTypes = dataProcessor.GetUniqueValuesForPeriod("Funding", dateFrom, dateTo)
                cbFunding.Items.AddRange(fundingTypes.ToArray())

                SetComboBoxSelectedValue(cbFunding, selectedValue)

                Debug.WriteLine($"ScheduleFilterManager: დაფინანსების ComboBox შევსებულია - {fundingTypes.Count} ელემენტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: დაფინანსების ComboBox-ის შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ბენეფიციარის გვარების ComboBox-ის რესეტი - დახურული მდგომარეობაში
        ''' </summary>
        Private Sub ResetBeneficiarySurnameComboBox()
            Try
                If cbBeneSurname Is Nothing Then Return

                cbBeneSurname.Items.Clear()
                cbBeneSurname.Items.Add("ყველა")
                cbBeneSurname.SelectedIndex = 0
                cbBeneSurname.Enabled = False

                Debug.WriteLine("ScheduleFilterManager: ბენეფიციარის გვარის ComboBox დარესეტებულია და დახურულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის გვარის ComboBox-ის რესეტის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ბენეფიციარის გვარების ComboBox-ის განახლება არჩეული სახელის მიხედვით
        ''' </summary>
        ''' <param name="selectedName">არჩეული სახელი</param>
        Public Sub UpdateBeneficiarySurnames(selectedName As String)
            Try
                If cbBeneSurname Is Nothing OrElse dataProcessor Is Nothing Then Return

                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის გვარების განახლება სახელისთვის: '{selectedName}'")

                cbBeneSurname.Items.Clear()

                If String.IsNullOrWhiteSpace(selectedName) OrElse selectedName = "ყველა" Then
                    ' თუ სახელი არ არის არჩეული, გვარების ComboBox დახურული
                    cbBeneSurname.Items.Add("ყველა")
                    cbBeneSurname.SelectedIndex = 0
                    cbBeneSurname.Enabled = False
                    Debug.WriteLine("ScheduleFilterManager: გვარების ComboBox დახურულია - სახელი არ არის არჩეული")
                Else
                    ' მიმდინარე თარიღის პერიოდის მიღება
                    Dim dateFrom = If(dtpFrom?.Value.Date, DateTime.Today.AddDays(-30))
                    Dim dateTo = If(dtpTo?.Value.Date, DateTime.Today)

                    ' კონკრეტული სახელის შესაბამისი გვარების მიღება
                    Dim surnames = dataProcessor.GetSurnamesForNameInPeriod(selectedName, dateFrom, dateTo)
                    cbBeneSurname.Items.AddRange(surnames.ToArray())
                    cbBeneSurname.SelectedIndex = 0
                    cbBeneSurname.Enabled = True

                    Debug.WriteLine($"ScheduleFilterManager: სახელი '{selectedName}'-ისთვის ნაპოვნია {surnames.Count - 1} გვარი, ComboBox გააქტიურებულია")
                End If

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: UpdateBeneficiarySurnames შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ის მიმდინარე არჩეული მნიშვნელობის მიღება
        ''' </summary>
        Private Function GetComboBoxSelectedValue(comboBox As ComboBox) As String
            Try
                If comboBox IsNot Nothing AndAlso comboBox.SelectedItem IsNot Nothing Then
                    Return comboBox.SelectedItem.ToString()
                End If
                Return "ყველა"
            Catch
                Return "ყველა"
            End Try
        End Function

        ''' <summary>
        ''' ComboBox-ის არჩეული მნიშვნელობის დაყენება
        ''' </summary>
        Private Sub SetComboBoxSelectedValue(comboBox As ComboBox, value As String)
            Try
                If comboBox Is Nothing OrElse String.IsNullOrEmpty(value) Then Return

                ' ვეცდებით იპოვნოთ მნიშვნელობა ComboBox-ში
                For i As Integer = 0 To comboBox.Items.Count - 1
                    If String.Equals(comboBox.Items(i).ToString(), value, StringComparison.OrdinalIgnoreCase) Then
                        comboBox.SelectedIndex = i
                        Return
                    End If
                Next

                ' თუ არ მოიძებნა, დავაყენოთ "ყველა"
                comboBox.SelectedIndex = 0

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: SetComboBoxSelectedValue შეცდომა: {ex.Message}")
                If comboBox IsNot Nothing AndAlso comboBox.Items.Count > 0 Then
                    comboBox.SelectedIndex = 0
                End If
            End Try
        End Sub

        ''' <summary>
        ''' ყველა ComboBox-ის განახლება თარიღის შეცვლის შემდეგ
        ''' </summary>
        Public Sub RefreshAllComboBoxes()
            Try
                Debug.WriteLine("ScheduleFilterManager: ყველა ComboBox-ის განახლება თარიღის შეცვლის შემდეგ")

                If dataProcessor IsNot Nothing Then
                    PopulateFilterComboBoxes(dataProcessor)
                Else
                    Debug.WriteLine("ScheduleFilterManager: dataProcessor არის Nothing - ComboBox-ების განახლება შეუძლებელია")
                End If

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: RefreshAllComboBoxes შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ფილტრის კრიტერიუმების მიღება მომხმარებლის ჩანაწერებიდან
        ''' </summary>
        ''' <returns>ფილტრის კრიტერიუმები</returns>
        Public Function GetFilterCriteria() As ScheduleDataProcessor.FilterCriteria
            Try
                Dim criteria As New ScheduleDataProcessor.FilterCriteria()

                ' თარიღების მიღება
                criteria.DateFrom = If(dtpFrom?.Value.Date, DateTime.Today.AddDays(-30))
                criteria.DateTo = If(dtpTo?.Value.Date, DateTime.Today)

                ' მონიშნული სტატუსების მიღება
                criteria.SelectedStatuses = GetSelectedStatuses()

                ' ComboBox-ების მნიშვნელობები
                criteria.BeneficiaryName = GetComboBoxFilterValue(cbBeneName)
                criteria.BeneficiarySurname = GetComboBoxFilterValue(cbBeneSurname)
                criteria.TherapistName = GetComboBoxFilterValue(cbTherapist)
                criteria.TherapyType = GetComboBoxFilterValue(cbTherapyType)
                criteria.Space = GetComboBoxFilterValue(cbSpace)
                criteria.Funding = GetComboBoxFilterValue(cbFunding)

                Debug.WriteLine($"ScheduleFilterManager: ფილტრის კრიტერიუმები - თარიღები: {criteria.DateFrom:dd.MM.yyyy}-{criteria.DateTo:dd.MM.yyyy}, " &
                               $"სტატუსები: {criteria.SelectedStatuses.Count}, სახელი: '{criteria.BeneficiaryName}', გვარი: '{criteria.BeneficiarySurname}'")

                Return criteria

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: GetFilterCriteria შეცდომა: {ex.Message}")
                Return New ScheduleDataProcessor.FilterCriteria()
            End Try
        End Function

        ''' <summary>
        ''' ComboBox-ის ფილტრის მნიშვნელობის მიღება ("ყველა" = ცარიელი სტრიქონი)
        ''' </summary>
        Private Function GetComboBoxFilterValue(comboBox As ComboBox) As String
            Try
                If comboBox IsNot Nothing AndAlso comboBox.SelectedItem IsNot Nothing Then
                    Dim selectedValue = comboBox.SelectedItem.ToString()
                    Return If(selectedValue = "ყველა", "", selectedValue)
                End If
                Return ""
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' მონიშნული სტატუსების სიის მიღება
        ''' </summary>
        Private Function GetSelectedStatuses() As List(Of String)
            Dim selectedStatuses As New List(Of String)

            Try
                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing AndAlso checkBox.Checked Then
                        Dim statusText = checkBox.Text.Trim()
                        If Not String.IsNullOrEmpty(statusText) Then
                            selectedStatuses.Add(statusText)
                        End If
                    End If
                Next

                Debug.WriteLine($"ScheduleFilterManager: მონიშნული სტატუსები: {String.Join(", ", selectedStatuses)}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: GetSelectedStatuses შეცდომა: {ex.Message}")
            End Try

            Return selectedStatuses
        End Function

        ''' <summary>
        ''' გვერდის ზომის მიღება მონიშნული რადიობუტონიდან
        ''' </summary>
        ''' <returns>გვერდის ზომა (ნაგულისხმევად 20)</returns>
        Public Function GetPageSize() As Integer
            Try
                For Each rb In rbPageSizes
                    If rb IsNot Nothing AndAlso rb.Checked Then
                        ' რადიობუტონის სახელიდან ზომის განსაზღვრა
                        Select Case rb.Name.ToUpper()
                            Case "RB20"
                                Return 20
                            Case "RB50"
                                Return 50
                            Case "RB100"
                                Return 100
                            Case Else
                                ' ტექსტიდან რიცხვის მოძებნის მცდელობა
                                Dim numberText = System.Text.RegularExpressions.Regex.Match(rb.Text, "\d+").Value
                                If Not String.IsNullOrEmpty(numberText) Then
                                    Dim size As Integer
                                    If Integer.TryParse(numberText, size) Then
                                        Return size
                                    End If
                                End If
                        End Select
                    End If
                Next

                Return 20 ' ნაგულისხმევი ზომა

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: GetPageSize შეცდომა: {ex.Message}")
                Return 20
            End Try
        End Function

        ''' <summary>
        ''' ყველა სტატუსის CheckBox-ის მონიშვნა/განიშვნა
        ''' </summary>
        ''' <param name="checkAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
        Public Sub SetAllStatusCheckBoxes(checkAll As Boolean)
            Try
                Debug.WriteLine($"ScheduleFilterManager: ყველა CheckBox-ის დაყენება: {checkAll}")

                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing Then
                        checkBox.Checked = checkAll
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: SetAllStatusCheckBoxes შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' კონკრეტული სტატუსის CheckBox-ის მონიშვნა/განიშვნა
        ''' </summary>
        ''' <param name="statusText">სტატუსის ტექსტი</param>
        ''' <param name="isChecked">მონიშნული თუ არა</param>
        Public Sub SetStatusCheckBox(statusText As String, isChecked As Boolean)
            Try
                For Each checkBox In statusCheckBoxes
                    If checkBox IsNot Nothing AndAlso
                       String.Equals(checkBox.Text.Trim(), statusText.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        checkBox.Checked = isChecked
                        Debug.WriteLine($"ScheduleFilterManager: CheckBox '{statusText}' დაყენებული: {isChecked}")
                        Exit For
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: SetStatusCheckBox შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონიშნული სტატუსების რაოდენობის მიღება
        ''' </summary>
        ''' <returns>მონიშნული CheckBox-ების რაოდენობა</returns>
        Public ReadOnly Property SelectedStatusCount As Integer
            Get
                Try
                    Return statusCheckBoxes.Where(Function(cb) cb IsNot Nothing AndAlso cb.Checked).Count()
                Catch
                    Return 0
                End Try
            End Get
        End Property

        ''' <summary>
        ''' ფილტრების რესეტი - ყველაფრის საწყის მდგომარეობაში დაბრუნება
        ''' </summary>
        Public Sub ResetFilters()
            Try
                Debug.WriteLine("ScheduleFilterManager: ფილტრების რესეტი")

                ' თარიღების რესეტი
                InitializeDatePickers()

                ' ComboBox-ების რესეტი - ყველა "ყველა"-ზე დაყენება
                Dim comboBoxes As ComboBox() = {cbBeneName, cbBeneSurname, cbTherapist, cbTherapyType, cbSpace, cbFunding}
                For Each cb In comboBoxes
                    If cb IsNot Nothing AndAlso cb.Items.Count > 0 Then
                        cb.SelectedIndex = 0 ' "ყველა"
                    End If
                Next

                ' ბენეფიციარის გვარის ComboBox-ის დახურვა
                ResetBeneficiarySurnameComboBox()

                ' ყველა სტატუსის მონიშვნა
                SetAllStatusCheckBoxes(True)

                ' პირველი რადიობუტონის მონიშვნა
                If rbPageSizes.Count > 0 AndAlso rbPageSizes(0) IsNot Nothing Then
                    rbPageSizes(0).Checked = True
                End If

                ' ComboBox-ების განახლება თუ dataProcessor ხელმისაწვდომია
                If dataProcessor IsNot Nothing Then
                    PopulateFilterComboBoxes(dataProcessor)
                End If

                Debug.WriteLine("ScheduleFilterManager: ფილტრების რესეტი დასრულდა")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ResetFilters შეცდომა: {ex.Message}")
            End Try
        End Sub

#Region "ივენთ ჰენდლერები"

        ''' <summary>
        ''' თარიღის DatePicker-ის შეცვლის ივენთი
        ''' </summary>
        Private Sub DatePicker_ValueChanged(sender As Object, e As EventArgs)
            Try
                Debug.WriteLine("ScheduleFilterManager: თარიღის ფილტრი შეიცვალა - ComboBox-ების განახლება")

                ' ComboBox-ების განახლება ახალი თარიღის პერიოდისთვის
                RefreshAllComboBoxes()

                ' ფილტრის ღონისძიება
                RaiseEvent FilterChanged()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: DatePicker_ValueChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ბენეფიციარის სახელის ComboBox-ის შეცვლის ივენთი
        ''' </summary>
        Private Sub BeneficiaryName_SelectedIndexChanged(sender As Object, e As EventArgs)
            Try
                Dim selectedName = GetComboBoxFilterValue(cbBeneName)

                ' "ყველა" = ცარიელი სტრიქონი, ასე რომ ვიყენებთ ზუსტ მნიშვნელობას
                Dim actualSelectedName = If(cbBeneName?.SelectedItem?.ToString(), "ყველა")

                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის სახელი შეიცვალა: '{actualSelectedName}'")

                ' ბენეფიციარის გვარების ComboBox-ის განახლება
                UpdateBeneficiarySurnames(actualSelectedName)

                ' ფილტრის ღონისძიება
                RaiseEvent FilterChanged()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: BeneficiaryName_SelectedIndexChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ის შეცვლის ივენთი (ბენეფიციარის სახელის გარდა)
        ''' </summary>
        Private Sub ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            Try
                Debug.WriteLine("ScheduleFilterManager: ComboBox ფილტრი შეიცვალა")
                RaiseEvent FilterChanged()
            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ComboBox_SelectedIndexChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სტატუსის CheckBox-ის შეცვლის ივენთი
        ''' </summary>
        Private Sub StatusCheckBox_CheckedChanged(sender As Object, e As EventArgs)
            Try
                Debug.WriteLine("ScheduleFilterManager: სტატუსის CheckBox შეიცვალა")
                RaiseEvent FilterChanged()
            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: StatusCheckBox_CheckedChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' რადიობუტონის შეცვლის ივენთი
        ''' </summary>
        Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
            Try
                Dim radioButton As RadioButton = CType(sender, RadioButton)
                If radioButton.Checked Then
                    Debug.WriteLine($"ScheduleFilterManager: გვერდის ზომა შეიცვალა: {GetPageSize()}")
                    RaiseEvent PageSizeChanged()
                End If
            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: RadioButton_CheckedChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

#End Region

    End Class
End Namespace