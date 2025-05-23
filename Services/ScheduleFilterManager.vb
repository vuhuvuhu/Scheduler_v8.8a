' ===========================================
' 📄 Services/ScheduleFilterManager.vb
' -------------------------------------------
' განრიგის ფილტრების მართვის სერვისი
' UC_Schedule-დან გატანილი ფილტრების ლოგიკა
' ===========================================
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის ფილტრების მართვის სერვისი
    ''' მართავს CheckBox-ებს, ComboBox-ებს, DatePicker-ებს და რადიობუტონებს
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

        ' ღირებულების დელეგატები
        Public Delegate Sub FilterChangedEventHandler()
        Public Delegate Sub BeneficiaryNameChangedEventHandler(selectedName As String)
        Public Delegate Sub ComboBoxRefreshRequestedEventHandler()

        ''' <summary>ფილტრის შეცვლის ღონისძიება</summary>
        Public Event FilterChanged As FilterChangedEventHandler

        ''' <summary>გვერდის ზომის შეცვლის ღონისძიება</summary>
        Public Event PageSizeChanged As FilterChangedEventHandler

        ''' <summary>ბენეფიციარის სახელის შეცვლის ღონისძიება</summary>
        Public Event BeneficiaryNameChanged As BeneficiaryNameChangedEventHandler

        ''' <summary>ComboBox-ების განახლების მოთხოვნის ღონისძიება</summary>
        Public Event ComboBoxRefreshRequested As ComboBoxRefreshRequestedEventHandler

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

                ' ComboBox-ების ივენთების მიბმა
                BindComboBoxEvents()

                ' რადიობუტონების ინიციალიზაცია
                InitializeRadioButtons()

                Debug.WriteLine("ScheduleFilterManager: ფილტრების ინიციალიზაცია დასრულდა")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ინიციალიზაციის შეცდომა: {ex.Message}")
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
                    AddHandler dtpFrom.ValueChanged, AddressOf DatePicker_ValueChanged
                End If

                If dtpTo IsNot Nothing Then
                    dtpTo.Value = lastDayOfMonth
                    AddHandler dtpTo.ValueChanged, AddressOf DatePicker_ValueChanged
                End If

                Debug.WriteLine($"ScheduleFilterManager: თარიღები დაყენებულია {firstDayOfMonth:dd.MM.yyyy} - {lastDayOfMonth:dd.MM.yyyy}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: თარიღების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ების ივენთების მიბმა
        ''' </summary>
        Private Sub BindComboBoxEvents()
            Try
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

                Debug.WriteLine("ScheduleFilterManager: ComboBox ივენთები მიბმულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ComboBox ივენთების მიბმის შეცდომა: {ex.Message}")
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

                ' ივენთების მიბმა
                For Each rb In rbPageSizes
                    If rb IsNot Nothing Then
                        AddHandler rb.CheckedChanged, AddressOf RadioButton_CheckedChanged
                    End If
                Next

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
        ''' ComboBox-ების შევსება დინამიური ფილტრაციით
        ''' მხოლოდ მოცემულ პერიოდში მონაწილე მნიშვნელობები
        ''' </summary>
        ''' <param name="dataProcessor">მონაცემების დამუშავების სერვისი</param>
        Public Sub PopulateFilterComboBoxes(dataProcessor As ScheduleDataProcessor)
            Try
                Debug.WriteLine("ScheduleFilterManager: დინამიური ComboBox-ების შევსება")

                ' მიმდინარე თარიღის პერიოდის მიღება
                Dim dateFrom = If(dtpFrom?.Value.Date, DateTime.Today.AddDays(-30))
                Dim dateTo = If(dtpTo?.Value.Date, DateTime.Today)

                ' ბენეფიციარის სახელები - მხოლოდ მოცემულ პერიოდში
                If cbBeneName IsNot Nothing Then
                    cbBeneName.Items.Clear()
                    Dim names = dataProcessor.GetUniqueValuesForPeriod("BeneficiaryName", dateFrom, dateTo)
                    cbBeneName.Items.AddRange(names.ToArray())
                    cbBeneName.SelectedIndex = 0 ' "ყველა"
                End If

                ' ბენეფიციარის გვარები - დახურული (ცარიელი) სახელის არჩევამდე
                If cbBeneSurname IsNot Nothing Then
                    cbBeneSurname.Items.Clear()
                    cbBeneSurname.Items.Add("ყველა")
                    cbBeneSurname.SelectedIndex = 0
                    cbBeneSurname.Enabled = False ' დახურული
                End If

                ' თერაპევტები - მხოლოდ მოცემულ პერიოდში
                If cbTherapist IsNot Nothing Then
                    cbTherapist.Items.Clear()
                    Dim therapists = dataProcessor.GetUniqueValuesForPeriod("Therapist", dateFrom, dateTo)
                    cbTherapist.Items.AddRange(therapists.ToArray())
                    cbTherapist.SelectedIndex = 0
                End If

                ' თერაპიის ტიპები - მხოლოდ მოცემულ პერიოდში
                If cbTherapyType IsNot Nothing Then
                    cbTherapyType.Items.Clear()
                    Dim therapyTypes = dataProcessor.GetUniqueValuesForPeriod("TherapyType", dateFrom, dateTo)
                    cbTherapyType.Items.AddRange(therapyTypes.ToArray())
                    cbTherapyType.SelectedIndex = 0
                End If

                ' სივრცეები - მხოლოდ მოცემულ პერიოდში
                If cbSpace IsNot Nothing Then
                    cbSpace.Items.Clear()
                    Dim spaces = dataProcessor.GetUniqueValuesForPeriod("Space", dateFrom, dateTo)
                    cbSpace.Items.AddRange(spaces.ToArray())
                    cbSpace.SelectedIndex = 0
                End If

                ' დაფინანსების ტიპები - მხოლოდ მოცემულ პერიოდში
                If cbFunding IsNot Nothing Then
                    cbFunding.Items.Clear()
                    Dim fundingTypes = dataProcessor.GetUniqueValuesForPeriod("Funding", dateFrom, dateTo)
                    cbFunding.Items.AddRange(fundingTypes.ToArray())
                    cbFunding.SelectedIndex = 0
                End If

                Debug.WriteLine("ScheduleFilterManager: დინამიური ComboBox-ები შევსებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: ComboBox-ების შევსების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ბენეფიციარის გვარების ComboBox-ის განახლება არჩეული სახელის მიხედვით
        ''' </summary>
        ''' <param name="selectedName">არჩეული სახელი</param>
        ''' <param name="dataProcessor">მონაცემების დამუშავების სერვისი</param>
        Public Sub UpdateBeneficiarySurnames(selectedName As String, dataProcessor As ScheduleDataProcessor)
            Try
                If cbBeneSurname Is Nothing Then Return

                cbBeneSurname.Items.Clear()

                If String.IsNullOrWhiteSpace(selectedName) OrElse selectedName = "ყველა" Then
                    ' თუ სახელი არ არის არჩეული, გვარების ComboBox დახურული
                    cbBeneSurname.Items.Add("ყველა")
                    cbBeneSurname.SelectedIndex = 0
                    cbBeneSurname.Enabled = False
                    Debug.WriteLine("ScheduleFilterManager: გვარების ComboBox დახურულია")
                Else
                    ' მიმდინარე თარიღის პერიოდის მიღება
                    Dim dateFrom = If(dtpFrom?.Value.Date, DateTime.Today.AddDays(-30))
                    Dim dateTo = If(dtpTo?.Value.Date, DateTime.Today)

                    ' კონკრეტული სახელის შესაბამისი გვარების მიღება
                    Dim surnames = dataProcessor.GetSurnamesForNameInPeriod(selectedName, dateFrom, dateTo)
                    cbBeneSurname.Items.AddRange(surnames.ToArray())
                    cbBeneSurname.SelectedIndex = 0
                    cbBeneSurname.Enabled = True
                    Debug.WriteLine($"ScheduleFilterManager: სახელი '{selectedName}'-ისთვის ნაპოვნია {surnames.Count - 1} გვარი")
                End If

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: UpdateBeneficiarySurnames შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თარიღის შეცვლისას ყველა ComboBox-ის განახლება
        ''' </summary>
        ''' <param name="dataProcessor">მონაცემების დამუშავების სერვისი</param>
        Public Sub RefreshAllComboBoxes(dataProcessor As ScheduleDataProcessor)
            Try
                Debug.WriteLine("ScheduleFilterManager: ყველა ComboBox-ის განახლება თარიღის შეცვლის შემდეგ")
                PopulateFilterComboBoxes(dataProcessor)
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
                criteria.BeneficiaryName = GetComboBoxValue(cbBeneName)
                criteria.BeneficiarySurname = GetComboBoxValue(cbBeneSurname)
                criteria.TherapistName = GetComboBoxValue(cbTherapist)
                criteria.TherapyType = GetComboBoxValue(cbTherapyType)
                criteria.Space = GetComboBoxValue(cbSpace)
                criteria.Funding = GetComboBoxValue(cbFunding)

                Debug.WriteLine($"ScheduleFilterManager: ფილტრის კრიტერიუმები მოძიებულია - თარიღები: {criteria.DateFrom:dd.MM.yyyy} - {criteria.DateTo:dd.MM.yyyy}, სტატუსები: {criteria.SelectedStatuses.Count}")

                Return criteria

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: GetFilterCriteria შეცდომა: {ex.Message}")
                Return New ScheduleDataProcessor.FilterCriteria()
            End Try
        End Function

        ''' <summary>
        ''' ComboBox-ის მნიშვნელობის მიღება
        ''' </summary>
        ''' <param name="comboBox">ComboBox</param>
        ''' <returns>მონიშნული მნიშვნელობა ან ცარიელი სტრიქონი</returns>
        Private Function GetComboBoxValue(comboBox As ComboBox) As String
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
        ''' <returns>მონიშნული სტატუსების სია</returns>
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

                ' ComboBox-ების რესეტი
                Dim comboBoxes As ComboBox() = {cbBeneName, cbBeneSurname, cbTherapist, cbTherapyType, cbSpace, cbFunding}
                For Each cb In comboBoxes
                    If cb IsNot Nothing AndAlso cb.Items.Count > 0 Then
                        cb.SelectedIndex = 0 ' "ყველა"
                    End If
                Next

                ' ყველა სტატუსის მონიშვნა
                SetAllStatusCheckBoxes(True)

                ' პირველი რადიობუტონის მონიშვნა
                If rbPageSizes.Count > 0 AndAlso rbPageSizes(0) IsNot Nothing Then
                    rbPageSizes(0).Checked = True
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

                ' მოთხოვა ComboBox-ების განახლებისთვის
                RaiseEvent ComboBoxRefreshRequested()

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
                Dim selectedName = GetComboBoxValue(cbBeneName)
                Debug.WriteLine($"ScheduleFilterManager: ბენეფიციარის სახელი შეიცვალა: '{selectedName}'")

                ' მოთხოვა გვარების განახლებისთვის
                RaiseEvent BeneficiaryNameChanged(selectedName)

                ' ფილტრის ღონისძიება
                RaiseEvent FilterChanged()
            Catch ex As Exception
                Debug.WriteLine($"ScheduleFilterManager: BeneficiaryName_SelectedIndexChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ComboBox-ის შეცვლის ივენთი
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