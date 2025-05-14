Imports Scheduler_v8._8a.Scheduler_v8_8a.Services
Imports System.Text.RegularExpressions

' ===========================================
' 📄 Forms/AddBeneForm.vb
' -------------------------------------------
' ფორმა ახალი ბენეფიციარის დასამატებლად
' ===========================================

Public Class AddBeneForm
    ' მონაცემთა სერვისის მითითება
    Private ReadOnly dataService As IDataService
    ' დამატების შედეგის სტატუსი - წარმატებულია თუ არა
    Private _isSuccess As Boolean = False
    ' დამატებული ბენეფიციარის სახელი
    Private _addedName As String = ""
    ' დამატებული ბენეფიციარის გვარი
    Private _addedSurname As String = ""

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    ''' <param name="dataService">მონაცემთა სერვისი</param>
    Public Sub New(dataService As IDataService)
        InitializeComponent()
        Me.dataService = dataService
    End Sub

    ''' <summary>
    ''' დამატების ოპერაციის სტატუსი
    ''' </summary>
    Public ReadOnly Property IsSuccess As Boolean
        Get
            Return _isSuccess
        End Get
    End Property

    ''' <summary>
    ''' დამატებული ბენეფიციარის სახელი
    ''' </summary>
    Public ReadOnly Property AddedName As String
        Get
            Return _addedName
        End Get
    End Property

    ''' <summary>
    ''' დამატებული ბენეფიციარის გვარი
    ''' </summary>
    Public ReadOnly Property AddedSurname As String
        Get
            Return _addedSurname
        End Get
    End Property

    ''' <summary>
    ''' ფორმის ჩატვირთვის ივენთი
    ''' </summary>
    Private Sub AddBeneForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' მაქსიმალური ID-ის მიღება და გაზრდა 1-ით
            Dim maxId As Integer = GetMaxBeneId()
            LN.Text = (maxId + 1).ToString()

            ' DateTimePicker-ის კონფიგურაცია ქართული ფორმატით
            ConfigureDateTimePicker()

            ' რადიობატონების საწყისი მდგომარეობა
            RBM.Checked = True  ' ნაგულისხმევად მამრობითი სქესი

            ' სტატუსის ლეიბლის გასუფთავება
            LblStatus.Text = ""

            ' ფოკუსის დაყენება სახელის ველზე
            TName.Focus()

        Catch ex As Exception
            MessageBox.Show($"ფორმის ჩატვირთვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"AddBeneForm_Load შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DateTimePicker-ის კონფიგურაცია ქართული ფორმატით
    ''' </summary>
    Private Sub ConfigureDateTimePicker()
        Try
            ' დავაყენოთ მიმდინარე თარიღი
            DTPBene.Value = DateTime.Today

            ' ფორმატის კონფიგურაცია
            DTPBene.Format = DateTimePickerFormat.Custom
            DTPBene.CustomFormat = "dd.MM.yyyy"

            ' თუ გვინდა ქართული ფონტი (Sylfaen)
            DTPBene.Font = New Font("Sylfaen", 9)

        Catch ex As Exception
            Debug.WriteLine($"ConfigureDateTimePicker შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' იღებს მაქსიმალურ ID-ს DB-Bene ცხრილიდან
    ''' </summary>
    ''' <returns>მაქსიმალური ID</returns>
    Private Function GetMaxBeneId() As Integer
        Try
            ' A სვეტის წაკითხვა DB-Bene ცხრილიდან
            Dim rows = dataService.GetData("DB-Bene!A2:A")
            Dim maxId As Integer = 0

            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    ' თუ მწკრივს აქვს მინიმუმ ერთი სვეტი
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

            Debug.WriteLine($"GetMaxBeneId: მაქსიმალური ID = {maxId}")
            Return maxId
        Catch ex As Exception
            Debug.WriteLine($"GetMaxBeneId შეცდომა: {ex.Message}")
            Return 0 ' შეცდომის შემთხვევაში ვაბრუნებთ 0
        End Try
    End Function

    ''' <summary>
    ''' პირადი ნომრის ფორმატის შემოწმება - მხოლოდ ციფრები, მაქსიმუმ 11 სიმბოლო
    ''' </summary>
    ''' <param name="personalNumber">პირადი ნომერი</param>
    ''' <returns>True თუ ფორმატი სწორია</returns>
    Private Function IsValidPersonalNumber(personalNumber As String) As Boolean
        ' შევამოწმოთ სიცარიელე
        If String.IsNullOrWhiteSpace(personalNumber) Then
            Return True ' ცარიელი ველი დასაშვებია
        End If

        ' მხოლოდ ციფრები და მაქსიმუმ 11 სიმბოლო
        Return Regex.IsMatch(personalNumber, "^\d{1,11}$")
    End Function

    ''' <summary>
    ''' ტელეფონის ნომრის ფორმატის შემოწმება - მხოლოდ ციფრები და + სიმბოლო
    ''' </summary>
    ''' <param name="phoneNumber">ტელეფონის ნომერი</param>
    ''' <returns>True თუ ფორმატი სწორია</returns>
    Private Function IsValidPhoneNumber(phoneNumber As String) As Boolean
        ' შევამოწმოთ სიცარიელე
        If String.IsNullOrWhiteSpace(phoneNumber) Then
            Return True ' ცარიელი ველი დასაშვებია
        End If

        ' მხოლოდ ციფრები და + სიმბოლო
        Return Regex.IsMatch(phoneNumber, "^[0-9+]+$")
    End Function

    ''' <summary>
    ''' ელ-ფოსტის ფორმატის შემოწმება
    ''' </summary>
    ''' <param name="email">ელ-ფოსტა</param>
    ''' <returns>True თუ ფორმატი სწორია</returns>
    Private Function IsValidEmail(email As String) As Boolean
        ' შევამოწმოთ სიცარიელე
        If String.IsNullOrWhiteSpace(email) Then
            Return True ' ცარიელი ველი დასაშვებია
        End If

        ' ელ-ფოსტის ფორმატის რეგულარული გამოსახულება
        Dim pattern As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        Return Regex.IsMatch(email, pattern)
    End Function

    ''' <summary>
    ''' შეყვანილი ველების ვალიდაცია
    ''' </summary>
    ''' <returns>True თუ ფორმატი სწორია</returns>
    Private Function ValidateInputs() As Boolean
        ' სახელი და გვარი აუცილებელია
        If String.IsNullOrWhiteSpace(TName.Text) Then
            LblStatus.Text = "გთხოვთ შეიყვანოთ სახელი"
            LblStatus.ForeColor = Color.Red
            TName.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(TSurname.Text) Then
            LblStatus.Text = "გთხოვთ შეიყვანოთ გვარი"
            LblStatus.ForeColor = Color.Red
            TSurname.Focus()
            Return False
        End If

        ' პირადი ნომერის შემოწმება
        If Not IsValidPersonalNumber(TPN.Text) Then
            LblStatus.Text = "პირადი ნომერი უნდა შეიცავდეს მხოლოდ ციფრებს (მაქს. 11)"
            LblStatus.ForeColor = Color.Red
            TPN.Focus()
            Return False
        End If

        ' მზრუნველის პირადი ნომერის შემოწმება
        If Not IsValidPersonalNumber(TCarmPN.Text) Then
            LblStatus.Text = "მზრუნველის პირ. ნომერი უნდა შეიცავდეს მხოლოდ ციფრებს (მაქს. 11)"
            LblStatus.ForeColor = Color.Red
            TCarmPN.Focus()
            Return False
        End If

        ' ტელეფონის ნომრის შემოწმება
        If Not IsValidPhoneNumber(TTel.Text) Then
            LblStatus.Text = "ტელეფონი უნდა შეიცავდეს მხოლოდ ციფრებს და + სიმბოლოს"
            LblStatus.ForeColor = Color.Red
            TTel.Focus()
            Return False
        End If

        ' ელ-ფოსტის შემოწმება
        If Not IsValidEmail(TMail.Text) Then
            LblStatus.Text = "გთხოვთ შეიყვანოთ სწორი ფორმატის ელ-ფოსტა"
            LblStatus.ForeColor = Color.Red
            TMail.Focus()
            Return False
        End If

        ' ვალიდაცია წარმატებულია
        LblStatus.Text = ""
        Return True
    End Function

    ''' <summary>
    ''' პირადი ნომრის KeyPress ივენთი - მხოლოდ ციფრების დაშვება
    ''' </summary>
    Private Sub TPN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TPN.KeyPress
        ' მხოლოდ ციფრები და წაშლის კლავიში
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' მზრუნველის პირადი ნომრის KeyPress ივენთი - მხოლოდ ციფრების დაშვება
    ''' </summary>
    Private Sub TCarmPN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TCarmPN.KeyPress
        ' მხოლოდ ციფრები და წაშლის კლავიში
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' ტელეფონის ნომრის KeyPress ივენთი - მხოლოდ ციფრებისა და + სიმბოლოს დაშვება
    ''' </summary>
    Private Sub TTel_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TTel.KeyPress
        ' მხოლოდ ციფრები, + სიმბოლო და წაშლის კლავიში
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "+"c AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' "გასუფთავება" ღილაკის Click ივენთი
    ''' </summary>
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        ' ველების გასუფთავება
        TName.Clear()
        TSurname.Clear()
        TPN.Clear()
        DTPBene.Value = DateTime.Today
        RBM.Checked = True
        TCarm.Clear()
        TCarmPN.Clear()
        TTel.Clear()
        TMail.Clear()

        ' სტატუსის ლეიბლის გასუფთავება
        LblStatus.Text = ""

        ' ფოკუსის დაყენება სახელის ველზე
        TName.Focus()
    End Sub

    ''' <summary>
    ''' "დამატება" ღილაკის Click ივენთი
    ''' </summary>
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Try
            ' შევამოწმოთ ველების ვალიდურობა
            If Not ValidateInputs() Then
                Return
            End If

            ' ბენეფიციარის ID
            Dim id As Integer = Integer.Parse(LN.Text)

            ' ფორმატირებული ID-ები 1410 და 1412 პრეფიქსებით
            Dim id1410 As String = "1410-" & id.ToString("000")
            Dim id1412 As String = "1412-" & id.ToString("000")

            ' მწკრივის მომზადება
            Dim rowData As New List(Of Object)

            ' A სვეტი - ID
            rowData.Add(id)

            ' B სვეტი - სახელი
            rowData.Add(TName.Text.Trim())

            ' C სვეტი - გვარი
            rowData.Add(TSurname.Text.Trim())

            ' D სვეტი - 1410 პრეფიქსით ID
            rowData.Add(id1410)

            ' E სვეტი - 1412 პრეფიქსით ID
            rowData.Add(id1412)

            ' F სვეტი - პირადი ნომერი
            rowData.Add(TPN.Text.Trim())

            ' G სვეტი - დაბადების თარიღი
            rowData.Add(DTPBene.Value.ToString("dd.MM.yyyy"))

            ' H სვეტი - სქესი (m ან f)
            rowData.Add(If(RBM.Checked, "m", "f"))

            ' I სვეტი - მზრუნველის სახელი
            rowData.Add(TCarm.Text.Trim())

            ' J სვეტი - მზრუნველის პირადი ნომერი
            rowData.Add(TCarmPN.Text.Trim())

            ' K სვეტი - ამ ეტაპზე ცარიელი
            rowData.Add("")

            ' L სვეტი - ტელეფონი
            rowData.Add(TTel.Text.Trim())

            ' M სვეტი - ელ-ფოსტა
            rowData.Add(TMail.Text.Trim())

            ' შევამოწმოთ, ხომ არ არის უკვე ისეთივე სახელისა და გვარის ბენეფიციარი
            If BeneExists(TName.Text.Trim(), TSurname.Text.Trim()) Then
                If MessageBox.Show("ბენეფიციარი ამ სახელით და გვარით უკვე არსებობს. გსურთ მაინც დამატება?",
                              "დუბლიკატი", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                    Return
                End If
            End If

            ' მონაცემების დამატება DB-Bene ცხრილში
            dataService.AppendData("DB-Bene!A:M", rowData)

            ' შევინახოთ დამატებული ბენეფიციარის ინფორმაცია
            _addedName = TName.Text.Trim()
            _addedSurname = TSurname.Text.Trim()
            _isSuccess = True

            ' შეტყობინება წარმატებული ოპერაციის შესახებ
            LblStatus.Text = "ბენეფიციარი წარმატებით დაემატა!"
            LblStatus.ForeColor = Color.Green

            ' დავხუროთ ფორმა წარმატების კოდით
            MessageBox.Show("ბენეფიციარი წარმატებით დაემატა!", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            LblStatus.Text = "შეცდომა: " & ex.Message
            LblStatus.ForeColor = Color.Red
            Debug.WriteLine($"BtnAdd_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ბენეფიციარის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ამოწმებს არსებობს თუ არა ბენეფიციარი მითითებული სახელითა და გვარით
    ''' </summary>
    ''' <param name="name">სახელი</param>
    ''' <param name="surname">გვარი</param>
    ''' <returns>True თუ ბენეფიციარი უკვე არსებობს</returns>
    Private Function BeneExists(name As String, surname As String) As Boolean
        Try
            ' მონაცემების წაკითხვა DB-Bene ცხრილიდან (B და C სვეტები)
            Dim rows = dataService.GetData("DB-Bene!B2:C")

            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    If row.Count >= 2 AndAlso
                       String.Equals(row(0).ToString().Trim(), name, StringComparison.OrdinalIgnoreCase) AndAlso
                       String.Equals(row(1).ToString().Trim(), surname, StringComparison.OrdinalIgnoreCase) Then
                        Return True ' ბენეფიციარი ნაპოვნია
                    End If
                Next
            End If

            Return False ' ბენეფიციარი არ არსებობს

        Catch ex As Exception
            Debug.WriteLine($"BeneExists შეცდომა: {ex.Message}")
            Return False ' შეცდომის შემთხვევაში ვაბრუნებთ False (ანუ ვამატებთ ახალს)
        End Try
    End Function
End Class