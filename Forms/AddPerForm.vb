' ===========================================
' 📄 Forms/AddPerForm.vb
' -------------------------------------------
' თერაპევტის დამატების ფორმა - ინფორმაცია იწერება DB-Personal ფურცელზე
' ===========================================
Imports System.Text.RegularExpressions
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class AddPerForm

    ' მონაცემთა სერვისი
    Private ReadOnly dataService As IDataService
    ' დამატების წარმატების ფლაგი
    Public Property IsSuccess As Boolean = False
    ' დამატებული თერაპევტის სახელი და გვარი
    Public Property AddedName As String = ""
    Public Property AddedSurname As String = ""

    ''' <summary>
    ''' კონსტრუქტორი - მონაცემთა სერვისის მითითებით
    ''' </summary>
    ''' <param name="dataService">მონაცემთა სერვისი</param>
    Public Sub New(dataService As IDataService)
        ' დიზაინერით გენერირებული კოდი
        InitializeComponent()

        ' მონაცემთა სერვისის მითითება
        Me.dataService = dataService

        ' ფორმის ზედა ზოლში მხოლოდ დახურვის ღილაკის დატოვება
        Me.MinimizeBox = False
        Me.MaximizeBox = False

        ' ფორმის სათაური
        Me.Text = "თერაპევტის დამატება"
    End Sub

    ''' <summary>
    ''' ფორმის ჩატვირთვის ივენთი
    ''' </summary>
    Private Sub AddPerForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' ახალი ID-ის მიღება
        LN.Text = (GetMaxRecordId() + 1).ToString()

        ' DateTimePicker-ის საწყისი თარიღის დაყენება 
        DTPPer.Value = DateTime.Today.AddYears(-30) ' ნაგულისხმევად 30 წლის

        ' აქტიური სტატუსის არჩევა ნაგულისხმევად
        RBActive.Checked = True

        ' ფორმის ვალიდაცია
        ValidateForm()
    End Sub

    ''' <summary>
    ''' მაქსიმალური ID-ის მოძიება DB-Personal ფურცლიდან
    ''' </summary>
    Private Function GetMaxRecordId() As Integer
        Try
            ' A სვეტის წაკითხვა DB-Personal ფურცლიდან
            Dim rows = dataService.GetData("DB-Personal!A2:A")
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
            Debug.WriteLine($"GetMaxRecordId შეცდომა: {ex.Message}")
            MessageBox.Show($"ID-ის მოძიების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' ამოწმებს არსებობს თუ არა უკვე მოცემული სახელისა და გვარის თერაპევტი
    ''' </summary>
    ''' <returns>True თუ არსებობს, False თუ არ არსებობს</returns>
    Private Function CheckTherapistExists(name As String, surname As String) As Boolean
        Try
            ' B და C სვეტების წაკითხვა DB-Personal ფურცლიდან
            Dim rows = dataService.GetData("DB-Personal!B2:C")

            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    If row.Count >= 2 Then
                        ' შევამოწმოთ ემთხვევა თუ არა სახელი და გვარი
                        Dim rowName = row(0).ToString().Trim()
                        Dim rowSurname = row(1).ToString().Trim()

                        If String.Equals(rowName, name, StringComparison.OrdinalIgnoreCase) AndAlso
                           String.Equals(rowSurname, surname, StringComparison.OrdinalIgnoreCase) Then
                            Return True ' ნაპოვნია დამთხვევა
                        End If
                    End If
                Next
            End If

            Return False ' არ არის ნაპოვნი დამთხვევა
        Catch ex As Exception
            Debug.WriteLine($"CheckTherapistExists შეცდომა: {ex.Message}")
            ' შეცდომის შემთხვევაში ვაბრუნებთ False-ს, რომ არ შევაფერხოთ პროცესი
            Return False
        End Try
    End Function

    ''' <summary>
    ''' TName ველის TextChanged ივენთი
    ''' </summary>
    Private Sub TName_TextChanged(sender As Object, e As EventArgs) Handles TName.TextChanged
        ValidateForm()
    End Sub

    ''' <summary>
    ''' TSurname ველის TextChanged ივენთი
    ''' </summary>
    Private Sub TSurname_TextChanged(sender As Object, e As EventArgs) Handles TSurname.TextChanged
        ValidateForm()
    End Sub

    ''' <summary>
    ''' პირადი ნომრის ველის KeyPress ივენთი
    ''' - მხოლოდ ციფრების შეყვანის კონტროლი
    ''' </summary>
    Private Sub TPN_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TPN.KeyPress
        ' დავუშვათ მხოლოდ ციფრები და წაშლის (Backspace) ღილაკი
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' პირადი ნომრის ველის TextChanged ივენთი
    ''' - სიგრძის კონტროლი (მაქსიმუმ 11 სიმბოლო)
    ''' </summary>
    Private Sub TPN_TextChanged(sender As Object, e As EventArgs) Handles TPN.TextChanged
        ' შევამოწმოთ შეყვანილი პირადი ნომერი
        Dim pn = TPN.Text.Trim()

        ' მხოლოდ ციფრები და მაქსიმუმ 11 სიმბოლო
        If Not String.IsNullOrEmpty(pn) AndAlso Not Regex.IsMatch(pn, "^\d+$") Then
            TPN.BackColor = Color.MistyRose ' შეცდომის მითითება
        ElseIf pn.Length > 0 AndAlso pn.Length < 11 Then
            TPN.BackColor = Color.MistyRose ' შეცდომის მითითება
        Else
            TPN.BackColor = SystemColors.Window ' ნორმალური ფონი
        End If

        ValidateForm()
    End Sub

    ''' <summary>
    ''' ტელეფონის ნომრის ველის TextChanged ივენთი
    ''' </summary>
    Private Sub TTel_TextChanged(sender As Object, e As EventArgs) Handles TTel.TextChanged
        ' შევამოწმოთ ტელეფონის ნომერი
        Dim tel = TTel.Text.Trim()

        ' ტელეფონის ნომრის მარტივი ვალიდაცია: მხოლოდ ციფრები, +-() და გარკვეული სიმბოლოები
        If Not String.IsNullOrEmpty(tel) AndAlso Not Regex.IsMatch(tel, "^[0-9\+\-\(\)\s]+$") Then
            TTel.BackColor = Color.MistyRose ' შეცდომის მითითება
        Else
            TTel.BackColor = SystemColors.Window ' ნორმალური ფონი
        End If

        ValidateForm()
    End Sub

    ''' <summary>
    ''' ელ-ფოსტის ველის TextChanged ივენთი
    ''' </summary>
    Private Sub TMail_TextChanged(sender As Object, e As EventArgs) Handles TMail.TextChanged
        ' შევამოწმოთ ელ-ფოსტა
        Dim email = TMail.Text.Trim()

        ' ელ-ფოსტის ვალიდაცია რეგულარული გამოსახულებით
        If Not String.IsNullOrEmpty(email) AndAlso Not Regex.IsMatch(email, "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$") Then
            TMail.BackColor = Color.MistyRose ' შეცდომის მითითება
        Else
            TMail.BackColor = SystemColors.Window ' ნორმალური ფონი
        End If

        ValidateForm()
    End Sub

    ''' <summary>
    ''' ფორმის ვალიდაცია - ამოწმებს ყველა ველს
    ''' </summary>
    Private Sub ValidateForm()
        ' სავალდებულო ველების შემოწმება - სახელი და გვარი
        Dim isValid As Boolean = Not String.IsNullOrWhiteSpace(TName.Text) AndAlso
                                Not String.IsNullOrWhiteSpace(TSurname.Text)

        ' ველების ვალიდაცია
        ' პირადი ნომერი - თუ შევსებულია, უნდა იყოს 11 ციფრი
        If Not String.IsNullOrEmpty(TPN.Text.Trim()) Then
            isValid = isValid AndAlso Regex.IsMatch(TPN.Text.Trim(), "^\d{11}$")
        End If

        ' ტელეფონის ნომერი - თუ შევსებულია, უნდა იყოს ვალიდური ფორმატი
        If Not String.IsNullOrEmpty(TTel.Text.Trim()) Then
            isValid = isValid AndAlso Regex.IsMatch(TTel.Text.Trim(), "^[0-9\+\-\(\)\s]+$")
        End If

        ' ელ-ფოსტა - თუ შევსებულია, უნდა იყოს ვალიდური ფორმატი
        If Not String.IsNullOrEmpty(TMail.Text.Trim()) Then
            isValid = isValid AndAlso Regex.IsMatch(TMail.Text.Trim(), "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$")
        End If

        ' შევამოწმოთ მსგავსი ჩანაწერი აქვე, თუ სახელი და გვარი უკვე შევსებულია
        If isValid Then
            Dim name = TName.Text.Trim()
            Dim surname = TSurname.Text.Trim()

            If CheckTherapistExists(name, surname) Then
                isValid = False
                LblStatus.Text = $"თერაპევტი {name} {surname} უკვე არსებობს!"
                LblStatus.ForeColor = Color.Red
                BtnAdd.Visible = False
                Return
            End If
        End If

        ' დამატების ღილაკის ხილვადობის კონტროლი
        BtnAdd.Visible = isValid

        ' სტატუსის ტექსტის განახლება
        If isValid Then
            LblStatus.Text = "ფორმა ვალიდურია"
            LblStatus.ForeColor = Color.Green
        Else
            ' თუ სტატუსი უკვე არ არის დაყენებული (მსგავსი ჩანაწერის შემთხვევაში)
            If LblStatus.Text <> $"თერაპევტი {TName.Text.Trim()} {TSurname.Text.Trim()} უკვე არსებობს!" Then
                LblStatus.Text = "შეავსეთ ყველა საჭირო ველი"
                LblStatus.ForeColor = Color.Red
            End If
        End If
    End Sub

    ''' <summary>
    ''' BtnAdd ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Try
            ' ფორმის ხელახალი ვალიდაცია
            ValidateForm()

            If Not BtnAdd.Visible Then
                MessageBox.Show("გთხოვთ შეავსოთ ყველა სავალდებულო ველი", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' კიდევ ერთხელ შევამოწმოთ არსებობს თუ არა მსგავსი ჩანაწერი
            Dim name = TName.Text.Trim()
            Dim surname = TSurname.Text.Trim()

            If CheckTherapistExists(name, surname) Then
                MessageBox.Show($"თერაპევტი {name} {surname} უკვე არსებობს!", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                LblStatus.Text = $"თერაპევტი {name} {surname} უკვე არსებობს!"
                LblStatus.ForeColor = Color.Red
                BtnAdd.Visible = False
                Return
            End If

            ' მწკრივის მომზადება Google Sheets-ისთვის
            Dim rowData As New List(Of Object)()

            ' A სვეტი - ID
            rowData.Add(Integer.Parse(LN.Text))

            ' B სვეტი - სახელი
            rowData.Add(TName.Text.Trim())

            ' C სვეტი - გვარი
            rowData.Add(TSurname.Text.Trim())

            ' D სვეტი - პირადი ნომერი
            rowData.Add(TPN.Text.Trim())

            ' E სვეტი - დაბადების თარიღი
            rowData.Add(DTPPer.Value.ToString("dd.MM.yyyy"))

            ' F სვეტი - ტელეფონი
            rowData.Add(TTel.Text.Trim())

            ' G სვეტი - ელ-ფოსტა
            rowData.Add(TMail.Text.Trim())

            ' H სვეტი - სტატუსი
            rowData.Add(If(RBActive.Checked, "active", "passive"))

            ' დამატება Google Sheets-ში
            dataService.AppendData("DB-Personal!A:H", rowData)

            ' წარმატების ინდიკატორების დაყენება
            IsSuccess = True
            AddedName = TName.Text.Trim()
            AddedSurname = TSurname.Text.Trim()

            MessageBox.Show($"თერაპევტი {TName.Text} {TSurname.Text} წარმატებით დაემატა", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' დავხუროთ ფორმა წარმატებული ოპერაციის შემდეგ - უნდა დაბრუნდეს DialogResult.OK
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            Debug.WriteLine($"BtnAdd_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"თერაპევტის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' BtnClear ღილაკზე დაჭერის ივენთი
    ''' </summary>
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        Try
            ' ველების გასუფთავება
            TName.Text = ""
            TSurname.Text = ""
            TPN.Text = ""
            TTel.Text = ""
            TMail.Text = ""

            ' DateTimePicker-ის საწყისი თარიღის დაყენება 
            DTPPer.Value = DateTime.Today.AddYears(-30) ' ნაგულისხმევად 30 წლის

            ' აქტიური სტატუსის არჩევა ნაგულისხმევად
            RBActive.Checked = True

            ' ველების ფონების გასუფთავება
            TName.BackColor = SystemColors.Window
            TSurname.BackColor = SystemColors.Window
            TPN.BackColor = SystemColors.Window
            TTel.BackColor = SystemColors.Window
            TMail.BackColor = SystemColors.Window

            ' ფორმის ვალიდაცია
            ValidateForm()

            MessageBox.Show("ფორმა გასუფთავდა", "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Debug.WriteLine($"BtnClear_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"ფორმის გასუფთავების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class