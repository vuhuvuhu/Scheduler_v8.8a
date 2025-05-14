' ===========================================
' 📄 Forms/AddTerForm.vb
' -------------------------------------------
' თერაპიის ტიპის დამატების ფორმა - მონაცემები იწერება DB-Therapy ფურცელზე
' ===========================================
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class AddTerForm
    ' მონაცემთა სერვისი
    Private ReadOnly dataService As IDataService
    ' დამატების წარმატების ფლაგი
    Public Property IsSuccess As Boolean = False
    ' დამატებული თერაპიის ტიპი
    Public Property AddedTherapy As String = ""

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
        Me.Text = "თერაპიის ტიპის დამატება"
    End Sub

    ''' <summary>
    ''' ფორმის ჩატვირთვის ივენთი
    ''' </summary>
    Private Sub AddTerForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' ახალი ID-ის მიღება
        LN.Text = (GetMaxRecordId() + 1).ToString()

        ' ფორმის ვალიდაცია
        ValidateForm()
    End Sub

    ''' <summary>
    ''' მაქსიმალური ID-ის მოძიება DB-Therapy ფურცლიდან
    ''' </summary>
    Private Function GetMaxRecordId() As Integer
        Try
            ' A სვეტის წაკითხვა DB-Therapy ფურცლიდან
            Dim rows = dataService.GetData("DB-Therapy!A2:A")
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
    ''' ამოწმებს არსებობს თუ არა უკვე მოცემული სახელწოდების თერაპიის ტიპი
    ''' </summary>
    ''' <returns>True თუ არსებობს, False თუ არ არსებობს</returns>
    Private Function CheckTherapyExists(therapyName As String) As Boolean
        Try
            ' B სვეტის წაკითხვა DB-Therapy ფურცლიდან
            Dim rows = dataService.GetData("DB-Therapy!B2:B")

            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                For Each row In rows
                    If row.Count >= 1 Then
                        ' შევამოწმოთ ემთხვევა თუ არა თერაპიის ტიპი
                        Dim rowTherapy = row(0).ToString().Trim()

                        If String.Equals(rowTherapy, therapyName, StringComparison.OrdinalIgnoreCase) Then
                            Return True ' ნაპოვნია დამთხვევა
                        End If
                    End If
                Next
            End If

            Return False ' არ არის ნაპოვნი დამთხვევა
        Catch ex As Exception
            Debug.WriteLine($"CheckTherapyExists შეცდომა: {ex.Message}")
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
    ''' ფორმის ვალიდაცია - ამოწმებს თერაპიის ტიპის ველს
    ''' </summary>
    Private Sub ValidateForm()
        ' სავალდებულო ველების შემოწმება - თერაპიის ტიპი
        Dim isValid As Boolean = Not String.IsNullOrWhiteSpace(TName.Text)

        ' შევამოწმოთ მსგავსი ჩანაწერი აქვე, თუ თერაპიის ტიპი უკვე შევსებულია
        If isValid Then
            Dim therapyName = TName.Text.Trim()

            If CheckTherapyExists(therapyName) Then
                isValid = False
                LblStatus.Text = $"თერაპიის ტიპი '{therapyName}' უკვე არსებობს!"
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
            If LblStatus.Text <> $"თერაპიის ტიპი '{TName.Text.Trim()}' უკვე არსებობს!" Then
                LblStatus.Text = "შეავსეთ თერაპიის ტიპის ველი"
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
                MessageBox.Show("გთხოვთ შეავსოთ თერაპიის ტიპის ველი", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' კიდევ ერთხელ შევამოწმოთ არსებობს თუ არა მსგავსი ჩანაწერი
            Dim therapyName = TName.Text.Trim()

            If CheckTherapyExists(therapyName) Then
                MessageBox.Show($"თერაპიის ტიპი '{therapyName}' უკვე არსებობს!", "გაფრთხილება", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                LblStatus.Text = $"თერაპიის ტიპი '{therapyName}' უკვე არსებობს!"
                LblStatus.ForeColor = Color.Red
                BtnAdd.Visible = False
                Return
            End If

            ' მწკრივის მომზადება Google Sheets-ისთვის
            Dim rowData As New List(Of Object)()

            ' A სვეტი - ID
            rowData.Add(Integer.Parse(LN.Text))

            ' B სვეტი - თერაპიის ტიპი
            rowData.Add(therapyName)

            ' დამატება Google Sheets-ში
            dataService.AppendData("DB-Therapy!A:B", rowData)

            ' წარმატების ინდიკატორების დაყენება
            IsSuccess = True
            AddedTherapy = therapyName

            MessageBox.Show($"თერაპიის ტიპი '{therapyName}' წარმატებით დაემატა", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' დავხუროთ ფორმა წარმატებული ოპერაციის შემდეგ - უნდა დაბრუნდეს DialogResult.OK
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            Debug.WriteLine($"BtnAdd_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"თერაპიის ტიპის დამატების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class