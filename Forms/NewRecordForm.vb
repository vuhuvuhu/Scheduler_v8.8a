' NewRecordForm.vb - ახალი ფორმა ჩანაწერისთვის
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class NewRecordForm
    ' მონაცემთა სერვისის მითითება
    Private ReadOnly dataService As IDataService
    ' ჩანაწერის ტიპი (მაგ: "სესია", "დაბადების დღე", "დავალება")
    Private ReadOnly recordType As String
    Private WithEvents BtnOK As Button
    ' კონსტრუქტორი
    Public Sub New(dataService As IDataService, recordType As String)
        InitializeComponent()
        Me.dataService = dataService
        Me.recordType = recordType

        ' ფორმის ინიციალიზაცია ტიპის მიხედვით
        InitializeFormByType()
    End Sub

    ' ტიპის მიხედვით ფორმის შევსება და მორგება
    Private Sub InitializeFormByType()
        Select Case recordType.ToLower()
            Case "დაბადების დღე"
                ' დაბადების დღის ფორმის ელემენტები
                Me.Text = "ახალი დაბადების დღის დამატება"
                ' შესაბამისი ინიციალიზაცია...

            Case "სესია"
                ' სესიის ფორმის ელემენტები
                Me.Text = "ახალი სესიის დამატება"
                ' შესაბამისი ინიციალიზაცია...

            Case "დავალება"
                ' დავალების ფორმის ელემენტები
                Me.Text = "ახალი დავალების დამატება"
                ' შესაბამისი ინიციალიზაცია...

            Case Else
                MessageBox.Show("უცნობი ჩანაწერის ტიპი", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
        End Select
    End Sub

    ' OK ღილაკზე დაჭერა
    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
        If ValidateInputs() Then
            SaveRecord()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ' ინფუთების ვალიდაცია
    Private Function ValidateInputs() As Boolean
        ' ვალიდაციის ლოგიკა ჩანაწერის ტიპის მიხედვით
        ' ...
        Return True ' ან False თუ ვალიდაცია ვერ გაიარა
    End Function

    ' ახალი ჩანაწერის შენახვა
    Private Sub SaveRecord()
        Try
            Select Case recordType.ToLower()
                Case "დაბადების დღე"
                    SaveBirthday()
                Case "სესია"
                    SaveSession()
                Case "დავალება"
                    SaveTask()
                Case Else
                    Throw New Exception("უცნობი ჩანაწერის ტიპი")
            End Select

            MessageBox.Show("ჩანაწერი წარმატებით დაემატა", "წარმატება", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"შეცდომა ჩანაწერის შენახვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' დაბადების დღის შენახვა
    Private Sub SaveBirthday()
        ' მნიშვნელობების წაკითხვა ფორმიდან
        Dim firstName As String '= TxtFirstName.Text.Trim()
        Dim lastName As String '= TxtLastName.Text.Trim()
        Dim birthDate As DateTime '= DtpBirthDate.Value
        Dim email As String '= TxtEmail.Text.Trim()
        Dim phone As String '= TxtPhone.Text.Trim()

        ' ახალი მოდელის შექმნა
        Dim birthday As New BirthdayModel()
        birthday.PersonName = firstName
        birthday.PersonSurname = lastName
        birthday.BirthDate = birthDate
        birthday.Email = email
        birthday.Phone = phone

        ' შევქმნათ მონაცემების რიგი Google Sheets-ისთვის
        Dim rowData = birthday.ToSheetRow()

        ' დავამატოთ ახალი მწკრივი
        dataService.AppendData("DB-Personal!B:G", rowData)
    End Sub

    ' სესიის შენახვა
    Private Sub SaveSession()
        ' ანალოგიური ლოგიკა სესიისთვის
        ' ...
    End Sub

    ' დავალების შენახვა
    Private Sub SaveTask()
        ' ანალოგიური ლოგიკა დავალებისთვის
        ' ...
    End Sub
End Class