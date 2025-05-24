' ===========================================
' 📄 Forms/ColumnSelectionForm.vb
' -------------------------------------------
' ბეჭდვისთვის სვეტების მონიშვნის ფორმა
' ===========================================
Imports System.Windows.Forms

''' <summary>
''' სვეტების მონიშვნის ფორმა ბეჭდვისთვის
''' მომხმარებელს საშუალება აქვს აირჩიოს რომელი სვეტები გინდა დაიბეჭდოს
''' </summary>
Public Class ColumnSelectionForm
    Inherits Form

    ' UI კომპონენტები
    Private lblTitle As Label
    Private lblInstructions As Label
    Private clbColumns As CheckedListBox
    Private btnSelectAll As Button
    Private btnDeselectAll As Button
    Private btnOK As Button
    Private btnCancel As Button
    Private pnlButtons As Panel

    ' მონაცემები
    Private ReadOnly sourceDataGridView As DataGridView
    Private selectedColumns As List(Of DataGridViewColumn)

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    ''' <param name="dgv">წყარო DataGridView</param>
    Public Sub New(dgv As DataGridView)
        sourceDataGridView = dgv
        selectedColumns = New List(Of DataGridViewColumn)
        InitializeComponents()
        PopulateColumnsList()
    End Sub

    ''' <summary>
    ''' UI კომპონენტების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeComponents()
        Try
            ' ფორმის ძირითადი პარამეტრები
            Me.Text = "სვეტების მონიშვნა ბეჭდვისთვის"
            Me.Size = New Size(450, 500)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New Font("Sylfaen", 10)

            ' სათაური
            lblTitle = New Label()
            lblTitle.Text = "ბეჭდვისთვის სვეტების მონიშვნა"
            lblTitle.Font = New Font("Sylfaen", 12, FontStyle.Bold)
            lblTitle.Location = New Point(10, 10)
            lblTitle.Size = New Size(420, 25)
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            lblTitle.ForeColor = Color.DarkBlue

            ' ინსტრუქციები
            lblInstructions = New Label()
            lblInstructions.Text = "მონიშნეთ ის სვეტები, რომლებიც გსურთ დაიბეჭდოს:"
            lblInstructions.Location = New Point(15, 45)
            lblInstructions.Size = New Size(410, 20)
            lblInstructions.ForeColor = Color.Black

            ' სვეტების CheckedListBox
            clbColumns = New CheckedListBox()
            clbColumns.Location = New Point(15, 75)
            clbColumns.Size = New Size(410, 300)
            clbColumns.CheckOnClick = True
            clbColumns.Font = New Font("Sylfaen", 9)

            ' ღილაკების პანელი
            pnlButtons = New Panel()
            pnlButtons.Location = New Point(15, 385)
            pnlButtons.Size = New Size(410, 35)

            ' ყველას მონიშვნა
            btnSelectAll = New Button()
            btnSelectAll.Text = "ყველას მონიშვნა"
            btnSelectAll.Location = New Point(0, 5)
            btnSelectAll.Size = New Size(100, 25)
            btnSelectAll.Font = New Font("Sylfaen", 9)
            AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click

            ' ყველას განიშვნა
            btnDeselectAll = New Button()
            btnDeselectAll.Text = "ყველას განიშვნა"
            btnDeselectAll.Location = New Point(110, 5)
            btnDeselectAll.Size = New Size(100, 25)
            btnDeselectAll.Font = New Font("Sylfaen", 9)
            AddHandler btnDeselectAll.Click, AddressOf BtnDeselectAll_Click

            ' OK ღილაკი
            btnOK = New Button()
            btnOK.Text = "OK"
            btnOK.Location = New Point(250, 5)
            btnOK.Size = New Size(75, 25)
            btnOK.Font = New Font("Sylfaen", 9)
            btnOK.DialogResult = DialogResult.OK
            AddHandler btnOK.Click, AddressOf BtnOK_Click

            ' Cancel ღილაკი
            btnCancel = New Button()
            btnCancel.Text = "გაუქმება"
            btnCancel.Location = New Point(335, 5)
            btnCancel.Size = New Size(75, 25)
            btnCancel.Font = New Font("Sylfaen", 9)
            btnCancel.DialogResult = DialogResult.Cancel

            ' კომპონენტების დამატება ფორმაზე
            pnlButtons.Controls.AddRange({btnSelectAll, btnDeselectAll, btnOK, btnCancel})
            Me.Controls.AddRange({lblTitle, lblInstructions, clbColumns, pnlButtons})

            ' ნაგულისხმევი ღილაკები
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            Debug.WriteLine("ColumnSelectionForm: UI კომპონენტები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: InitializeComponents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სვეტების სიის შევსება DataGridView-დან
    ''' </summary>
    Private Sub PopulateColumnsList()
        Try
            Debug.WriteLine("ColumnSelectionForm: სვეტების სიის შევსება")

            If sourceDataGridView Is Nothing Then
                Debug.WriteLine("ColumnSelectionForm: sourceDataGridView არის Nothing")
                Return
            End If

            clbColumns.Items.Clear()

            ' ყველა ხილული სვეტის დამატება (რედაქტირების ღილაკის გარდა)
            For Each column As DataGridViewColumn In sourceDataGridView.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    ' სვეტის ინფორმაციის შენახვა Tag-ში
                    Dim item As New ColumnItem(column)
                    clbColumns.Items.Add(item, True) ' ნაგულისხმევად ყველა მონიშნული
                    Debug.WriteLine($"ColumnSelectionForm: დაემატა სვეტი '{column.HeaderText}'")
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: სულ დაემატა {clbColumns.Items.Count} სვეტი")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: PopulateColumnsList შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სვეტის მონიშვნა
    ''' </summary>
    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        Try
            For i As Integer = 0 To clbColumns.Items.Count - 1
                clbColumns.SetItemChecked(i, True)
            Next
            Debug.WriteLine("ColumnSelectionForm: ყველა სვეტი მონიშნულია")
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnSelectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სვეტის განიშვნა
    ''' </summary>
    Private Sub BtnDeselectAll_Click(sender As Object, e As EventArgs)
        Try
            For i As Integer = 0 To clbColumns.Items.Count - 1
                clbColumns.SetItemChecked(i, False)
            Next
            Debug.WriteLine("ColumnSelectionForm: ყველა სვეტი განიშნულია")
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnDeselectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' OK ღილაკზე დაჭერა - მონიშნული სვეტების შეგროვება
    ''' </summary>
    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        Try
            selectedColumns.Clear()

            ' მონიშნული სვეტების შეგროვება
            For i As Integer = 0 To clbColumns.Items.Count - 1
                If clbColumns.GetItemChecked(i) Then
                    Dim item As ColumnItem = CType(clbColumns.Items(i), ColumnItem)
                    selectedColumns.Add(item.Column)
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: მონიშნულია {selectedColumns.Count} სვეტი")

            ' შემოწმება - მინიმუმ ერთი სვეტი უნდა იყოს მონიშნული
            If selectedColumns.Count = 0 Then
                MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ბეჭდვისთვის",
                               "ინფორმაცია", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.None ' არ დახუროს ფორმა
                Return
            End If

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnOK_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"სვეტების მონიშვნის შეცდომა: {ex.Message}", "შეცდომა",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' მონიშნული სვეტების მიღება
    ''' </summary>
    ''' <returns>მონიშნული სვეტების სია</returns>
    Public Function GetSelectedColumns() As List(Of DataGridViewColumn)
        Return New List(Of DataGridViewColumn)(selectedColumns)
    End Function

    ''' <summary>
    ''' სვეტის ინფორმაციის შესანახი კლასი
    ''' </summary>
    Private Class ColumnItem
        Public Property Column As DataGridViewColumn

        Public Sub New(col As DataGridViewColumn)
            Column = col
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Column.HeaderText} ({Column.Name})"
        End Function
    End Class

End Class