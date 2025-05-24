' ===========================================
' 📄 Forms/ColumnSelectionForm.vb
' -------------------------------------------
' მარტივი სვეტების მონიშვნის ფორმა
' Design-ის პრობლემების თავიდან აცილებით
' ===========================================
Imports System.Windows.Forms

Public Class ColumnSelectionForm
    Inherits Form

    ' DataGridView წყარო
    Private ReadOnly sourceGrid As DataGridView

    ' UI კონტროლები
    Private checkListBox As CheckedListBox
    Private btnOK As Button
    Private btnCancel As Button
    Private btnSelectAll As Button
    Private btnDeselectAll As Button
    Private lblTitle As Label
    Private lblInfo As Label

    ' შედეგი
    Private selectedColumns As New List(Of DataGridViewColumn)

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    Public Sub New(dataGridView As DataGridView)
        sourceGrid = dataGridView
        CreateForm()
        LoadColumns()
    End Sub

    ''' <summary>
    ''' ფორმის შექმნა კოდით
    ''' </summary>
    Private Sub CreateForm()
        Try
            ' ფორმის ძირითადი პარამეტრები
            Me.Text = "სვეტების არჩევანი"
            Me.Size = New Size(400, 500)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False

            ' სათაური
            lblTitle = New Label()
            lblTitle.Text = "აირჩიეთ PDF-ში ჩასართველი სვეტები:"
            lblTitle.Font = New Font("Sylfaen", 12, FontStyle.Bold)
            lblTitle.Location = New Point(20, 20)
            lblTitle.Size = New Size(350, 25)
            Me.Controls.Add(lblTitle)

            ' ინფორმაცია
            lblInfo = New Label()
            lblInfo.Text = "მონიშნეთ სასურველი სვეტები და დააჭირეთ OK-ს"
            lblInfo.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            lblInfo.Location = New Point(20, 50)
            lblInfo.Size = New Size(350, 20)
            Me.Controls.Add(lblInfo)

            ' CheckedListBox
            checkListBox = New CheckedListBox()
            checkListBox.Location = New Point(20, 80)
            checkListBox.Size = New Size(350, 280)
            checkListBox.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            checkListBox.CheckOnClick = True
            Me.Controls.Add(checkListBox)

            ' ღილაკები
            btnSelectAll = New Button()
            btnSelectAll.Text = "ყველას მონიშვნა"
            btnSelectAll.Location = New Point(20, 380)
            btnSelectAll.Size = New Size(100, 30)
            AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
            Me.Controls.Add(btnSelectAll)

            btnDeselectAll = New Button()
            btnDeselectAll.Text = "ყველას განიშვნა"
            btnDeselectAll.Location = New Point(130, 380)
            btnDeselectAll.Size = New Size(100, 30)
            AddHandler btnDeselectAll.Click, AddressOf BtnDeselectAll_Click
            Me.Controls.Add(btnDeselectAll)

            btnOK = New Button()
            btnOK.Text = "OK"
            btnOK.Location = New Point(250, 380)
            btnOK.Size = New Size(60, 30)
            btnOK.DialogResult = DialogResult.OK
            AddHandler btnOK.Click, AddressOf BtnOK_Click
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "გაუქმება"
            btnCancel.Location = New Point(320, 380)
            btnCancel.Size = New Size(70, 30)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            ' Default ღილაკები
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            Debug.WriteLine("ColumnSelectionForm: ფორმა შეიქმნა")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: CreateForm შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სვეტების ჩატვირთვა
    ''' </summary>
    Private Sub LoadColumns()
        Try
            If sourceGrid Is Nothing Then Return

            checkListBox.Items.Clear()

            ' ხილული სვეტები (Edit-ის გარდა)
            For Each column As DataGridViewColumn In sourceGrid.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    checkListBox.Items.Add(column.HeaderText, True) ' ყველა მონიშნული
                End If
            Next

            UpdateInfo()
            Debug.WriteLine($"ColumnSelectionForm: ჩატვირთულია {checkListBox.Items.Count} სვეტი")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: LoadColumns შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველას მონიშვნა
    ''' </summary>
    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        Try
            For i As Integer = 0 To checkListBox.Items.Count - 1
                checkListBox.SetItemChecked(i, True)
            Next
            UpdateInfo()
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnSelectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ყველას განიშვნა
    ''' </summary>
    Private Sub BtnDeselectAll_Click(sender As Object, e As EventArgs)
        Try
            For i As Integer = 0 To checkListBox.Items.Count - 1
                checkListBox.SetItemChecked(i, False)
            Next
            UpdateInfo()
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnDeselectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' OK ღილაკი
    ''' </summary>
    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        Try
            Dim checkedCount As Integer = checkListBox.CheckedItems.Count

            If checkedCount = 0 Then
                MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი", "გაფრთხილება",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Debug.WriteLine($"ColumnSelectionForm: OK - მონიშნულია {checkedCount} სვეტი")
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: BtnOK_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ინფორმაციის განახლება
    ''' </summary>
    Private Sub UpdateInfo()
        Try
            Dim checkedCount As Integer = checkListBox.CheckedItems.Count
            lblInfo.Text = $"მონიშნული: {checkedCount} სვეტი"
            btnOK.Enabled = (checkedCount > 0)
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: UpdateInfo შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' მონიშნული სვეტების მიღება
    ''' </summary>
    Public Function GetSelectedColumns() As List(Of DataGridViewColumn)
        Try
            selectedColumns.Clear()

            If sourceGrid Is Nothing Then Return selectedColumns

            ' ხილული სვეტების სია
            Dim visibleColumns As New List(Of DataGridViewColumn)
            For Each column As DataGridViewColumn In sourceGrid.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    visibleColumns.Add(column)
                End If
            Next

            ' მონიშნული ინდექსებისთვის შესაბამისი სვეტები
            For i As Integer = 0 To checkListBox.Items.Count - 1
                If checkListBox.GetItemChecked(i) AndAlso i < visibleColumns.Count Then
                    selectedColumns.Add(visibleColumns(i))
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: დაბრუნდება {selectedColumns.Count} მონიშნული სვეტი")
            Return selectedColumns

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: GetSelectedColumns შეცდომა: {ex.Message}")
            Return New List(Of DataGridViewColumn)
        End Try
    End Function

    ''' <summary>
    ''' CheckedListBox-ის ცვლილება
    ''' </summary>
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        Try
            ' ItemCheck ივენთის მიბმა Load-ის შემდეგ
            AddHandler checkListBox.ItemCheck, AddressOf CheckListBox_ItemCheck
            UpdateInfo()
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: OnLoad შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' CheckListBox ცვლილების ივენთი
    ''' </summary>
    Private Sub CheckListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        Try
            ' ასინქრონული განახლება
            Me.BeginInvoke(New Action(AddressOf UpdateInfo))
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: CheckListBox_ItemCheck შეცდომა: {ex.Message}")
        End Try
    End Sub

End Class