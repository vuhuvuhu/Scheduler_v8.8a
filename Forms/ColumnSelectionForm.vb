' ===========================================
' 📄 Forms/ColumnSelectionForm.vb
' -------------------------------------------
' სვეტების მონიშვნის ფორმა ბეჭდვისა და PDF ექსპორტისთვის
' მომხმარებელს აძლევს საშუალებას აირჩიოს რომელი სვეტები უნდა
' ===========================================
Imports System.Windows.Forms
Imports System.Drawing

Public Class ColumnSelectionForm

    ' UI კომპონენტები
    Private WithEvents btnOK As Button
    Private WithEvents btnCancel As Button
    Private WithEvents btnSelectAll As Button
    Private WithEvents btnDeselectAll As Button
    Private checkListColumns As CheckedListBox
    Private lblTitle As Label
    Private lblInstructions As Label

    ' მონაცემები
    Private ReadOnly sourceDataGridView As DataGridView
    Private selectedColumns As List(Of DataGridViewColumn)

    ''' <summary>
    ''' კონსტრუქტორი
    ''' </summary>
    ''' <param name="dgv">DataGridView რომლის სვეტებიც უნდა მოვარჩიოთ</param>
    Public Sub New(dgv As DataGridView)
        sourceDataGridView = dgv
        selectedColumns = New List(Of DataGridViewColumn)()

        ' ჯერ Designer-ის InitializeComponent
        InitializeComponent()

        ' შემდეგ ჩვენი UI-ის დამატებითი ინიციალიზაცია
        InitializeCustomComponents()
        PopulateColumnList()
    End Sub

    ''' <summary>
    ''' დამატებითი UI კომპონენტების ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeCustomComponents()
        Try
            ' ფორმის ძირითადი პარამეტრები
            Me.Text = "სვეტების მონიშვნა ბეჭდვისთვის"
            Me.Size = New Size(450, 550)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New Font("Sylfaen", 10, FontStyle.Regular)
            Me.BackColor = Color.FromArgb(245, 245, 250)

            ' იკონის მცდელობა
            Try
                Me.Icon = SystemIcons.Application
            Catch
                ' იკონის გარეშე
            End Try

            ' სათაური
            lblTitle = New Label()
            lblTitle.Text = "📋 აირჩიეთ სვეტები ბეჭდვისთვის"
            lblTitle.Font = New Font("Sylfaen", 14, FontStyle.Bold)
            lblTitle.Location = New Point(20, 15)
            lblTitle.Size = New Size(400, 30)
            lblTitle.ForeColor = Color.DarkBlue
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            Me.Controls.Add(lblTitle)

            ' ინსტრუქციები
            lblInstructions = New Label()
            lblInstructions.Text = "მონიშნეთ ის სვეტები რომლებიც გსურთ ბეჭდვაში ან PDF-ში ნახოთ:"
            lblInstructions.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            lblInstructions.Location = New Point(20, 50)
            lblInstructions.Size = New Size(400, 40)
            lblInstructions.ForeColor = Color.DarkGray
            lblInstructions.AutoSize = False
            Me.Controls.Add(lblInstructions)

            ' CheckedListBox სვეტებისთვის
            checkListColumns = New CheckedListBox()
            checkListColumns.Location = New Point(20, 100)
            checkListColumns.Size = New Size(400, 320)
            checkListColumns.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            checkListColumns.CheckOnClick = True
            checkListColumns.BorderStyle = BorderStyle.FixedSingle
            checkListColumns.BackColor = Color.White
            Me.Controls.Add(checkListColumns)

            ' "ყველას მონიშვნა" ღილაკი
            btnSelectAll = New Button()
            btnSelectAll.Text = "✓ ყველას მონიშვნა"
            btnSelectAll.Location = New Point(20, 435)
            btnSelectAll.Size = New Size(130, 35)
            btnSelectAll.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            btnSelectAll.BackColor = Color.LightGreen
            btnSelectAll.FlatStyle = FlatStyle.Flat
            btnSelectAll.FlatAppearance.BorderColor = Color.Green
            btnSelectAll.Cursor = Cursors.Hand
            Me.Controls.Add(btnSelectAll)

            ' "ყველას განიშვნა" ღილაკი
            btnDeselectAll = New Button()
            btnDeselectAll.Text = "✗ ყველას განიშვნა"
            btnDeselectAll.Location = New Point(160, 435)
            btnDeselectAll.Size = New Size(130, 35)
            btnDeselectAll.Font = New Font("Sylfaen", 9, FontStyle.Regular)
            btnDeselectAll.BackColor = Color.LightCoral
            btnDeselectAll.FlatStyle = FlatStyle.Flat
            btnDeselectAll.FlatAppearance.BorderColor = Color.Red
            btnDeselectAll.Cursor = Cursors.Hand
            Me.Controls.Add(btnDeselectAll)

            ' OK ღილაკი
            btnOK = New Button()
            btnOK.Text = "✓ OK"
            btnOK.Location = New Point(240, 480)
            btnOK.Size = New Size(90, 35)
            btnOK.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            btnOK.BackColor = Color.LimeGreen
            btnOK.ForeColor = Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.FlatAppearance.BorderColor = Color.Green
            btnOK.Cursor = Cursors.Hand
            Me.Controls.Add(btnOK)

            ' Cancel ღილაკი
            btnCancel = New Button()
            btnCancel.Text = "✗ გაუქმება"
            btnCancel.Location = New Point(340, 480)
            btnCancel.Size = New Size(90, 35)
            btnCancel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
            btnCancel.BackColor = Color.Crimson
            btnCancel.ForeColor = Color.White
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.FlatAppearance.BorderColor = Color.DarkRed
            btnCancel.Cursor = Cursors.Hand
            Me.Controls.Add(btnCancel)

            ' Dialog Result-ები
            btnOK.DialogResult = DialogResult.OK
            btnCancel.DialogResult = DialogResult.Cancel
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            Debug.WriteLine("ColumnSelectionForm: დამატებითი UI კომპონენტები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: InitializeCustomComponents შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სვეტების სიის შევსება CheckedListBox-ში
    ''' </summary>
    Private Sub PopulateColumnList()
        Try
            Debug.WriteLine("ColumnSelectionForm: სვეტების სიის შევსება")

            If sourceDataGridView Is Nothing Then
                Debug.WriteLine("ColumnSelectionForm: sourceDataGridView არის Nothing")
                Return
            End If

            checkListColumns.Items.Clear()
            columnIndexDictionary.Clear() ' 🔧 Dictionary-ის გასუფთავება

            Dim itemIndex As Integer = 0

            ' 🔧 ჯერ ყველა ხილული სვეტის სახელი ვნახოთ
            Debug.WriteLine("ColumnSelectionForm: ხილული სვეტების სია:")
            For Each column As DataGridViewColumn In sourceDataGridView.Columns
                If column.Visible Then
                    Debug.WriteLine($"  - {column.Name} (HeaderText: '{column.HeaderText}')")
                End If
            Next

            ' ყველა ხილული სვეტის დამატება (Edit ღილაკის გარდა)
            For Each column As DataGridViewColumn In sourceDataGridView.Columns
                If column.Visible AndAlso column.Name <> "Edit" Then
                    ' CheckedListBox-ში ვამატებთ სვეტის სათაურს
                    checkListColumns.Items.Add(column.HeaderText)

                    ' 🔧 Dictionary-ში ვინახავთ ინდექსი -> column დაკავშირებას
                    columnIndexDictionary(itemIndex) = column

                    ' ნაგულისხმევად მონიშნული (რჩევები: ყველაზე მნიშვნელოვანი სვეტები)
                    Dim shouldBeChecked As Boolean = ShouldColumnBeCheckedByDefault(column.Name)
                    checkListColumns.SetItemChecked(itemIndex, shouldBeChecked)

                    Debug.WriteLine($"ColumnSelectionForm: დაემატა სვეტი #{itemIndex} '{column.HeaderText}' (Name: {column.Name}), მონიშნული: {shouldBeChecked}")

                    itemIndex += 1 ' 🔧 ინდექსის ზრდა ხელით
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: სულ დაემატა {checkListColumns.Items.Count} სვეტი, Dictionary: {columnIndexDictionary.Count}")

            ' 🔧 დებაგირებისთვის - ვნახოთ რამდენი სვეტია მონიშნული
            Dim checkedCount As Integer = 0
            For i As Integer = 0 To checkListColumns.Items.Count - 1
                If checkListColumns.GetItemChecked(i) Then
                    checkedCount += 1
                End If
            Next
            Debug.WriteLine($"ColumnSelectionForm: ნაგულისხმევად მონიშნულია {checkedCount} სვეტი")

            ' 🔧 თუ არცერთი სვეტი არ არის მონიშნული, პირველი 3 მონიშნულია
            If checkedCount = 0 Then
                Debug.WriteLine("ColumnSelectionForm: არცერთი სვეტი არ არის მონიშნული, ვმონიშნავთ პირველ 3-ს...")
                Dim maxToCheck = Math.Min(3, checkListColumns.Items.Count)
                For i As Integer = 0 To maxToCheck - 1
                    checkListColumns.SetItemChecked(i, True)
                    Debug.WriteLine($"ColumnSelectionForm: ავტომატურად მონიშნული - #{i} '{checkListColumns.Items(i)}'")
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: PopulateColumnList შეცდომა: {ex.Message}")
        End Try
    End Sub

    ' Dictionary რაც დაგვეხმარება column ობიექტების შენახვაში
    Private columnIndexDictionary As New Dictionary(Of Integer, DataGridViewColumn)()

    ''' <summary>
    ''' განსაზღვრავს უნდა იყოს თუ არა სვეტი ნაგულისხმევად მონიშნული
    ''' </summary>
    ''' <param name="columnName">სვეტის სახელი</param>
    ''' <returns>True თუ უნდა იყოს მონიშნული</returns>
    Private Function ShouldColumnBeCheckedByDefault(columnName As String) As Boolean
        Try
            Debug.WriteLine($"ColumnSelectionForm: ShouldColumnBeCheckedByDefault ამოწმებს სვეტს '{columnName}'")

            ' 🔧 ყველაზე მნიშვნელოვანი სვეტები - დავამატოთ მეტი ვარიანტი
            Dim importantColumns As String() = {
                "N",             ' სესიის ID
                "BeneName",      ' ბენეფიციარის სახელი  
                "BeneSurname",   ' ბენეფიციარის გვარი
                "DateTime",      ' თარიღი და დრო
                "Duration",      ' ხანგრძლივობა
                "Therapist",     ' თერაპევტი
                "TherapyType",   ' თერაპიის ტიპი
                "Status",        ' სტატუსი
                "Price",         ' ფასი
                "Space",         ' სივრცე
                "Funding",       ' დაფინანსება
                "IsGroup",       ' ჯგუფური
                "Author",        ' ავტორი
                "Comments"       ' კომენტარი
            }

            ' შევამოწმოთ ემთხვევა თუ არა
            Dim shouldCheck = importantColumns.Contains(columnName, StringComparer.OrdinalIgnoreCase)
            Debug.WriteLine($"ColumnSelectionForm: სვეტი '{columnName}' - მონიშვნა: {shouldCheck}")

            Return shouldCheck

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: ShouldColumnBeCheckedByDefault შეცდომა: {ex.Message}")
            Return True ' 🔧 შეცდომის შემთხვევაში ყველა მონიშნული
        End Try
    End Function

    ''' <summary>
    ''' მონიშნული სვეტების მიღება
    ''' </summary>
    ''' <returns>მონიშნული სვეტების სია</returns>
    Public Function GetSelectedColumns() As List(Of DataGridViewColumn)
        Try
            Debug.WriteLine("ColumnSelectionForm: მონიშნული სვეტების მიღება")

            ' 🔧 თუ უკვე შენახულია selectedColumns (ფორმის დახურვის შემდეგ), გამოვიყენოთ ის
            If selectedColumns IsNot Nothing AndAlso selectedColumns.Count > 0 Then
                Debug.WriteLine($"ColumnSelectionForm: გამოყენებულია შენახული selectedColumns - {selectedColumns.Count} სვეტი")
                Return selectedColumns
            End If

            ' 🔧 წინააღმდეგ შემთხვევაში, ხელახლა გამოვთვალოთ
            Dim tempSelectedColumns As New List(Of DataGridViewColumn)()

            ' 🔧 ვალიდაცია - Dictionary და CheckedListBox-ის სინქრონიზაცია
            Debug.WriteLine($"ColumnSelectionForm: CheckedListBox items: {checkListColumns.Items.Count}, Dictionary items: {columnIndexDictionary.Count}")

            For i As Integer = 0 To checkListColumns.Items.Count - 1
                Dim isChecked = checkListColumns.GetItemChecked(i)
                Debug.WriteLine($"ColumnSelectionForm: Item #{i} '{checkListColumns.Items(i)}' - მონიშნული: {isChecked}")

                If isChecked Then
                    If columnIndexDictionary.ContainsKey(i) Then
                        Dim column = columnIndexDictionary(i)
                        tempSelectedColumns.Add(column)
                        Debug.WriteLine($"ColumnSelectionForm: ✓ მონიშნული სვეტი დაემატა - {column.HeaderText} (Name: {column.Name})")
                    Else
                        Debug.WriteLine($"ColumnSelectionForm: ❌ Dictionary-ში ვერ მოიძებნა ინდექსი {i}")
                    End If
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: საბოლოო შედეგი - მონიშნულია {tempSelectedColumns.Count} სვეტი სულ")

            ' 🔧 დაბრუნებამდე ბოლო შემოწმება
            If tempSelectedColumns.Count = 0 Then
                Debug.WriteLine("ColumnSelectionForm: ❌ საბოლოო შედეგი ცარიელია!")
            End If

            Return tempSelectedColumns

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: GetSelectedColumns შეცდომა: {ex.Message}")
            Debug.WriteLine($"ColumnSelectionForm: Exception StackTrace: {ex.StackTrace}")
            Return New List(Of DataGridViewColumn)()
        End Try
    End Function

    ''' <summary>
    ''' "ყველას მონიშვნა" ღილაკის ივენთი
    ''' </summary>
    Private Sub btnSelectAll_Click(sender As Object, e As EventArgs) Handles btnSelectAll.Click
        Try
            Debug.WriteLine("ColumnSelectionForm: ყველას მონიშვნა")

            For i As Integer = 0 To checkListColumns.Items.Count - 1
                checkListColumns.SetItemChecked(i, True)
            Next

            Debug.WriteLine($"ColumnSelectionForm: მონიშნულია {checkListColumns.Items.Count} სვეტი")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: btnSelectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' "ყველას განიშვნა" ღილაკის ივენთი
    ''' </summary>
    Private Sub btnDeselectAll_Click(sender As Object, e As EventArgs) Handles btnDeselectAll.Click
        Try
            Debug.WriteLine("ColumnSelectionForm: ყველას განიშვნა")

            For i As Integer = 0 To checkListColumns.Items.Count - 1
                checkListColumns.SetItemChecked(i, False)
            Next

            Debug.WriteLine("ColumnSelectionForm: ყველა სვეტი განიშნულია")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: btnDeselectAll_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' OK ღილაკის ივენთი - გაუმჯობესებული ვალიდაციით
    ''' Dictionary-ის შენარჩუნებით
    ''' </summary>
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Try
            Debug.WriteLine("ColumnSelectionForm: OK ღილაკზე დაჭერა")

            ' 🔧 რეალური დროის შემოწმება - GetSelectedColumns-ს გამოძახება
            Dim tempSelectedColumns = GetSelectedColumns()

            If tempSelectedColumns.Count = 0 Then
                MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ბეჭდვისთვის!" & Environment.NewLine & Environment.NewLine &
                               "დააჭირეთ სვეტის სახელს ან გამოიყენეთ 'ყველას მონიშვნა' ღილაკი.",
                               "სვეტები არ არის მონიშნული!",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' 🔧 ფოკუსი CheckedListBox-ზე
                checkListColumns.Focus()
                Return ' არ დავხუროთ ფორმა
            End If

            ' 🔧 ᲛᲜᲘᲨᲕᲜᲔᲚᲝᲕᲐᲜᲘ: selectedColumns-ს შენახვა დახურვამდე
            selectedColumns.Clear()
            selectedColumns.AddRange(tempSelectedColumns)

            Debug.WriteLine($"ColumnSelectionForm: ვალიდაცია გავლილია - მონიშნულია {tempSelectedColumns.Count} სვეტი, ფორმა იხურება")
            Debug.WriteLine($"ColumnSelectionForm: selectedColumns-ში შენახულია {selectedColumns.Count} სვეტი")

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: btnOK_Click შეცდომა: {ex.Message}")
            MessageBox.Show($"შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Cancel ღილაკის ივენთი
    ''' </summary>
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Try
            Debug.WriteLine("ColumnSelectionForm: Cancel ღილაკზე დაჭერა")
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: btnCancel_Click შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფორმის ჩატვირთვისას - გაუმჯობესებული ვერსია
    ''' </summary>
    Protected Overrides Sub OnLoad(e As EventArgs)
        Try
            MyBase.OnLoad(e)

            ' ფორმის ცენტრირება
            If Me.Owner IsNot Nothing Then
                Me.Location = New Point(
                    Me.Owner.Location.X + (Me.Owner.Width - Me.Width) \ 2,
                    Me.Owner.Location.Y + (Me.Owner.Height - Me.Height) \ 2
                )
            End If

            ' ფოკუსი CheckedListBox-ზე
            checkListColumns.Focus()

            ' 🔧 დებაგირებისთვის - იტემებისა და მონიშნულებების რაოდენობა
            Dim checkedCount As Integer = 0
            For i As Integer = 0 To checkListColumns.Items.Count - 1
                If checkListColumns.GetItemChecked(i) Then
                    checkedCount += 1
                End If
            Next

            Debug.WriteLine($"ColumnSelectionForm: ფორმა ჩაიტვირთა - Items: {checkListColumns.Items.Count}, Dictionary: {columnIndexDictionary.Count}, Checked: {checkedCount}")

            ' 🔧 გარანტირებული მონიშვნა - თუ არცერთი არ არის მონიშნული
            If checkedCount = 0 Then
                Debug.WriteLine("ColumnSelectionForm: არცერთი სვეტი არ არის მონიშნული, ვმონიშნავთ ყველას...")

                ' ყველა სვეტის მონიშვნა (გარანტირებული მეთოდი)
                For i As Integer = 0 To checkListColumns.Items.Count - 1
                    checkListColumns.SetItemChecked(i, True)
                    Debug.WriteLine($"ColumnSelectionForm: ავტომატურად მონიშნული - {checkListColumns.Items(i)}")
                Next

                ' ახლა კიდევ ერთხელ ვითვლით
                checkedCount = checkListColumns.Items.Count
                Debug.WriteLine($"ColumnSelectionForm: ყველა სვეტი მონიშნულია - სულ: {checkedCount}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: OnLoad შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ფორმის დახურვისას - რესურსების გათავისუფლება
    ''' 🔧 Dictionary და selectedColumns არ გავასუფთავოთ, რადგან UC_Schedule-მა ჯერ უნდა წაიკითხოს
    ''' </summary>
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            MyBase.OnFormClosed(e)

            ' 🔧 მხოლოდ UI რესურსები დავათავისუფლოთ
            ' Dictionary და selectedColumns კი დავტოვოთ UC_Schedule-სთვის
            Debug.WriteLine($"ColumnSelectionForm: რესურსები განთავისუფლებულია - selectedColumns: {selectedColumns?.Count}, Dictionary: {columnIndexDictionary?.Count}")

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: OnFormClosed შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კლავიშის დაჭერის ივენთი - Enter და Escape კლავიშების დასამუშავებლად
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Try
            Select Case keyData
                Case Keys.Enter
                    ' Enter კლავიშზე დაჭერისას OK-ის შესრულება
                    btnOK_Click(Nothing, Nothing)
                    Return True

                Case Keys.Escape
                    ' Escape კლავიშზე დაჭერისას Cancel-ის შესრულება
                    btnCancel_Click(Nothing, Nothing)
                    Return True

                Case Keys.A Or Keys.Control
                    ' Ctrl+A - ყველას მონიშვნა
                    btnSelectAll_Click(Nothing, Nothing)
                    Return True

                Case Keys.D Or Keys.Control
                    ' Ctrl+D - ყველას განიშვნა
                    btnDeselectAll_Click(Nothing, Nothing)
                    Return True

            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)

        Catch ex As Exception
            Debug.WriteLine($"ColumnSelectionForm: ProcessCmdKey შეცდომა: {ex.Message}")
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Try
    End Function

    ''' <summary>
    ''' სვეტების რაოდენობის მიღება
    ''' </summary>
    Public ReadOnly Property ColumnCount As Integer
        Get
            Try
                Return checkListColumns?.Items.Count
            Catch
                Return 0
            End Try
        End Get
    End Property

    ''' <summary>
    ''' მონიშნული სვეტების რაოდენობის მიღება
    ''' </summary>
    Public ReadOnly Property SelectedColumnCount As Integer
        Get
            Try
                Dim count As Integer = 0
                For i As Integer = 0 To checkListColumns.Items.Count - 1
                    If checkListColumns.GetItemChecked(i) Then
                        count += 1
                    End If
                Next
                Return count
            Catch
                Return 0
            End Try
        End Get
    End Property

End Class