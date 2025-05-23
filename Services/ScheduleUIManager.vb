' ===========================================
' 📄 Services/ScheduleUIManager.vb
' -------------------------------------------
' განრიგის UI-ის მართვის სერვისი
' UC_Schedule-დან გატანილი UI ლოგიკა
' ===========================================
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Globalization
Imports Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის მომხმარებლის ინტერფეისის მართვის სერვისი
    ''' პასუხისმგებელია DataGridView-ის, ფილტრების და ნავიგაციის მართვაზე
    ''' </summary>
    Public Class ScheduleUIManager

        ' UI კონტროლები
        Private ReadOnly dgvSchedule As DataGridView
        Private ReadOnly lblPage As Label
        Private ReadOnly btnPrev As Button
        Private ReadOnly btnNext As Button

        ' ღირებულები ფერებისთვის
        Private ReadOnly headerColor As Color = Color.FromArgb(60, 80, 150)
        Private ReadOnly alternateRowColor As Color = Color.FromArgb(245, 245, 250)
        Private ReadOnly selectionColor As Color = Color.FromArgb(180, 200, 255)

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dataGridView">მთავარი DataGridView</param>
        ''' <param name="pageLabel">გვერდის ლეიბლი</param>
        ''' <param name="prevButton">წინა გვერდის ღილაკი</param>
        ''' <param name="nextButton">შემდეგი გვერდის ღილაკი</param>
        Public Sub New(dataGridView As DataGridView,
                      pageLabel As Label,
                      prevButton As Button,
                      nextButton As Button)
            dgvSchedule = dataGridView
            lblPage = pageLabel
            btnPrev = prevButton
            btnNext = nextButton
        End Sub

        ''' <summary>
        ''' DataGridView-ის საწყისი კონფიგურაცია
        ''' </summary>
        Public Sub ConfigureDataGridView()
            Try
                If dgvSchedule Is Nothing Then Return

                Debug.WriteLine("ScheduleUIManager: DataGridView-ის კონფიგურაცია")

                With dgvSchedule
                    .AutoGenerateColumns = False
                    .AllowUserToAddRows = False
                    .AllowUserToDeleteRows = False
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .MultiSelect = False

                    ' თავის სტილი
                    .EnableHeadersVisualStyles = False
                    .ColumnHeadersDefaultCellStyle.BackColor = headerColor
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Sylfaen", 10, FontStyle.Bold)

                    ' მწკრივების სტილი
                    .RowsDefaultCellStyle.BackColor = Color.White
                    .AlternatingRowsDefaultCellStyle.BackColor = alternateRowColor
                    .DefaultCellStyle.SelectionBackColor = selectionColor
                    .DefaultCellStyle.SelectionForeColor = Color.Black
                End With

                Debug.WriteLine("ScheduleUIManager: DataGridView კონფიგურაცია დასრულდა")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: DataGridView კონფიგურაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' DataGridView სვეტების შექმნა მომხმარებლის როლის მიხედვით
        ''' </summary>
        ''' <param name="userRole">მომხმარებლის როლი (1=ადმინი, 2=მენეჯერი, 6=მომხმარებელი)</param>
        Public Sub CreateColumns(userRole As Integer)
            Try
                Debug.WriteLine($"ScheduleUIManager: სვეტების შექმნა, მომხმარებლის როლი={userRole}")

                dgvSchedule.Columns.Clear()

                With dgvSchedule.Columns
                    ' ძირითადი სვეტები ყველა მომხმარებლისთვის
                    .Add("N", "N")
                    .Add("Tarigi", "თარიღი")
                    .Add("Duri", "ხანგძლივობა")
                    .Add("Bene", "ბენეფიციარი")
                    .Add("Per", "თერაპევტი")
                    .Add("Ter", "თერაპია")
                    .Add("Group", "ჯგუფური")
                    .Add("Space", "სივრცე")
                    .Add("Price", "თანხა")
                    .Add("Status", "სტატუსი")
                    .Add("Program", "პროგრამა")
                    .Add("Coment", "კომენტარი")

                    ' ადმინისა და მენეჯერისთვის დამატებითი სვეტები
                    If userRole = 1 OrElse userRole = 2 Then
                        .Add("EditDate", "რედ. თარიღი")
                        .Add("Author", "ავტორი")
                        Debug.WriteLine("ScheduleUIManager: დაემატა ადმინის/მენეჯერის სვეტები")
                    End If

                    ' რედაქტირების ღილაკი
                    Dim editBtn As New DataGridViewButtonColumn()
                    editBtn.Name = "Edit"
                    editBtn.HeaderText = ""
                    editBtn.Text = "✎"
                    editBtn.UseColumnTextForButtonValue = True
                    .Add(editBtn)
                End With

                ' სვეტების ზომების დაყენება
                SetColumnWidths(userRole)

                Debug.WriteLine($"ScheduleUIManager: შეიქმნა {dgvSchedule.Columns.Count} სვეტი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: სვეტების შექმნის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სვეტების სიგანეების დაყენება
        ''' </summary>
        ''' <param name="userRole">მომხმარებლის როლი</param>
        Private Sub SetColumnWidths(userRole As Integer)
            Try
                With dgvSchedule.Columns
                    .Item("N").Width = 40
                    .Item("Tarigi").Width = 110
                    .Item("Duri").Width = 40
                    .Item("Bene").Width = 180
                    .Item("Per").Width = 185
                    .Item("Ter").Width = 185
                    .Item("Group").Width = 50
                    .Item("Space").Width = 80
                    .Item("Price").Width = 60
                    .Item("Status").Width = 130
                    .Item("Program").Width = 130
                    .Item("Coment").Width = 120
                    .Item("Coment").DefaultCellStyle.WrapMode = DataGridViewTriState.False
                    .Item("Coment").ToolTipText = "სრული კომენტარი გამოჩნდება მაუსის მიტანისას"

                    ' ადმინისა და მენეჯერისთვის დამატებითი სვეტები
                    If userRole = 1 OrElse userRole = 2 Then
                        If .Contains("EditDate") Then .Item("EditDate").Width = 110
                        If .Contains("Author") Then .Item("Author").Width = 120
                    End If

                    .Item("Edit").Width = 24
                End With

                Debug.WriteLine("ScheduleUIManager: სვეტების სიგანეები დაყენებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: სვეტების სიგანეების დაყენების შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონაცემების ჩატვირთვა DataGridView-ში
        ''' </summary>
        ''' <param name="data">ჩასატვირთი მონაცემები</param>
        ''' <param name="userRole">მომხმარებლის როლი</param>
        Public Sub LoadDataToGrid(data As List(Of IList(Of Object)), userRole As Integer)
            Try
                Debug.WriteLine($"ScheduleUIManager: მონაცემების ჩატვირთვა - {data.Count} ჩანაწერი")

                ' ცხრილის გასუფთავება
                dgvSchedule.Rows.Clear()

                ' თუ სვეტები არ არსებობს, შევქმნათ
                If dgvSchedule.Columns.Count = 0 Then
                    CreateColumns(userRole)
                End If

                ' ყოველი ჩანაწერის დამატება
                For i As Integer = 0 To data.Count - 1
                    Try
                        AddRowToGrid(data(i), userRole, i)
                    Catch ex As Exception
                        Debug.WriteLine($"ScheduleUIManager: მწკრივი {i}-ის დამატების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next

                Debug.WriteLine($"ScheduleUIManager: ჩატვირთულია {dgvSchedule.Rows.Count} მწკრივი")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: LoadDataToGrid შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ერთი მწკრივის დამატება DataGridView-ში
        ''' </summary>
        ''' <param name="rowData">მწკრივის მონაცემები</param>
        ''' <param name="userRole">მომხმარებლის როლი</param>
        ''' <param name="rowIndex">მწკრივის ინდექსი ფერის დასაყენებლად</param>
        Private Sub AddRowToGrid(rowData As IList(Of Object), userRole As Integer, rowIndex As Integer)
            Try
                ' ბენეფიციარის სრული სახელი
                Dim beneficiaryFullName As String = $"{rowData(3).ToString().Trim()} {rowData(4).ToString().Trim()}"

                ' ძირითადი მონაცემების მომზადება
                Dim dgvRowData As New List(Of Object) From {
                    rowData(0),                                                ' N (ID)
                    rowData(5),                                                ' თარიღი
                    If(rowData.Count > 6, rowData(6).ToString(), "60"),       ' ხანგძლივობა
                    beneficiaryFullName,                                       ' ბენეფიციარი
                    If(rowData.Count > 8, rowData(8).ToString(), ""),         ' თერაპევტი
                    If(rowData.Count > 9, rowData(9).ToString(), ""),         ' თერაპია
                    If(rowData.Count > 7, rowData(7).ToString(), ""),         ' ჯგუფური
                    If(rowData.Count > 10, rowData(10).ToString(), ""),       ' სივრცე
                    If(rowData.Count > 11, rowData(11).ToString(), "0"),      ' ფასი
                    If(rowData.Count > 12, rowData(12).ToString(), "დაგეგმილი"), ' სტატუსი
                    If(rowData.Count > 13, rowData(13).ToString(), ""),       ' პროგრამა
                    If(rowData.Count > 14, rowData(14).ToString(), "")        ' კომენტარი
                }

                ' ადმინისა და მენეჯერისთვის დამატებითი მონაცემები
                If userRole = 1 OrElse userRole = 2 Then
                    Dim editDate = If(rowData.Count > 1, rowData(1).ToString(), "")
                    Dim author = If(rowData.Count > 2, rowData(2).ToString(), "")
                    dgvRowData.Add(editDate)  ' რედაქტირების თარიღი
                    dgvRowData.Add(author)    ' ავტორი
                End If

                ' რედაქტირების ღილაკი
                dgvRowData.Add("✎")

                ' მწკრივის დამატება DataGridView-ში
                Dim addedRowIndex As Integer = dgvSchedule.Rows.Add(dgvRowData.ToArray())

                ' მწკრივის ფერის დაყენება სტატუსის მიხედვით
                ApplyRowStatusColor(addedRowIndex, rowData)

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: AddRowToGrid შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' მწკრივის ფერის დაყენება სტატუსის მიხედვით
        ''' </summary>
        ''' <param name="rowIndex">მწკრივის ინდექსი</param>
        ''' <param name="rowData">მწკრივის მონაცემები</param>
        Private Sub ApplyRowStatusColor(rowIndex As Integer, rowData As IList(Of Object))
            Try
                If rowIndex < 0 OrElse rowIndex >= dgvSchedule.Rows.Count Then Return

                ' სტატუსის მიღება
                Dim statusText As String = If(rowData.Count > 12, rowData(12).ToString().Trim(), "დაგეგმილი")
                Dim currentTime As DateTime = DateTime.Now
                Dim sessionTime As DateTime

                ' სესიის თარიღის პარსინგი
                Dim dateFormats As String() = {"dd.MM.yyyy HH:mm", "dd.MM.yyyy", "dd.MM.yy HH:mm", "d.M.yyyy HH:mm", "d/M/yyyy H:mm:ss"}
                If Not DateTime.TryParseExact(rowData(5).ToString().Trim(), dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, sessionTime) Then
                    DateTime.TryParse(rowData(5).ToString().Trim(), sessionTime)
                End If

                ' სტატუსის მიხედვით ფერის დაყენება
                Dim statusColor As Color = GetStatusColor(statusText, sessionTime, currentTime)
                dgvSchedule.Rows(rowIndex).DefaultCellStyle.BackColor = statusColor

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: ApplyRowStatusColor შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' სტატუსის მიხედვით ფერის დაბრუნება
        ''' </summary>
        ''' <param name="statusText">სტატუსის ტექსტი</param>
        ''' <param name="sessionTime">სესიის დრო</param>
        ''' <param name="currentTime">მიმდინარე დრო</param>
        ''' <returns>შესაბამისი ფერი</returns>
        Private Function GetStatusColor(statusText As String, sessionTime As DateTime, currentTime As DateTime) As Color
            Try
                Select Case statusText.ToLower().Trim()
                    Case "შესრულებული"
                        Return Color.LightGreen
                    Case "აღდგენა"
                        Return Color.Honeydew
                    Case "გაცდენა არასაპატიო"
                        Return Color.Plum
                    Case "გაცდენა საპატიო"
                        Return Color.LightYellow
                    Case "პროგრამით გატარება"
                        Return Color.LightGray
                    Case "გაუქმება", "გაუქმებული"
                        Return Color.LightCoral
                    Case "დაგეგმილი"
                        ' დაგეგმილი სესიებისთვის - თუ დრო გასულია, ღია წითელი, თუ არა - ღია ლურჯი
                        If sessionTime < currentTime Then
                            Return Color.LightCoral ' ღია წითელი - ვადაგადაცილებული
                        Else
                            Return Color.LightBlue ' ღია ლურჯი - დაგეგმილი
                        End If
                    Case Else
                        ' უცნობი სტატუსებისთვის სტანდარტული ფერი
                        Return Color.White
                End Select
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: GetStatusColor შეცდომა: {ex.Message}")
                Return Color.White
            End Try
        End Function

        ''' <summary>
        ''' გვერდის ლეიბლის განახლება
        ''' </summary>
        ''' <param name="currentPage">მიმდინარე გვერდი</param>
        ''' <param name="totalPages">სულ გვერდები</param>
        Public Sub UpdatePageLabel(currentPage As Integer, totalPages As Integer)
            Try
                If lblPage IsNot Nothing Then
                    lblPage.Text = $"{currentPage} / {totalPages}"
                End If
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: UpdatePageLabel შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ნავიგაციის ღილაკების მდგომარეობის განახლება
        ''' </summary>
        ''' <param name="currentPage">მიმდინარე გვერდი</param>
        ''' <param name="totalPages">სულ გვერდები</param>
        Public Sub UpdateNavigationButtons(currentPage As Integer, totalPages As Integer)
            Try
                ' წინა ღილაკი
                If btnPrev IsNot Nothing Then
                    btnPrev.Enabled = (currentPage > 1)
                End If

                ' შემდეგი ღილაკი
                If btnNext IsNot Nothing Then
                    btnNext.Enabled = (currentPage < totalPages)
                End If
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: UpdateNavigationButtons შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ნავიგაციის ღილაკების სტილის დაყენება
        ''' </summary>
        Public Sub ConfigureNavigationButtons()
            Try
                ' წინა ღილაკის სტილი
                If btnPrev IsNot Nothing Then
                    btnPrev.FlatStyle = FlatStyle.Flat
                    btnPrev.BackColor = headerColor
                    btnPrev.ForeColor = Color.White
                    btnPrev.Text = ""
                End If

                ' შემდეგი ღილაკის სტილი
                If btnNext IsNot Nothing Then
                    btnNext.FlatStyle = FlatStyle.Flat
                    btnNext.BackColor = headerColor
                    btnNext.ForeColor = Color.White
                    btnNext.Text = ""
                End If

                Debug.WriteLine("ScheduleUIManager: ნავიგაციის ღილაკები კონფიგურირებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: ConfigureNavigationButtons შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონაცემების ცხრილის გასუფთავება
        ''' </summary>
        Public Sub ClearGrid()
            Try
                dgvSchedule.Rows.Clear()
                Debug.WriteLine("ScheduleUIManager: DataGridView გასუფთავებულია")
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: ClearGrid შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონიშნული მწკრივის ID-ის მიღება
        ''' </summary>
        ''' <returns>მონიშნული სესიის ID ან Nothing</returns>
        Public Function GetSelectedSessionId() As Integer?
            Try
                If dgvSchedule.SelectedRows.Count > 0 Then
                    Dim selectedRow = dgvSchedule.SelectedRows(0)
                    Dim sessionIdValue = selectedRow.Cells("N").Value

                    If sessionIdValue IsNot Nothing AndAlso IsNumeric(sessionIdValue) Then
                        Return CInt(sessionIdValue)
                    End If
                End If

                Return Nothing
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: GetSelectedSessionId შეცდომა: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' ღილაკზე დაჭერის შემოწმება
        ''' </summary>
        ''' <param name="e">CellClick ივენთის არგუმენტები</param>
        ''' <returns>True თუ რედაქტირების ღილაკზე დაიჭირა</returns>
        Public Function IsEditButtonClicked(e As DataGridViewCellEventArgs) As Boolean
            Try
                Return e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso
                       dgvSchedule.Columns(e.ColumnIndex).Name = "Edit"
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: IsEditButtonClicked შეცდომა: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' კონკრეტული უჯრის მნიშვნელობის მიღება
        ''' </summary>
        ''' <param name="columnName">სვეტის სახელი</param>
        ''' <param name="rowIndex">მწკრივის ინდექსი</param>
        ''' <returns>უჯრის მნიშვნელობა</returns>
        Public Function GetCellValue(columnName As String, rowIndex As Integer) As Object
            Try
                If rowIndex >= 0 AndAlso rowIndex < dgvSchedule.Rows.Count AndAlso
                   dgvSchedule.Columns.Contains(columnName) Then
                    Return dgvSchedule.Rows(rowIndex).Cells(columnName).Value
                End If

                Return Nothing
            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: GetCellValue შეცდომა: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' DataGridView-ის მწკრივების რაოდენობის მიღება
        ''' </summary>
        ''' <returns>მწკრივების რაოდენობა</returns>
        Public ReadOnly Property RowCount As Integer
            Get
                Try
                    Return dgvSchedule.Rows.Count
                Catch
                    Return 0
                End Try
            End Get
        End Property

        ''' <summary>
        ''' ფონის ფერების დაყენება კონტროლებისთვის
        ''' </summary>
        ''' <param name="controls">კონტროლების სია</param>
        Public Shared Sub SetBackgroundColors(ParamArray controls As Control())
            Try
                ' ღია გამჭვირვალე თეთრი ფერი
                Dim transparentWhite As Color = Color.FromArgb(200, Color.White)

                For Each control In controls
                    If control IsNot Nothing Then
                        control.BackColor = transparentWhite
                    End If
                Next

                Debug.WriteLine($"ScheduleUIManager: ფონის ფერები დაყენებულია {controls.Length} კონტროლისთვის")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleUIManager: SetBackgroundColors შეცდომა: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace