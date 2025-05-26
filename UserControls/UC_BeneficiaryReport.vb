' ===========================================
' 📄 UserControls/UC_BeneficiaryReport.vb - კოდის ნაწილი მხოლოდ
' -------------------------------------------
' ბენეფიციარის ანგარიშის UserControl
' Designer.vb-ის ელემენტებთან სამუშაოდ
' 
' 🎯 ხელმისაწვდომია როლებისთვის: 1 (ადმინი), 2 (მენეჯერი), 3
' ===========================================
Imports System.ComponentModel
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports System.Text
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_BeneficiaryReport

    ' სერვისები და მონაცემები
    Private dataService As IDataService
    Private allSessions As List(Of SessionModel)
    Private filteredSessions As List(Of SessionModel)
    Private currentBeneficiarySessions As List(Of SessionModel)
    Private printService As AdvancedDataGridViewPrintService

    ' მომხმარებლის ინფორმაცია
    Private userEmail As String = ""
    Private userRole As String = ""

    ''' <summary>
    ''' ბენეფიციარის ინფორმაციის კლასი
    ''' </summary>
    Public Class BeneficiaryInfo
        Public Property FullName As String
        Public Property FirstName As String
        Public Property LastName As String
        Public Property TotalSessions As Integer
        Public Property CompletedSessions As Integer
        Public Property CancelledSessions As Integer
        Public Property PendingSessions As Integer
        Public Property TotalAmount As Decimal
        Public Property LastSessionDate As DateTime?
        Public Property FirstSessionDate As DateTime?
    End Class

    ''' <summary>
    ''' მონაცემთა სერვისის მითითება
    ''' </summary>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_BeneficiaryReport: მონაცემთა სერვისი მითითებულია")
    End Sub

    ''' <summary>
    ''' მომხმარებლის ინფორმაციის მითითება
    ''' </summary>
    Public Sub SetUserInfo(email As String, role As String)
        userEmail = email
        userRole = role
        Debug.WriteLine($"UC_BeneficiaryReport: მომხმარებლის ინფორმაცია - Email: {email}, Role: {role}")
    End Sub

    ''' <summary>
    ''' UserControl-ის ჩატვირთვისას
    ''' </summary>
    Private Sub UC_BeneficiaryReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            InitializeControls()
            LoadInitialData()
        Catch ex As Exception
            Debug.WriteLine($"UC_BeneficiaryReport_Load: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' კონტროლების საწყისი ინიციალიზაცია
    ''' </summary>
    Private Sub InitializeControls()
        Try
            ' თარიღების საწყისი მნიშვნელობები
            DtpDan.Value = DateTime.Today.AddMonths(-1) ' 1 თვის წინ
            DtpMde.Value = DateTime.Today ' დღეს

            ' ComboBox-ების საწყისი მდგომარეობა
            CBBeneName.Enabled = False
            CBBeneSurname.Enabled = False

            CBBeneName.Items.Clear()
            CBBeneSurname.Items.Clear()

            ' სესიების DataGridView-ის საწყისი მდგომარეობა
            If DgvSessions IsNot Nothing Then
                DgvSessions.DataSource = Nothing
            End If

            Debug.WriteLine("UC_BeneficiaryReport: კონტროლები ინიციალიზებულია")

        Catch ex As Exception
            Debug.WriteLine($"InitializeControls: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' საწყისი მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadInitialData()
        Try
            If dataService Is Nothing Then
                Debug.WriteLine("LoadInitialData: dataService არ არის ხელმისაწვდომი")
                Return
            End If

            ' ყველა სესიის წამოღება
            allSessions = dataService.GetAllSessions()
            Debug.WriteLine($"LoadInitialData: წამოღებულია {allSessions.Count} სესია")

            ' საწყისი ფილტრაცია
            FilterSessionsByDateRange()

        Catch ex As Exception
            Debug.WriteLine($"LoadInitialData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DtpDan თარიღის შეცვლისას
    ''' </summary>
    Private Sub DtpDan_ValueChanged(sender As Object, e As EventArgs) Handles DtpDan.ValueChanged
        Try
            ' ვალიდაცია - საწყისი თარიღი არ უნდა იყოს უფრო გვიან ვიდრე დასასრული
            If DtpDan.Value > DtpMde.Value Then
                DtpMde.Value = DtpDan.Value
            End If

            FilterSessionsByDateRange()
            LoadBeneficiaryNames()

        Catch ex As Exception
            Debug.WriteLine($"DtpDan_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' DtpMde თარიღის შეცვლისას
    ''' </summary>
    Private Sub DtpMde_ValueChanged(sender As Object, e As EventArgs) Handles DtpMde.ValueChanged
        Try
            ' ვალიდაცია - დასასრული თარიღი არ უნდა იყოს უფრო ადრე ვიდრე საწყისი
            If DtpMde.Value < DtpDan.Value Then
                DtpDan.Value = DtpMde.Value
            End If

            FilterSessionsByDateRange()
            LoadBeneficiaryNames()

        Catch ex As Exception
            Debug.WriteLine($"DtpMde_ValueChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ფილტრაცია თარიღების მიხედვით
    ''' </summary>
    Private Sub FilterSessionsByDateRange()
        Try
            If allSessions Is Nothing OrElse allSessions.Count = 0 Then
                filteredSessions = New List(Of SessionModel)()
                Return
            End If

            Dim startDate As DateTime = DtpDan.Value.Date
            Dim endDate As DateTime = DtpMde.Value.Date.AddDays(1).AddTicks(-1) ' დღის ბოლომდე

            filteredSessions = allSessions.Where(Function(s) s.DateTime >= startDate AndAlso s.DateTime <= endDate).ToList()

            Debug.WriteLine($"FilterSessionsByDateRange: ფილტრირებულია {filteredSessions.Count} სესია პერიოდისთვის {startDate:dd.MM.yyyy} - {DtpMde.Value.Date:dd.MM.yyyy}")

        Catch ex As Exception
            Debug.WriteLine($"FilterSessionsByDateRange: შეცდომა - {ex.Message}")
            filteredSessions = New List(Of SessionModel)()
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარების სახელების ჩატვირთვა CBBeneName-ში
    ''' </summary>
    Private Sub LoadBeneficiaryNames()
        Try
            CBBeneName.Items.Clear()
            CBBeneSurname.Items.Clear()
            CBBeneName.Text = ""
            CBBeneSurname.Text = ""

            If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                CBBeneName.Enabled = False
                CBBeneSurname.Enabled = False
                Debug.WriteLine("LoadBeneficiaryNames: ფილტრირებული სესიები ცარიელია")
                Return
            End If

            ' უნიკალური სახელების მიღება ანბანის მიხედვით
            Dim uniqueNames = filteredSessions.Where(Function(s) Not String.IsNullOrEmpty(s.BeneficiaryName)) _
                                           .Select(Function(s) s.BeneficiaryName.Trim()) _
                                           .Distinct() _
                                           .OrderBy(Function(name) name) _
                                           .ToList()

            ' სახელების დამატება ComboBox-ში
            For Each Name In uniqueNames
                CBBeneName.Items.Add(Name)
            Next

            CBBeneName.Enabled = True
            CBBeneSurname.Enabled = False

            Debug.WriteLine($"LoadBeneficiaryNames: ჩატვირთულია {uniqueNames.Count} უნიკალური სახელი")

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiaryNames: შეცდომა - {ex.Message}")
            CBBeneName.Enabled = False
            CBBeneSurname.Enabled = False
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneName-ის მნიშვნელობის შეცვლისას
    ''' </summary>
    Private Sub CBBeneName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneName.SelectedIndexChanged
        Try
            CBBeneSurname.Items.Clear()
            CBBeneSurname.Text = ""

            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) Then
                CBBeneSurname.Enabled = False
                ClearBeneficiaryData()
                Return
            End If

            LoadBeneficiarySurnames(selectedName)

        Catch ex As Exception
            Debug.WriteLine($"CBBeneName_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' არჩეული სახელის შესაბამისი გვარების ჩატვირთვა
    ''' </summary>
    ''' <param name="selectedName">არჩეული სახელი</param>
    Private Sub LoadBeneficiarySurnames(selectedName As String)
        Try
            If filteredSessions Is Nothing OrElse filteredSessions.Count = 0 Then
                CBBeneSurname.Enabled = False
                Return
            End If

            ' არჩეული სახელის შესაბამისი უნიკალური გვარების მიღება
            Dim uniqueSurnames = filteredSessions.Where(Function(s) s.BeneficiaryName.Trim().Equals(selectedName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                Not String.IsNullOrEmpty(s.BeneficiarySurname)) _
                                                 .Select(Function(s) s.BeneficiarySurname.Trim()) _
                                                 .Distinct() _
                                                 .OrderBy(Function(surname) surname) _
                                                 .ToList()

            ' გვარების დამატება ComboBox-ში
            For Each surname In uniqueSurnames
                CBBeneSurname.Items.Add(surname)
            Next

            CBBeneSurname.Enabled = True

            ' თუ მხოლოდ ერთი გვარია, ავტომატურად აირჩიოს
            If uniqueSurnames.Count = 1 Then
                CBBeneSurname.SelectedIndex = 0
            End If

            Debug.WriteLine($"LoadBeneficiarySurnames: სახელისთვის '{selectedName}' ჩატვირთულია {uniqueSurnames.Count} უნიკალური გვარი")

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiarySurnames: შეცდომა - {ex.Message}")
            CBBeneSurname.Enabled = False
        End Try
    End Sub

    ''' <summary>
    ''' CBBeneSurname-ის მნიშვნელობის შეცვლისას
    ''' </summary>
    Private Sub CBBeneSurname_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBeneSurname.SelectedIndexChanged
        Try
            Dim selectedName As String = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname As String = CBBeneSurname.SelectedItem?.ToString()

            If String.IsNullOrEmpty(selectedName) OrElse String.IsNullOrEmpty(selectedSurname) Then
                ClearBeneficiaryData()
                Return
            End If

            LoadBeneficiaryData(selectedName, selectedSurname)

        Catch ex As Exception
            Debug.WriteLine($"CBBeneSurname_SelectedIndexChanged: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის მონაცემების ჩატვირთვა
    ''' </summary>
    ''' <param name="name">ბენეფიციარის სახელი</param>
    ''' <param name="surname">ბენეფიციარის გვარი</param>
    Private Sub LoadBeneficiaryData(name As String, surname As String)
        Try
            ' ბენეფიციარის სესიების ფილტრაცია
            currentBeneficiarySessions = filteredSessions.Where(Function(s)
                                                                    s.BeneficiaryName.Trim().Equals(name, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                s.BeneficiarySurname.Trim().Equals(surname, StringComparison.OrdinalIgnoreCase)
                                                                End Function) _
                                                        .OrderByDescending(Function(s) s.DateTime) _
                                                        .ToList()

            If currentBeneficiarySessions.Count = 0 Then
                Debug.WriteLine($"LoadBeneficiaryData: ვერ მოიძებნა სესიები ბენეფიციარისთვის '{name} {surname}'")
                ClearBeneficiaryData()
                Return
            End If

            ' ბენეფიციარის ინფორმაციის შექმნა
            Dim beneficiaryInfo = CreateBeneficiaryInfo(name, surname, currentBeneficiarySessions)

            ' სტატისტიკის განახლება
            UpdateBeneficiaryStatistics(beneficiaryInfo)

            ' სესიების ჩვენება DataGridView-ში
            LoadSessionsToGrid(currentBeneficiarySessions)

            ' ღილაკების ააქტიურება
            EnableActionButtons(True)

            Debug.WriteLine($"LoadBeneficiaryData: ჩატვირთულია {currentBeneficiarySessions.Count} სესია ბენეფიციარისთვის '{name} {surname}'")

        Catch ex As Exception
            Debug.WriteLine($"LoadBeneficiaryData: შეცდომა - {ex.Message}")
            ClearBeneficiaryData()
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის ინფორმაციის შექმნა
    ''' </summary>
    ''' <param name="name">სახელი</param>
    ''' <param name="surname">გვარი</param>
    ''' <param name="sessions">სესიების სია</param>
    ''' <returns>ბენეფიციარის ინფორმაცია</returns>
    Private Function CreateBeneficiaryInfo(name As String, surname As String, sessions As List(Of SessionModel)) As BeneficiaryInfo
        Try
            Dim info As New BeneficiaryInfo()
            info.FirstName = name
            info.LastName = surname
            info.FullName = $"{name} {surname}"

            ' სტატისტიკის გამოთვლა
            info.TotalSessions = sessions.Count
            info.CompletedSessions = sessions.Where(Function(s) s.Status = "შესრულებული").Count()
            info.CancelledSessions = sessions.Where(Function(s) s.Status = "გაუქმებული").Count()
            info.PendingSessions = sessions.Where(Function(s) s.Status = "დაგეგმილი").Count()
            info.TotalAmount = sessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price)

            ' თარიღების დადგენა
            If sessions.Count > 0 Then
                info.FirstSessionDate = sessions.Min(Function(s) s.DateTime)
                info.LastSessionDate = sessions.Max(Function(s) s.DateTime)
            End If

            Return info

        Catch ex As Exception
            Debug.WriteLine($"CreateBeneficiaryInfo: შეცდომა - {ex.Message}")
            Return New BeneficiaryInfo()
        End Try
    End Function

    ''' <summary>
    ''' ბენეფიციარის სტატისტიკის განახლება UI-ზე
    ''' </summary>
    ''' <param name="info">ბენეფიციარის ინფორმაცია</param>
    Private Sub UpdateBeneficiaryStatistics(info As BeneficiaryInfo)
        Try
            ' ვამოწმებთ არსებობს თუ არა შესაბამისი ლეიბლები
            If LblBeneficiaryName IsNot Nothing Then
                LblBeneficiaryName.Text = $"ბენეფიციარი: {info.FullName}"
            End If

            If LblTotalSessions IsNot Nothing Then
                LblTotalSessions.Text = $"სულ სესიები: {info.TotalSessions}"
            End If

            If LblCompletedSessions IsNot Nothing Then
                LblCompletedSessions.Text = $"შესრულებული: {info.CompletedSessions}"
            End If

            If LblCancelledSessions IsNot Nothing Then
                LblCancelledSessions.Text = $"გაუქმებული: {info.CancelledSessions}"
            End If

            If LblPendingSessions IsNot Nothing Then
                LblPendingSessions.Text = $"დაგეგმილი: {info.PendingSessions}"
            End If

            If LblTotalAmount IsNot Nothing Then
                LblTotalAmount.Text = $"ჯამური თანხა: {info.TotalAmount:N2} ₾"
            End If

            If LblFirstSession IsNot Nothing Then
                LblFirstSession.Text = If(info.FirstSessionDate.HasValue,
                                       $"პირველი სესია: {info.FirstSessionDate.Value:dd.MM.yyyy}",
                                       "პირველი სესია: -")
            End If

            If LblLastSession IsNot Nothing Then
                LblLastSession.Text = If(info.LastSessionDate.HasValue,
                                      $"ბოლო სესია: {info.LastSessionDate.Value:dd.MM.yyyy}",
                                      "ბოლო სესია: -")
            End If

            Debug.WriteLine($"UpdateBeneficiaryStatistics: სტატისტიკა განახლებულია - {info.FullName}")

        Catch ex As Exception
            Debug.WriteLine($"UpdateBeneficiaryStatistics: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ჩატვირთვა DataGridView-ში
    ''' </summary>
    ''' <param name="sessions">სესიების სია</param>
    Private Sub LoadSessionsToGrid(sessions As List(Of SessionModel))
        Try
            If DgvSessions Is Nothing Then
                Debug.WriteLine("LoadSessionsToGrid: DgvSessions არ არსებობს")
                Return
            End If

            ' DataGridView-ის განახლება
            DgvSessions.DataSource = Nothing
            DgvSessions.DataSource = sessions

            ' მწკრივების ფერების განახლება არსებული SessionStatusColors კლასის გამოყენებით
            For Each row As DataGridViewRow In DgvSessions.Rows
                Try
                    Dim session = TryCast(row.DataBoundItem, SessionModel)
                    If session IsNot Nothing Then
                        ' გამოვიყენოთ არსებული SessionStatusColors კლასი
                        Dim statusColor = SessionStatusColors.GetStatusColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.BackColor = statusColor

                        ' ასევე ჩარჩოს ფერიც დავაყენოთ
                        Dim borderColor = SessionStatusColors.GetStatusBorderColor(session.Status, session.DateTime)
                        row.DefaultCellStyle.SelectionBackColor = borderColor
                    End If
                Catch
                    Continue For
                End Try
            Next

            Debug.WriteLine($"LoadSessionsToGrid: ჩატვირთულია {sessions.Count} სესია DataGridView-ში")

        Catch ex As Exception
            Debug.WriteLine($"LoadSessionsToGrid: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ღილაკების ააქტიურება/გათიშვა
    ''' </summary>
    ''' <param name="enabled">ააქტიურებული იყოს თუ არა</param>
    Private Sub EnableActionButtons(enabled As Boolean)
        Try
            If BtnExportToExcel IsNot Nothing Then
                BtnExportToExcel.Enabled = enabled
            End If

            If BtnPrint IsNot Nothing Then
                BtnPrint.Enabled = enabled
            End If

            If BtnGenerateReport IsNot Nothing Then
                BtnGenerateReport.Enabled = enabled
            End If

        Catch ex As Exception
            Debug.WriteLine($"EnableActionButtons: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ბენეფიციარის მონაცემების გასუფთავება
    ''' </summary>
    Private Sub ClearBeneficiaryData()
        Try
            currentBeneficiarySessions = New List(Of SessionModel)()

            ' სტატისტიკის გასუფთავება
            If LblBeneficiaryName IsNot Nothing Then LblBeneficiaryName.Text = "ბენეფიციარი: -"
            If LblTotalSessions IsNot Nothing Then LblTotalSessions.Text = "სულ სესიები: 0"
            If LblCompletedSessions IsNot Nothing Then LblCompletedSessions.Text = "შესრულებული: 0"
            If LblCancelledSessions IsNot Nothing Then LblCancelledSessions.Text = "გაუქმებული: 0"
            If LblPendingSessions IsNot Nothing Then LblPendingSessions.Text = "დაგეგმილი: 0"
            If LblTotalAmount IsNot Nothing Then LblTotalAmount.Text = "ჯამური თანხა: 0.00 ₾"
            If LblFirstSession IsNot Nothing Then LblFirstSession.Text = "პირველი სესია: -"
            If LblLastSession IsNot Nothing Then LblLastSession.Text = "ბოლო სესია: -"

            ' DataGridView-ის გასუფთავება
            If DgvSessions IsNot Nothing Then
                DgvSessions.DataSource = Nothing
            End If

            ' ღილაკების გათიშვა
            EnableActionButtons(False)

            Debug.WriteLine("ClearBeneficiaryData: ბენეფიციარის მონაცემები გაიწმინდა")

        Catch ex As Exception
            Debug.WriteLine($"ClearBeneficiaryData: შეცდომა - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Excel-ში ექსპორტის ღილაკზე დაჭერა (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnExportToExcel_Click(sender As Object, e As EventArgs) Handles BtnExportToExcel.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("ექსპორტისთვის ჯერ აირჩიეთ ბენეფიციარი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' CSV ექსპორტი
            ExportToCSV()

        Catch ex As Exception
            Debug.WriteLine($"BtnExportToExcel_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ექსპორტის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' CSV ფორმატში ექსპორტი
    ''' </summary>
    Private Sub ExportToCSV()
        Try
            Dim selectedName = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname = CBBeneSurname.SelectedItem?.ToString()

            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "CSV Files|*.csv"
                saveDialog.Title = "ბენეფიციარის ანგარიშის ექსპორტი"
                saveDialog.FileName = $"ბენეფიციარის_ანგარიში_{selectedName}_{selectedSurname}_{DateTime.Now:yyyyMMdd}.csv"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    Using writer As New System.IO.StreamWriter(saveDialog.FileName, False, System.Text.Encoding.UTF8)
                        ' BOM-ის დამატება Excel-ისთვის
                        writer.Write(System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble()))

                        ' სათაური ინფორმაცია
                        writer.WriteLine($"ბენეფიციარის ანგარიში,{selectedName} {selectedSurname}")
                        writer.WriteLine($"პერიოდი,{DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")
                        writer.WriteLine($"გენერირების თარიღი,{DateTime.Now:dd.MM.yyyy HH:mm}")
                        writer.WriteLine($"მომხმარებელი,{userEmail}")
                        writer.WriteLine()

                        ' სტატისტიკა
                        writer.WriteLine("სტატისტიკა")
                        writer.WriteLine($"სულ სესიები,{currentBeneficiarySessions.Count}")
                        writer.WriteLine($"შესრულებული,{currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Count()}")
                        writer.WriteLine($"გაუქმებული,{currentBeneficiarySessions.Where(Function(s) s.Status = "გაუქმებული").Count()}")
                        writer.WriteLine($"დაგეგმილი,{currentBeneficiarySessions.Where(Function(s) s.Status = "დაგეგმილი").Count()}")
                        writer.WriteLine($"ჯამური თანხა,{currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price):N2} ₾")
                        writer.WriteLine()

                        ' სესიების სათაურები
                        writer.WriteLine("ID,თარიღი,თერაპევტი,თერაპიის ტიპი,ხანგრძლივობა (წთ),ფასი (₾),სტატუსი,დაფინანსება,კომენტარები")

                        ' სესიების მონაცემები
                        For Each session In currentBeneficiarySessions
                            Dim line As String = $"{session.Id}," &
                                               $"{session.DateTime:dd.MM.yyyy HH:mm}," &
                                               $"""{session.TherapistName}""," &
                                               $"""{session.TherapyType}""," &
                                               $"{session.Duration}," &
                                               $"{session.Price:N2}," &
                                               $"""{session.Status}""," &
                                               $"""{session.Funding}""," &
                                               $"""{session.Comments.Replace("""", """""").Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")}"""
                            writer.WriteLine(line)
                        Next
                    End Using

                    MessageBox.Show($"ბენეფიციარის ანგარიში წარმატებით ექსპორტირდა:{Environment.NewLine}{saveDialog.FileName}",
                                  "ექსპორტი დასრულდა", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"ExportToCSV: შეცდომა - {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ბეჭდვის ღილაკზე დაჭერა (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("ბეჭდვისთვის ჯერ აირჩიეთ ბენეფიციარი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If DgvSessions Is Nothing Then
                MessageBox.Show("სესიების ცხრილი არ არის ხელმისაწვდომი", "შეცდომა",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ბეჭდვის სერვისის ინიციალიზაცია
            If printService IsNot Nothing Then
                printService.Dispose()
            End If

            printService = New AdvancedDataGridViewPrintService(DgvSessions)
            printService.ShowFullPrintDialog()

        Catch ex As Exception
            Debug.WriteLine($"BtnPrint_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დეტალური რეპორტის გენერირების ღილაკზე დაჭერა (თუ ღილაკი არსებობს)
    ''' </summary>
    Private Sub BtnGenerateReport_Click(sender As Object, e As EventArgs) Handles BtnGenerateReport.Click
        Try
            If currentBeneficiarySessions Is Nothing OrElse currentBeneficiarySessions.Count = 0 Then
                MessageBox.Show("რეპორტის გენერირებისთვის ჯერ აირჩიეთ ბენეფიციარი", "ინფორმაცია",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            GenerateDetailedReport()

        Catch ex As Exception
            Debug.WriteLine($"BtnGenerateReport_Click: შეცდომა - {ex.Message}")
            MessageBox.Show($"რეპორტის გენერირების შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' დეტალური რეპორტის გენერირება
    ''' </summary>
    Private Sub GenerateDetailedReport()
        Try
            Dim selectedName = CBBeneName.SelectedItem?.ToString()
            Dim selectedSurname = CBBeneSurname.SelectedItem?.ToString()

            ' დეტალური რეპორტის ტექსტის შექმნა
            Dim report As New StringBuilder()

            ' რეპორტის სათაური
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine($"           ბენეფიციარის დეტალური ანგარიში")
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine()

            ' ბენეფიციარის ინფორმაცია
            report.AppendLine("📋 ბენეფიციარის ინფორმაცია:")
            report.AppendLine($"   სახელი, გვარი: {selectedName} {selectedSurname}")
            report.AppendLine($"   პერიოდი: {DtpDan.Value:dd.MM.yyyy} - {DtpMde.Value:dd.MM.yyyy}")
            report.AppendLine($"   გენერირების თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}")
            report.AppendLine($"   მომხმარებელი: {userEmail}")
            report.AppendLine()

            ' სტატისტიკა
            report.AppendLine("📊 სტატისტიკა:")
            report.AppendLine($"   სულ სესიები: {currentBeneficiarySessions.Count}")
            report.AppendLine($"   შესრულებული: {currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Count()}")
            report.AppendLine($"   გაუქმებული: {currentBeneficiarySessions.Where(Function(s) s.Status = "გაუქმებული").Count()}")
            report.AppendLine($"   დაგეგმილი: {currentBeneficiarySessions.Where(Function(s) s.Status = "დაგეგმილი").Count()}")
            report.AppendLine($"   ჯამური თანხა: {currentBeneficiarySessions.Where(Function(s) s.Status = "შესრულებული").Sum(Function(s) s.Price):N2} ₾")

            If currentBeneficiarySessions.Count > 0 Then
                report.AppendLine($"   პირველი სესია: {currentBeneficiarySessions.Min(Function(s) s.DateTime):dd.MM.yyyy}")
                report.AppendLine($"   ბოლო სესია: {currentBeneficiarySessions.Max(Function(s) s.DateTime):dd.MM.yyyy}")
            End If
            report.AppendLine()

            ' სესიების დეტალები
            report.AppendLine("📅 სესიების დეტალები:")
            report.AppendLine("─────────────────────────────────────────────────────────────")

            For Each session In currentBeneficiarySessions
                report.AppendLine($"🔸 სესია #{session.Id}")
                report.AppendLine($"   თარიღი: {session.DateTime:dd.MM.yyyy HH:mm}")
                report.AppendLine($"   თერაპევტი: {session.TherapistName}")
                report.AppendLine($"   თერაპიის ტიპი: {session.TherapyType}")
                report.AppendLine($"   ხანგრძლივობა: {session.Duration} წუთი")
                report.AppendLine($"   ფასი: {session.Price:N2} ₾")
                report.AppendLine($"   სტატუსი: {session.Status}")
                report.AppendLine($"   დაფინანსება: {session.Funding}")
                If Not String.IsNullOrEmpty(session.Comments) Then
                    report.AppendLine($"   კომენტარები: {session.Comments}")
                End If
                report.AppendLine("─────────────────────────────────────────────────────────────")
            Next

            report.AppendLine()
            report.AppendLine("═══════════════════════════════════════════════════════════════")
            report.AppendLine($"რეპორტი გენერირებულია: {DateTime.Now:dd.MM.yyyy HH:mm:ss}")
            report.AppendLine("═══════════════════════════════════════════════════════════════")

            ' რეპორტის შენახვა ფაილად
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "Text Files|*.txt"
                saveDialog.Title = "რეპორტის შენახვა"
                saveDialog.FileName = $"ბენეფიციარის_ანგარიში_{selectedName}_{selectedSurname}_{DateTime.Now:yyyyMMdd}.txt"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    System.IO.File.WriteAllText(saveDialog.FileName, report.ToString(), System.Text.Encoding.UTF8)
                    MessageBox.Show($"რეპორტი შენახულია:{Environment.NewLine}{saveDialog.FileName}",
                                  "შენახვა დასრულდა", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"GenerateDetailedReport: შეცდომა - {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' რესურსების განთავისუფლება
    ''' </summary>
    Protected Overrides Sub Finalize()
        Try
            printService?.Dispose()
            allSessions?.Clear()
            filteredSessions?.Clear()
            currentBeneficiarySessions?.Clear()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class