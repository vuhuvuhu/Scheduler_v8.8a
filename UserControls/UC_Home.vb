' ===========================================
' 📄 UserControls/UC_Home.vb
' -------------------------------------------
' მთავარი გვერდი - გამოყენებული HomeViewModel-ით და Data Binding-ით
' ===========================================
Imports System.Windows.Forms
Imports System.Drawing
Imports Scheduler_v8_8a.Models
Imports System.Globalization
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Home
    Inherits UserControl

    ' ViewModel რომელზეც ხდება მიბმა
    Private ReadOnly viewModel As HomeViewModel

    ' კეშირების და ოპტიმიზაციის ცვლადები
    Private cachedSessions As List(Of SessionModel) = Nothing ' კეშირებული სესიები
    Private sessionsLoaded As Boolean = False ' მოხდა თუ არა სესიების ჩატვირთვა
    Private isLoadingData As Boolean = False ' მიმდინარეობს თუ არა მონაცემების ჩატვირთვა

    ''' <summary>
    ''' კონსტრუქტორი: იღებს HomeViewModel-ს Data Binding-ისთვის
    ''' </summary>
    ''' <param name="homeVm">HomeViewModel ობიექტი</param>
    Public Sub New(homeVm As HomeViewModel)
        ' ViewModel-ის შენახვა
        viewModel = homeVm

        ' UI ელემენტების ინიციალიზაცია
        InitializeComponent()
        GBTools.Visible = False
        GBTools.Enabled = False
        ' Timer-ის დაყენება
        Timer1.Interval = 1000
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1.Start()

        ' GroupBox-ების გამჭვირვალე ფონის დაყენება (50%)
        GBGreeting.BackColor = Color.FromArgb(200, Color.White)
        GBNow.BackColor = Color.FromArgb(200, Color.White)
        GBTools.BackColor = Color.FromArgb(200, Color.White)
        GBRedTasks.BackColor = Color.FromArgb(200, Color.White)
        ' მიბმა ViewModel-ზე
        BindToViewModel()

        ' საწყისი მონაცემების ჩატვირთვა
        LoadData()

        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)
    End Sub

    ''' <summary>
    ''' ხილვადობის დაყენება ინსტრუმენტების პანელისთვის (GBTools)
    ''' </summary>
    ''' <param name="visible">უნდა იყოს თუ არა ხილული</param>
    Public Sub SetToolsVisibility(visible As Boolean)
        ' ინსტრუმენტების პანელი მხოლოდ მაშინ ჩანს, როცა მომხმარებელს აქვს შესაბამისი როლი
        GBTools.Visible = visible

        ' დავარეგულიროთ სხვა კონტეინერების ზომები, რათა არ იყოს ცარიელი ადგილი
        If visible Then
            ' ჩვეულებრივი მდგომარეობა, როცა ყველა პანელი ჩანს
            GBGreeting.Width = 253 ' საწყისი სიგანე
        Else
            ' როცა GBTools არ ჩანს, GBGreeting-ი უფრო ფართო უნდა იყოს
            GBGreeting.Width = 253 ' გაფართოებული სიგანე
        End If
    End Sub

    ''' <summary>
    ''' Timer-ის მოვლენის დამმუშავებელი - ანახლებს დროს
    ''' </summary>
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Try
            ' ViewModel-ში დროის განახლება - დავრწმუნდეთ რომ ar ხდება Exception-ი
            viewModel.UpdateTime()

            ' UI ელემენტების განახლება ViewModels-დან
            UpdateTimeDisplay()
        Catch ex As Exception
            ' შეცდომის დამუშავება - შესაძლოა Timer გამოვრთოთ თუ მუდმივად არის შეცდომა
            Debug.WriteLine($"Timer1_Tick შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' UI ელემენტების მიბმა ViewModel-ზე
    ''' </summary>
    Private Sub BindToViewModel()
        ' მიბმა ViewModel-ის PropertyChanged ივენთზე
        AddHandler viewModel.PropertyChanged, AddressOf ViewModel_PropertyChanged

        ' მისალმების დაყენება
        LUserName.Text = viewModel.UserName
        LWish.Text = viewModel.Greeting

        ' დროის დაყენება
        UpdateTimeDisplay()
    End Sub

    ''' <summary>
    ''' დროის ჩვენების განახლება ViewModel-დან
    ''' </summary>
    Private Sub UpdateTimeDisplay()
        ' UI-ს ვანახლებთ მხოლოდ თუ კონტროლი ხილვადია და ინიციალიზებულია
        If Me.IsHandleCreated AndAlso Me.Visible Then
            ' UI ნაკადის შემოწმება
            If Me.InvokeRequired Then
                Me.Invoke(Sub() UpdateTimeDisplay())
            Else
                ' UI-ს განახლება ViewModel-ის მონაცემებით
                LTime.Text = viewModel.FormattedTime
                LDate.Text = viewModel.FormattedDate
                LWeekDay.Text = viewModel.WeekDayName
            End If
        End If
    End Sub

    ''' <summary>
    ''' ViewModel-ის PropertyChanged ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub ViewModel_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs)
        ' ივენთის მართვა UI-ის მთავარ ნაკადში
        If Me.InvokeRequired Then
            Me.Invoke(Sub() ViewModel_PropertyChanged(sender, e))
            Return
        End If

        ' შესაბამისი თვისების განახლება UI-ში
        Select Case e.PropertyName
            Case NameOf(viewModel.UserName)
                LUserName.Text = viewModel.UserName
            Case NameOf(viewModel.Greeting)
                LWish.Text = viewModel.Greeting
            Case NameOf(viewModel.FormattedTime)
                LTime.Text = viewModel.FormattedTime
            Case NameOf(viewModel.FormattedDate)
                LDate.Text = viewModel.FormattedDate
            Case NameOf(viewModel.WeekDayName)
                LWeekDay.Text = viewModel.WeekDayName
        End Select
    End Sub

    ''' <summary>
    ''' მონაცემების ჩატვირთვა
    ''' </summary>
    Private Sub LoadData()
        ' TODO: ViewModel-ის RefreshData მეთოდის გამოძახება
        ' მონაცემების ჩატვირთვა (დაბადების დღეები, სესიები, დავალებები)
        viewModel.RefreshData()

        ' TODO: დაემატოს კოდი სხვა მონაცემების დასაყენებლად UI-ში
        ' დაბადების დღეები, სესიები, დავალებები და ა.შ.
    End Sub

    ''' <summary>
    ''' Refresh ღილაკის Click ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        ' მონაცემების ხელახლა ჩატვირთვა
        LoadData()
    End Sub

    ''' <summary>
    ''' UserControl-ის დატვირთვის ივენთი, რომელიც გაეშვება იგი პირველად ჩნდება
    ''' </summary>
    Private Sub UC_Home_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Timer-ის ხელახლა დაწყება თუ გაჩერებული იყო
        If Not Timer1.Enabled Then
            Timer1.Start()
        End If

        ' ვადაგადაცილებული სესიების პანელის შექმნა
        CreateOverdueSessionsPanel()
    End Sub

    ''' <summary>
    ''' ვადაგადაცილებული სესიების პანელის შექმნა
    ''' </summary>
    Private Sub CreateOverdueSessionsPanel()
        Try
            ' შევქმნათ პანელი
            Dim sessionsPanel As New Panel()
            sessionsPanel.Size = New Size(750, 300)
            sessionsPanel.Location = New Point(10, 150)
            sessionsPanel.BorderStyle = BorderStyle.FixedSingle
            sessionsPanel.BackColor = Color.FromArgb(255, 245, 245) ' ღია წითელი ფონი
            sessionsPanel.AutoScroll = True ' დავამატოთ სქროლი
            sessionsPanel.Tag = "OverdueSessionsPanel" ' ტეგი

            ' სათაურის ლეიბლი
            Dim titleLabel As New Label()
            titleLabel.Text = "ვადაგადაცილებული სესიები"
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(10, 10)
            titleLabel.Font = New Font(titleLabel.Font.FontFamily, 12, FontStyle.Bold)

            sessionsPanel.Controls.Add(titleLabel)

            ' ღილაკი სესიების ჩატვირთვისთვის
            Dim btnLoadSessions As New Button()
            btnLoadSessions.Text = "სესიების ჩატვირთვა"
            btnLoadSessions.Size = New Size(150, 30)
            btnLoadSessions.Location = New Point(250, 10)
            btnLoadSessions.BackColor = Color.FromArgb(0, 122, 204)
            btnLoadSessions.ForeColor = Color.White

            ' ღილაკზე დაჭერის ფუნქცია
            AddHandler btnLoadSessions.Click, AddressOf BtnLoadSessions_Click

            sessionsPanel.Controls.Add(btnLoadSessions)

            ' სტატუსის ლეიბლი
            Dim lblStatus As New Label()
            lblStatus.Text = "დააჭირეთ ღილაკს სესიების ჩატვირთვისთვის"
            lblStatus.AutoSize = True
            lblStatus.Location = New Point(10, 50)
            lblStatus.Name = "lblStatus" ' სახელი ძიებისთვის

            sessionsPanel.Controls.Add(lblStatus)

            ' დავამატოთ UserControl-ზე
            Me.Controls.Add(sessionsPanel)
            sessionsPanel.BringToFront()
        Catch ex As Exception
            Debug.WriteLine($"CreateOverdueSessionsPanel შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' სესიების ჩატვირთვის ღილაკზე დაჭერა - ოპტიმიზებული ვერსია
    ''' </summary>
    Private Sub BtnLoadSessions_Click(sender As Object, e As EventArgs)
        Try
            ' შევამოწმოთ არის თუ არა უკვე მიმდინარე დატვირთვა
            If isLoadingData Then
                MessageBox.Show("მონაცემები უკვე იტვირთება, გთხოვთ დაიცადოთ")
                Return
            End If

            ' მოვძებნოთ პანელი
            Dim sessionsPanel As Panel = Nothing

            For Each ctrl As Control In Me.Controls
                If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing AndAlso ctrl.Tag.ToString() = "OverdueSessionsPanel" Then
                    sessionsPanel = DirectCast(ctrl, Panel)
                    Exit For
                End If
            Next

            If sessionsPanel Is Nothing Then
                MessageBox.Show("ვერ მოიძებნა სესიების პანელი")
                Return
            End If

            ' მოვძებნოთ სტატუსის ლეიბლი
            Dim lblStatus As Label = Nothing
            For Each ctrl As Control In sessionsPanel.Controls
                If TypeOf ctrl Is Label AndAlso ctrl.Name = "lblStatus" Then
                    lblStatus = DirectCast(ctrl, Label)
                    Exit For
                End If
            Next

            If lblStatus Is Nothing Then
                lblStatus = New Label()
                lblStatus.AutoSize = True
                lblStatus.Location = New Point(10, 50)
                lblStatus.Name = "lblStatus"
                sessionsPanel.Controls.Add(lblStatus)
            End If

            ' ფორმის განახლება იმწამიერად
            lblStatus.Text = "მიმდინარეობს მონაცემების ჩატვირთვა..."
            sessionsPanel.Refresh()
            Application.DoEvents()

            ' ვიწყებთ ფონურ ჩატვირთვას
            isLoadingData = True

            ' გავასუფთაოთ არსებული ბარათები
            Dim controlsToRemove As New List(Of Control)
            For Each ctrl As Control In sessionsPanel.Controls
                If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing AndAlso ctrl.Tag.ToString() = "SessionCard" Then
                    controlsToRemove.Add(ctrl)
                End If
            Next

            For Each ctrl In controlsToRemove
                sessionsPanel.Controls.Remove(ctrl)
            Next

            ' ფონურ ნაკადში ვტვირთავთ სესიებს
            System.Threading.Tasks.Task.Run(Sub()
                                                Try
                                                    ' თუ უკვე გვაქვს კეშირებული სესიები, აღარ ვტვირთავთ ხელახლა
                                                    If Not sessionsLoaded OrElse cachedSessions Is Nothing Then
                                                        ' გამოვიძახოთ API და მივიღოთ სესიები
                                                        Dim dataForm = Form1.ActiveForm
                                                        If dataForm Is Nothing Then
                                                            Me.Invoke(Sub() lblStatus.Text = "ვერ მოიძებნა აქტიური ფორმა")
                                                            isLoadingData = False
                                                            Return
                                                        End If

                                                        ' გამოვიყენოთ რეფლექშენი dataService-ის მისაღებად
                                                        Dim dataServiceField = dataForm.GetType().GetField("dataService", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                                                        If dataServiceField Is Nothing Then
                                                            Me.Invoke(Sub() lblStatus.Text = "ვერ მოიძებნა dataService ველი")
                                                            isLoadingData = False
                                                            Return
                                                        End If

                                                        Dim dataService = DirectCast(dataServiceField.GetValue(dataForm), IDataService)
                                                        If dataService Is Nothing Then
                                                            Me.Invoke(Sub() lblStatus.Text = "dataService არის Nothing")
                                                            isLoadingData = False
                                                            Return
                                                        End If

                                                        ' განვაახლოთ სტატუსი
                                                        Me.Invoke(Sub() lblStatus.Text = "სესიების ჩატვირთვა Google Sheets-დან...")

                                                        ' მივიღოთ სესიები ფონურად
                                                        cachedSessions = dataService.GetOverdueSessions()
                                                        sessionsLoaded = True
                                                    End If

                                                    ' ჩავხატოთ UI (კეშირებული სესიებიდან)
                                                    Me.Invoke(Sub()
                                                                  lblStatus.Text = $"ნაპოვნია {cachedSessions.Count} ვადაგადაცილებული სესია"

                                                                  ' თუ სესიები ცარიელია
                                                                  If cachedSessions.Count = 0 Then
                                                                      lblStatus.Text = "ვადაგადაცილებული სესიები არ მოიძებნა"
                                                                      isLoadingData = False
                                                                      Return
                                                                  End If

                                                                  ' შეზღუდული რაოდენობა მხოლოდ
                                                                  Dim displayCount = Math.Min(10, cachedSessions.Count)
                                                                  Dim yPos As Integer = 80

                                                                  ' მხოლოდ პირველი რამდენიმე სესიის ჩვენება
                                                                  For i As Integer = 0 To displayCount - 1
                                                                      Dim session = cachedSessions(i)

                                                                      ' ბარათის შექმნა
                                                                      Dim card As New Panel()
                                                                      card.Size = New Size(sessionsPanel.Width - 40, 80)
                                                                      card.Location = New Point(20, yPos)
                                                                      card.BorderStyle = BorderStyle.FixedSingle
                                                                      card.BackColor = Color.FromArgb(255, 220, 220) ' უფრო ღია წითელი
                                                                      card.Tag = "SessionCard" ' ტეგი

                                                                      ' დავამატოთ ინფორმაცია
                                                                      Dim lblDateTime As New Label()
                                                                      lblDateTime.Text = $"თარიღი: {session.FormattedDateTime}"
                                                                      lblDateTime.Location = New Point(10, 10)
                                                                      lblDateTime.AutoSize = True
                                                                      lblDateTime.Font = New Font(lblDateTime.Font, FontStyle.Bold)
                                                                      card.Controls.Add(lblDateTime)

                                                                      Dim lblBeneficiary As New Label()
                                                                      lblBeneficiary.Text = $"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}"
                                                                      lblBeneficiary.Location = New Point(10, 30)
                                                                      lblBeneficiary.AutoSize = True
                                                                      card.Controls.Add(lblBeneficiary)

                                                                      Dim lblTherapist As New Label()
                                                                      lblTherapist.Text = $"თერაპევტი: {session.TherapistName}"
                                                                      lblTherapist.Location = New Point(10, 50)
                                                                      lblTherapist.AutoSize = True
                                                                      card.Controls.Add(lblTherapist)

                                                                      Dim lblId As New Label()
                                                                      lblId.Text = $"ID: {session.Id}"
                                                                      lblId.Location = New Point(card.Width - 60, 10)
                                                                      lblId.AutoSize = True
                                                                      lblId.Font = New Font(lblId.Font, FontStyle.Italic)
                                                                      card.Controls.Add(lblId)

                                                                      ' დავამატოთ ბარათი პანელზე
                                                                      sessionsPanel.Controls.Add(card)
                                                                      yPos += 90 ' შემდეგი ბარათისთვის
                                                                  Next

                                                                  ' დამატებითი სესიების არსებობის შემთხვევაში დავამატოთ ღილაკი
                                                                  If cachedSessions.Count > displayCount Then
                                                                      Dim btnLoadMore As New Button()
                                                                      btnLoadMore.Text = $"კიდევ {cachedSessions.Count - displayCount} სესიის ჩვენება"
                                                                      btnLoadMore.Size = New Size(200, 30)
                                                                      btnLoadMore.Location = New Point((sessionsPanel.Width - 200) / 2, yPos + 10)

                                                                      ' ღილაკზე დაჭერა
                                                                      AddHandler btnLoadMore.Click, Sub(s, args)
                                                                                                        ShowAllSessions(sessionsPanel)
                                                                                                    End Sub

                                                                      sessionsPanel.Controls.Add(btnLoadMore)
                                                                  End If

                                                                  ' დავასრულოთ ჩატვირთვა
                                                                  isLoadingData = False
                                                              End Sub)

                                                Catch ex As Exception
                                                    ' შეცდომის შემთხვევაში განვაახლოთ UI
                                                    Me.Invoke(Sub()
                                                                  lblStatus.Text = $"შეცდომა: {ex.Message}"
                                                                  isLoadingData = False
                                                              End Sub)

                                                    Debug.WriteLine($"BtnLoadSessions_Click შეცდომა: {ex.Message}")
                                                End Try
                                            End Sub)

        Catch ex As Exception
            MessageBox.Show($"შეცდომა სესიების ჩატვირთვისას: {ex.Message}")
            Debug.WriteLine($"BtnLoadSessions_Click შეცდომა: {ex.Message}")
            isLoadingData = False
        End Try
    End Sub

    ''' <summary>
    ''' ყველა სესიის ჩვენება
    ''' </summary>
    Private Sub ShowAllSessions(sessionsPanel As Panel)
        Try
            ' გავასუფთაოთ ყველა ბარათი და ღილაკი
            Dim controlsToRemove As New List(Of Control)

            For Each ctrl As Control In sessionsPanel.Controls
                If (TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing AndAlso ctrl.Tag.ToString() = "SessionCard") OrElse
                   (TypeOf ctrl Is Button AndAlso ctrl.Text.Contains("კიდევ")) Then
                    controlsToRemove.Add(ctrl)
                End If
            Next

            For Each ctrl In controlsToRemove
                sessionsPanel.Controls.Remove(ctrl)
            Next

            ' ჩავხატოთ ყველა სესია
            Dim yPos As Integer = 80

            For i As Integer = 0 To cachedSessions.Count - 1
                Dim session = cachedSessions(i)

                ' ბარათის შექმნა
                Dim card As New Panel()
                card.Size = New Size(sessionsPanel.Width - 40, 80)
                card.Location = New Point(20, yPos)
                card.BorderStyle = BorderStyle.FixedSingle
                card.BackColor = Color.FromArgb(255, 220, 220) ' უფრო ღია წითელი
                card.Tag = "SessionCard" ' ტეგი

                ' დავამატოთ ინფორმაცია
                Dim lblDateTime As New Label()
                lblDateTime.Text = $"თარიღი: {session.FormattedDateTime}"
                lblDateTime.Location = New Point(10, 10)
                lblDateTime.AutoSize = True
                lblDateTime.Font = New Font(lblDateTime.Font, FontStyle.Bold)
                card.Controls.Add(lblDateTime)

                Dim lblBeneficiary As New Label()
                lblBeneficiary.Text = $"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}"
                lblBeneficiary.Location = New Point(10, 30)
                lblBeneficiary.AutoSize = True
                card.Controls.Add(lblBeneficiary)

                Dim lblTherapist As New Label()
                lblTherapist.Text = $"თერაპევტი: {session.TherapistName}"
                lblTherapist.Location = New Point(10, 50)
                lblTherapist.AutoSize = True
                card.Controls.Add(lblTherapist)

                Dim lblId As New Label()
                lblId.Text = $"ID: {session.Id}"
                lblId.Location = New Point(card.Width - 60, 10)
                lblId.AutoSize = True
                lblId.Font = New Font(lblId.Font, FontStyle.Italic)
                card.Controls.Add(lblId)

                ' დავამატოთ ბარათი პანელზე
                sessionsPanel.Controls.Add(card)
                yPos += 90 ' შემდეგი ბარათისთვის
            Next

            ' მაჩვენებელი, რომ ყველა სესია ნაჩვენებია
            Dim lblAllShown As New Label()
            lblAllShown.Text = "ნაჩვენებია ყველა სესია"
            lblAllShown.AutoSize = True
            lblAllShown.Location = New Point((sessionsPanel.Width - 150) / 2, yPos + 10)
            lblAllShown.Font = New Font(lblAllShown.Font, FontStyle.Italic)
            sessionsPanel.Controls.Add(lblAllShown)

        Catch ex As Exception
            Debug.WriteLine($"ShowAllSessions შეცდომა: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' UserControl-ის დამალვის ივენთი, რომელიც გაეშვება როდესაც კონტროლი იმალება
    ''' </summary>
    Private Sub UC_Home_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        ' Timer-ი ვმუშაობს მხოლოდ როცა ხილულია
        Timer1.Enabled = Me.Visible
    End Sub

    ''' <summary>
    ''' ვადაგადაცილებული სესიების ბარათების შექმნა და GBRedTasks-ში განთავსება
    ''' </summary>
    Public Sub PopulateOverdueSessions(sessions As List(Of SessionModel), isAuthorized As Boolean, userRole As String)
        Try
            ' დარწმუნდით, რომ ამ მეთოდის გამოძახება ხილული და დახატულ კონტროლზე ხდება
            If Not Me.IsHandleCreated OrElse Not Me.Visible Then
                Return
            End If

            ' შევამოწმოთ არის თუ არა უკვე შექმნილი ჩვენი კონტეინერი
            Dim overduePanel As Panel = Nothing

            ' მოვძებნოთ არსებული პანელი Tag-ით (თუ უკვე შექმნილია)
            For Each ctrl As Control In Me.Controls
                If TypeOf ctrl Is Panel AndAlso ctrl.Tag IsNot Nothing AndAlso ctrl.Tag.ToString() = "OverdueSessionsPanel" Then
                    overduePanel = DirectCast(ctrl, Panel)
                    Exit For
                End If
            Next

            ' თუ ვერ ვიპოვეთ, შევქმნათ ახალი პანელი
            If overduePanel Is Nothing Then
                overduePanel = New Panel()
                overduePanel.Size = New Size(750, 300)
                overduePanel.Location = New Point(10, 150)
                overduePanel.BorderStyle = BorderStyle.FixedSingle
                overduePanel.BackColor = Color.FromArgb(255, 235, 235) ' მოწითალო ფონი
                overduePanel.AutoScroll = True ' დავამატოთ სქროლი
                overduePanel.Tag = "OverdueSessionsPanel" ' ტეგი ძიებისთვის

                Me.Controls.Add(overduePanel)
            Else
                ' გავასუფთაოთ არსებული პანელი
                overduePanel.Controls.Clear()
            End If

            ' სათაური
            Dim titleLabel As New Label()
            titleLabel.Text = $"ვადაგადაცილებული სესიები ({sessions.Count})"
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(10, 10)
            titleLabel.Font = New Font(titleLabel.Font.FontFamily, 12, FontStyle.Bold)
            overduePanel.Controls.Add(titleLabel)

            ' თუ სესიები არ არის, დავამატოთ შეტყობინება
            If sessions Is Nothing OrElse sessions.Count = 0 Then
                Dim lblNoSessions As New Label()
                lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
                lblNoSessions.AutoSize = True
                lblNoSessions.Location = New Point(10, 40)
                overduePanel.Controls.Add(lblNoSessions)
                Return
            End If

            ' ვადაგადაცილებული სესიების ბარათები
            Dim yPos As Integer = 40 ' საწყისი y პოზიცია ბარათებისთვის

            For i As Integer = 0 To Math.Min(15, sessions.Count - 1) ' პირველი 15 ბარათი მაინც
                Dim session = sessions(i)

                ' შევქმნათ ბარათი
                Dim card As New Panel()
                card.Size = New Size(overduePanel.Width - 40, 80)
                card.Location = New Point(20, yPos)
                card.BorderStyle = BorderStyle.FixedSingle
                card.BackColor = Color.FromArgb(255, 220, 220) ' უფრო ღია წითელი

                ' დავამატოთ სესიის ინფორმაცია
                Dim lblDateTime As New Label()
                lblDateTime.Text = $"თარიღი: {session.FormattedDateTime}"
                lblDateTime.Location = New Point(10, 10)
                lblDateTime.AutoSize = True
                lblDateTime.Font = New Font(lblDateTime.Font, FontStyle.Bold)
                card.Controls.Add(lblDateTime)

                Dim lblBeneficiary As New Label()
                lblBeneficiary.Text = $"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}"
                lblBeneficiary.Location = New Point(10, 30)
                lblBeneficiary.AutoSize = True
                card.Controls.Add(lblBeneficiary)

                Dim lblTherapist As New Label()
                lblTherapist.Text = $"თერაპევტი: {session.TherapistName}"
                lblTherapist.Location = New Point(10, 50)
                lblTherapist.AutoSize = True
                card.Controls.Add(lblTherapist)

                Dim lblId As New Label()
                lblId.Text = $"ID: {session.Id}"
                lblId.Location = New Point(card.Width - 60, 10)
                lblId.AutoSize = True
                lblId.Font = New Font(lblId.Font, FontStyle.Italic)
                card.Controls.Add(lblId)

                ' დავამატოთ ბარათი პანელზე
                overduePanel.Controls.Add(card)
                yPos += 90 ' შემდეგი ბარათისთვის
            Next

            ' პანელის წინა პლანზე გამოტანა
            overduePanel.BringToFront()
            overduePanel.Refresh()
            Application.DoEvents()

        Catch ex As Exception
            Debug.WriteLine($"PopulateOverdueSessions შეცდომა: {ex.Message}")
            MessageBox.Show($"შეცდომა ვადაგადაცილებული სესიების ასახვისას: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' რედაქტირების ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnEditSession_Click(sender As Object, e As EventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim sessionId As Integer = CInt(btn.Tag)

        ' აქ იქნება სესიის რედაქტირების ლოგიკა
        MessageBox.Show($"სესიის რედაქტირება ID: {sessionId}")
        ' შემდეგში შეიძლება გარე კლასს გადავაწოდოთ ეს ID ან გამოვიძახოთ რედაქტირების ფორმა
    End Sub

    ' ViewModel-ის თვისებების წაკითხვის მეთოდები UI-დან
    Public ReadOnly Property UserName() As String
        Get
            Return viewModel.UserName
        End Get
    End Property

    Public Sub TestDirectControls()
        Try
            ' შევიტანოთ სათაური
            Dim lblTest As New Label()
            lblTest.Text = "პირდაპირი ტესტი UC_Home-ზე"
            lblTest.AutoSize = True
            lblTest.Font = New Font(lblTest.Font.FontFamily, 16, FontStyle.Bold)
            lblTest.Location = New Point(50, 200)
            lblTest.BackColor = Color.Orange

            ' დავამატოთ UserControl-ზე (არა GBRedTasks-ზე)
            Me.Controls.Add(lblTest)

            ' დავამატოთ ღილაკი
            Dim btnTest As New Button()
            btnTest.Text = "საცდელი ღილაკი"
            btnTest.Size = New Size(150, 40)
            btnTest.Location = New Point(50, 250)
            btnTest.BackColor = Color.Green
            btnTest.ForeColor = Color.White

            Me.Controls.Add(btnTest)

            ' ფორსირებული განახლება
            lblTest.BringToFront()
            btnTest.BringToFront()
            Me.Refresh()
            Application.DoEvents()
        Catch ex As Exception
            Debug.WriteLine($"TestDirectControls შეცდომა: {ex.Message}")
        End Try
    End Sub
End Class