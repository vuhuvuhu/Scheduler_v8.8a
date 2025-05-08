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

Public Class UC_Home
    Inherits UserControl

    ' ViewModel რომელზეც ხდება მიბმა
    Private ReadOnly viewModel As HomeViewModel

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

        ' ალტერნატიული პანელის შექმნა
        CreateAlternativePanel()
    End Sub

    Private Sub CreateAlternativePanel()
        Try
            ' შევქმნათ ახალი პანელი
            Dim altPanel As New Panel()
            altPanel.Size = New Size(300, 300)
            altPanel.Location = New Point(20, 300)
            altPanel.BorderStyle = BorderStyle.FixedSingle
            altPanel.BackColor = Color.LightBlue

            ' სათაურის ლეიბლი
            Dim titleLabel As New Label()
            titleLabel.Text = "ალტერნატიული პანელი"
            titleLabel.AutoSize = True
            titleLabel.Location = New Point(10, 10)
            titleLabel.Font = New Font(titleLabel.Font.FontFamily, 12, FontStyle.Bold)

            altPanel.Controls.Add(titleLabel)

            ' ღილაკები
            Dim btn1 As New Button()
            btn1.Text = "ღილაკი 1"
            btn1.Size = New Size(100, 30)
            btn1.Location = New Point(20, 50)

            Dim btn2 As New Button()
            btn2.Text = "ღილაკი 2"
            btn2.Size = New Size(100, 30)
            btn2.Location = New Point(150, 50)

            altPanel.Controls.Add(btn1)
            altPanel.Controls.Add(btn2)

            ' დავამატოთ UserControl-ზე
            Me.Controls.Add(altPanel)
            altPanel.BringToFront()

            Debug.WriteLine("CreateAlternativePanel: შეიქმნა ალტერნატიული პანელი")
        Catch ex As Exception
            Debug.WriteLine($"CreateAlternativePanel შეცდომა: {ex.Message}")
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
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაიწყო, სესიების რაოდენობა: {If(sessions Is Nothing, 0, sessions.Count)}")

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
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: შეიქმნა ახალი პანელი ვადაგადაცილებული სესიებისთვის")
            Else
                ' გავასუფთაოთ არსებული პანელი
                overduePanel.Controls.Clear()
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: არსებული პანელი გასუფთავდა")
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
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
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

                Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაემატა ბარათი სესიისთვის ID: {session.Id}")
            Next

            ' პანელის წინა პლანზე გამოტანა
            overduePanel.BringToFront()
            overduePanel.Refresh()
            Application.DoEvents()

            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაემატებულია {Math.Min(15, sessions.Count)} ბარათი")

        Catch ex As Exception
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: შეცდომა - {ex.Message}")
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

            Debug.WriteLine("TestDirectControls: დაემატა 2 კონტროლი პირდაპირ UC_Home-ზე")
        Catch ex As Exception
            Debug.WriteLine($"TestDirectControls შეცდომა: {ex.Message}")
        End Try
    End Sub
End Class