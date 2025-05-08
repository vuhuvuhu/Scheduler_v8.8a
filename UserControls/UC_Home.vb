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

            ' შევამოწმოთ კონტროლის მდგომარეობა
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: GBRedTasks.Size={GBRedTasks.Size}, Visible={GBRedTasks.Visible}, Dock={GBRedTasks.Dock}")

            ' გავასუფთაოთ არსებული ბარათები
            GBRedTasks.Controls.Clear()

            ' ბარათების ტესტისთვის
            Dim testHeaderPanel As New Panel()
            testHeaderPanel.BackColor = Color.Navy
            testHeaderPanel.Size = New Size(GBRedTasks.Width - 20, 30)
            testHeaderPanel.Location = New Point(10, 10)

            Dim testHeaderLabel As New Label()
            testHeaderLabel.Text = $"ნაპოვნია {sessions.Count} ვადაგადაცილებული სესია"
            testHeaderLabel.ForeColor = Color.White
            testHeaderLabel.AutoSize = True
            testHeaderLabel.Location = New Point(10, 5)

            testHeaderPanel.Controls.Add(testHeaderLabel)
            GBRedTasks.Controls.Add(testHeaderPanel)
            Debug.WriteLine("UC_Home.PopulateOverdueSessions: ტესტის სათაური დაემატა")

            ' თუ სესიები არ არის, დავამატოთ შეტყობინება
            If sessions Is Nothing OrElse sessions.Count = 0 Then
                Dim lblNoSessions As New Label()
                lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
                lblNoSessions.AutoSize = True
                lblNoSessions.Location = New Point(10, 50)
                GBRedTasks.Controls.Add(lblNoSessions)
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
                GBRedTasks.Visible = True
                GBRedTasks.Refresh()
                Return
            End If

            ' ყველა სესიისთვის შევქმნათ ბარათი, რომლებსაც პირდაპირ GBRedTasks-ში ვამატებთ
            Dim yPos As Integer = 50
            Dim cardCount As Integer = 0

            For i As Integer = 0 To Math.Min(10, sessions.Count - 1) ' პირველი 10 ბარათი მაინც
                Dim session = sessions(i)
                Debug.WriteLine($"UC_Home.PopulateOverdueSessions: ვქმნი ბარათს სესიისთვის ID: {session.Id}, თარიღი: {session.DateTime}, სტატუსი: {session.Status}")

                Dim card As New Panel()
                card.Size = New Size(GBRedTasks.Width - 40, 80)
                card.Location = New Point(20, yPos)
                card.BorderStyle = BorderStyle.FixedSingle
                card.BackColor = Color.FromArgb(255, 235, 235) ' მოწითალო ფერი

                ' დავამატოთ სესიის ინფორმაცია
                Dim lblDateTime As New Label()
                lblDateTime.Text = $"თარიღი: {session.FormattedDateTime}"
                lblDateTime.Location = New Point(10, 10)
                lblDateTime.AutoSize = True
                lblDateTime.Font = New Font(lblDateTime.Font, FontStyle.Bold)
                card.Controls.Add(lblDateTime)

                Dim lblBeneficiary As New Label()
                lblBeneficiary.Text = $"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}"
                lblBeneficiary.Location = New Point(10, 35)
                lblBeneficiary.AutoSize = True
                card.Controls.Add(lblBeneficiary)

                Dim lblId As New Label()
                lblId.Text = $"ID: {session.Id}"
                lblId.Location = New Point(10, 60)
                lblId.AutoSize = True
                lblId.Font = New Font(lblId.Font, FontStyle.Italic)
                card.Controls.Add(lblId)

                ' დავამატოთ ბარათი GBRedTasks-ში
                GBRedTasks.Controls.Add(card)
                cardCount += 1
                yPos += 90 ' შემდეგი ბარათისთვის Y პოზიცია

                Debug.WriteLine($"UC_Home.PopulateOverdueSessions: ბარათი დაემატა სესიისთვის ID: {session.Id}")
            Next

            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაემატა {cardCount} ბარათი")

            ' დარწმუნდეთ რომ GBRedTasks ხილვადია და განახლებულია
            GBRedTasks.Visible = True
            GBRedTasks.BringToFront()
            GBRedTasks.Refresh()
            Application.DoEvents()

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
End Class