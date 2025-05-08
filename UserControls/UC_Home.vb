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

            ' გავასუფთაოთ არსებული ბარათები
            GBRedTasks.Controls.Clear()

            ' თუ სესიები არ არის, დავამატოთ შეტყობინება
            If sessions Is Nothing OrElse sessions.Count = 0 Then
                Dim lblNoSessions As New Label()
                lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
                lblNoSessions.AutoSize = True
                lblNoSessions.Dock = DockStyle.Fill
                lblNoSessions.TextAlign = ContentAlignment.MiddleCenter
                GBRedTasks.Controls.Add(lblNoSessions)
                Debug.WriteLine("UC_Home.PopulateOverdueSessions: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
                Return
            End If

            ' ბარათების განსათავსებლად FlowLayoutPanel
            Dim flowPanel As New FlowLayoutPanel()
            flowPanel.Dock = DockStyle.Fill
            flowPanel.AutoScroll = True
            flowPanel.FlowDirection = FlowDirection.TopDown
            flowPanel.WrapContents = False
            GBRedTasks.Controls.Add(flowPanel)
            Debug.WriteLine("UC_Home.PopulateOverdueSessions: FlowLayoutPanel დაემატა")

            ' ყველა სესიისთვის შევქმნათ ბარათი
            For Each session In sessions
                Debug.WriteLine($"UC_Home.PopulateOverdueSessions: ვქმნი ბარათს სესიისთვის ID: {session.Id}, თარიღი: {session.DateTime}, სტატუსი: {session.Status}")

                Dim card As New Panel()
                card.Size = New Size(GBRedTasks.Width - 30, 120)
                card.BorderStyle = BorderStyle.FixedSingle
                card.BackColor = Color.FromArgb(255, 235, 235) ' მოწითალო ფერი
                card.Margin = New Padding(5)

                ' დავამატოთ სესიის ინფორმაცია
                Dim lblDateTime As New Label()
                lblDateTime.Text = $"თარიღი: {session.FormattedDateTime}"
                lblDateTime.Location = New Point(10, 10)
                lblDateTime.AutoSize = True
                lblDateTime.Font = New Font(lblDateTime.Font, FontStyle.Bold)
                card.Controls.Add(lblDateTime)

                Dim lblDuration As New Label()
                lblDuration.Text = $"ხანგრძლივობა: {session.Duration} წთ"
                lblDuration.Location = New Point(10, 30)
                lblDuration.AutoSize = True
                card.Controls.Add(lblDuration)

                Dim lblBeneficiary As New Label()
                lblBeneficiary.Text = $"ბენეფიციარი: {session.BeneficiaryName} {session.BeneficiarySurname}"
                lblBeneficiary.Location = New Point(10, 50)
                lblBeneficiary.AutoSize = True
                card.Controls.Add(lblBeneficiary)

                Dim lblTherapist As New Label()
                lblTherapist.Text = $"თერაპევტი: {session.TherapistName}"
                lblTherapist.Location = New Point(10, 70)
                lblTherapist.AutoSize = True
                card.Controls.Add(lblTherapist)

                Dim lblTherapyType As New Label()
                lblTherapyType.Text = $"თერაპია: {session.TherapyType}"
                lblTherapyType.Location = New Point(250, 10)
                lblTherapyType.AutoSize = True
                card.Controls.Add(lblTherapyType)

                Dim lblSpace As New Label()
                lblSpace.Text = $"სივრცე: {session.Space}"
                lblSpace.Location = New Point(250, 30)
                lblSpace.AutoSize = True
                card.Controls.Add(lblSpace)

                Dim lblFunding As New Label()
                lblFunding.Text = $"დაფინანსება: {session.Funding}"
                lblFunding.Location = New Point(250, 50)
                lblFunding.AutoSize = True
                card.Controls.Add(lblFunding)

                ' თუ მომხმარებელი ავტორიზებულია და როლი არის 1, 2 ან 3, დავამატოთ რედაქტირების ღილაკი
                If isAuthorized AndAlso (userRole = "1" OrElse userRole = "2" OrElse userRole = "3") Then
                    Dim btnEdit As New Button()
                    btnEdit.Text = "რედაქტირება"
                    btnEdit.Location = New Point(card.Width - 120, 70)
                    btnEdit.Size = New Size(100, 25)
                    btnEdit.Tag = session.Id ' შევინახოთ სესიის ID
                    AddHandler btnEdit.Click, AddressOf BtnEditSession_Click
                    card.Controls.Add(btnEdit)
                End If

                ' დავამატოთ ბარათი FlowLayoutPanel-ში
                flowPanel.Controls.Add(card)
                Debug.WriteLine($"UC_Home.PopulateOverdueSessions: ბარათი დაემატა სესიისთვის ID: {session.Id}")
            Next

            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაემატა {sessions.Count} ბარათი")

            ' დარწმუნდეთ რომ GBRedTasks ხილვადია
            GBRedTasks.Visible = True

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