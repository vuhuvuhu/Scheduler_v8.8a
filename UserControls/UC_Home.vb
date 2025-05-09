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

    ' პაგინაციის ცვლადები
    Private CurrentPage As Integer = 0
    Private TotalPages As Integer = 0
    Private CardsPerPage As Integer = 0

    ' სესიების სია მეხსიერებაში
    Private AllSessions As List(Of SessionModel) = New List(Of SessionModel)
    Private IsAuthorizedUser As Boolean = False
    Private UserRoleValue As String = ""

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

        ' პაგინაციის ღილაკების მომზადება
        BtnPrev.Text = "◄"
        BtnNext.Text = "►"
        BtnPrev.Enabled = False  ' საწყის მდგომარეობაში უკან ღილაკი გამორთულია
        LPage.Text = ""

        ' მიბმა ViewModel-ზე
        BindToViewModel()

        ' საწყისი მონაცემების ჩატვირთვა
        LoadData()

        ' საწყისი მდგომარეობაში ინსტრუმენტები დამალულია
        SetToolsVisibility(False)

        ' Resize ივენთის მიბმა
        AddHandler Me.Resize, AddressOf UC_Home_Resize
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
        ' ViewModel-ის RefreshData მეთოდის გამოძახება
        ' მონაცემების ჩატვირთვა (დაბადების დღეები, სესიები, დავალებები)
        viewModel.RefreshData()
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
    ''' UserControl-ის Resize ივენთის დამმუშავებელი
    ''' </summary>
    Private Sub UC_Home_Resize(sender As Object, e As EventArgs)
        ' შევამოწმოთ არის თუ არა ფორმა ინიციალიზებული
        If Not Me.IsHandleCreated OrElse GBRedTasks Is Nothing Then
            Return
        End If

        ' თუ სესიები გვაქვს, განვაახლოთ ბარათების ჩვენება
        If AllSessions.Count > 0 Then
            ' ვიანგარიშოთ თავიდან რამდენი ეტევა
            CalculateCardsPerPage()

            ' თუ მიმდინარე გვერდი აღარ არსებობს (მაგ. ფანჯარა გაიზარდა), დავაბრუნოთ პირველ გვერდზე
            If CurrentPage >= TotalPages Then
                CurrentPage = Math.Max(0, TotalPages - 1)
            End If

            ' ჩავტვირთოთ მიმდინარე გვერდის ბარათები
            ShowCurrentPageCards()

            ' განვაახლოთ პაგინაციის კონტროლები
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' ანგარიშობს რამდენი ბარათი ეტევა ერთ გვერდზე და საერთო გვერდების რაოდენობას
    ''' </summary>
    Private Sub CalculateCardsPerPage()
        Try
            ' ბარათის ზომები და დაშორებები
            Const CARD_WIDTH As Integer = 250
            Const CARD_HEIGHT As Integer = 140
            Const CARD_MARGIN As Integer = 15

            ' გამოვთვალოთ რამდენი ბარათი დაეტევა ერთ მწკრივში
            Dim availableWidth As Integer = GBRedTasks.Width - (2 * CARD_MARGIN)
            Dim cardsPerRow As Integer = Math.Max(1, availableWidth \ (CARD_WIDTH + CARD_MARGIN))

            ' გამოვთვალოთ რამდენი მწკრივი დაეტევა
            Dim availableHeight As Integer = GBRedTasks.Height - 50 ' გამოვაკლოთ ადგილი შეტყობინებებისთვის
            Dim rowsPerPage As Integer = Math.Max(1, availableHeight \ (CARD_HEIGHT + CARD_MARGIN))

            ' გამოვთვალოთ რამდენი ბარათი ეტევა ერთ გვერდზე
            CardsPerPage = cardsPerRow * rowsPerPage

            ' საერთო გვერდების რაოდენობა
            TotalPages = Math.Ceiling(AllSessions.Count / CardsPerPage)

            Debug.WriteLine($"CalculateCardsPerPage: cardsPerRow={cardsPerRow}, rowsPerPage={rowsPerPage}, " &
                          $"CardsPerPage={CardsPerPage}, TotalPages={TotalPages}")
        Catch ex As Exception
            Debug.WriteLine($"CalculateCardsPerPage: შეცდომა - {ex.Message}")
            CardsPerPage = 6 ' ნაგულისხმები მნიშვნელობა შეცდომის შემთხვევაში
            TotalPages = Math.Ceiling(AllSessions.Count / CardsPerPage)
        End Try
    End Sub

    ''' <summary>
    ''' აჩვენებს მიმდინარე გვერდის ბარათებს
    ''' </summary>
    Private Sub ShowCurrentPageCards()
        ' გასუფთავება ძველი ბარათებისგან
        GBRedTasks.Controls.Clear()

        ' თუ სესიები არ არის, დავამატოთ შეტყობინება
        If AllSessions Is Nothing OrElse AllSessions.Count = 0 Then
            Dim lblNoSessions As New Label()
            lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
            lblNoSessions.AutoSize = True
            lblNoSessions.Location = New Point(10, 20)
            GBRedTasks.Controls.Add(lblNoSessions)
            Debug.WriteLine("ShowCurrentPageCards: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
            Return
        End If

        ' ბარათის ზომები და დაშორებები
        Const CARD_WIDTH As Integer = 250
        Const CARD_HEIGHT As Integer = 180   ' ოდნავ გაზრდილი სიმაღლე
        Const CARD_MARGIN As Integer = 15
        Const HEADER_HEIGHT As Integer = 24

        ' გამოვთვალოთ რამდენი ბარათი დაეტევა ერთ მწკრივში
        Dim availableWidth As Integer = GBRedTasks.Width - (2 * CARD_MARGIN)
        Dim cardsPerRow As Integer = Math.Max(1, availableWidth \ (CARD_WIDTH + CARD_MARGIN))

        ' დავაკორექტიროთ ჰორიზონტალური დაშორება თუ კი საჭიროა
        Dim horizontalSpacing As Integer = CARD_MARGIN
        If cardsPerRow > 1 Then
            horizontalSpacing = (availableWidth - (cardsPerRow * CARD_WIDTH)) \ (cardsPerRow - 1)
        End If

        ' გვერდზე ბარათების დიაპაზონი
        Dim startIndex As Integer = CurrentPage * CardsPerPage
        Dim endIndex As Integer = Math.Min(startIndex + CardsPerPage - 1, AllSessions.Count - 1)

        Debug.WriteLine($"ShowCurrentPageCards: გვერდი {CurrentPage + 1}/{TotalPages}, " &
                    $"დიაპაზონი: {startIndex}-{endIndex}, " &
                    $"cardsPerRow={cardsPerRow}, horizontalSpacing={horizontalSpacing}")

        ' ვადაგადაცილებული სესიების ბარათები
        Dim xPos As Integer = CARD_MARGIN
        Dim yPos As Integer = CARD_MARGIN + 10 ' დამატებითი 10 პიქსელი ქვემოთ
        Dim cardCount As Integer = 0

        For i As Integer = startIndex To endIndex
            Dim session = AllSessions(i)

            ' დებაგინფო სესიის შესახებ
            Debug.WriteLine($"ShowCurrentPageCards: სესია #{i}: ID={session.Id}, " &
                       $"ბენეფიციარი={session.BeneficiaryName} {session.BeneficiarySurname}, " &
                       $"თარიღი={session.FormattedDateTime}, " &
                       $"სტატუსი={session.Status}")

            ' შევქმნათ ბარათი
            Dim card As New Panel()
            card.Size = New Size(CARD_WIDTH, CARD_HEIGHT)
            card.Location = New Point(xPos, yPos)
            card.BorderStyle = BorderStyle.None
            card.BackColor = Color.FromArgb(255, 235, 235) ' უფრო ღია წითელი

            ' მოვუმრგვალოთ კუთხეები ბარათს
            Dim path As New Drawing2D.GraphicsPath()
            Dim cornerRadius As Integer = 10
            path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
            path.AddArc(card.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
            path.AddArc(card.Width - cornerRadius * 2, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90)
            path.AddArc(0, card.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90)
            path.CloseFigure()
            card.Region = New Region(path)

            ' ზედა მუქი ზოლი
            Dim headerPanel As New Panel()
            headerPanel.Size = New Size(CARD_WIDTH, HEADER_HEIGHT)
            headerPanel.Location = New Point(0, 0)
            headerPanel.BackColor = Color.FromArgb(180, 0, 0) ' მუქი წითელი

            ' მოვუმრგვალოთ მხოლოდ ზედა კუთხეები header-ს
            Dim headerPath As New Drawing2D.GraphicsPath()
            headerPath.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90)
            headerPath.AddArc(headerPanel.Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90)
            headerPath.AddLine(headerPanel.Width, cornerRadius, headerPanel.Width, headerPanel.Height)
            headerPath.AddLine(headerPanel.Width, headerPanel.Height, 0, headerPanel.Height)
            headerPath.AddLine(0, headerPanel.Height, 0, cornerRadius)
            headerPath.CloseFigure()
            headerPanel.Region = New Region(headerPath)
            card.Controls.Add(headerPanel)

            ' თარიღი თეთრი ფონტით header-ში
            Dim lblDateTime As New Label()
            lblDateTime.Text = session.FormattedDateTime
            lblDateTime.Location = New Point(8, 4)
            lblDateTime.AutoSize = True
            lblDateTime.Font = New Font(lblDateTime.Font.FontFamily, 9, FontStyle.Regular)
            lblDateTime.ForeColor = Color.White
            headerPanel.Controls.Add(lblDateTime)

            ' ID მარჯვენა კუთხეში header-ში
            Dim lblId As New Label()
            lblId.Text = $"#{session.Id}"
            lblId.Location = New Point(headerPanel.Width - 40, 4)
            lblId.AutoSize = True
            lblId.Font = New Font(lblId.Font.FontFamily, 9, FontStyle.Regular)
            lblId.ForeColor = Color.White
            headerPanel.Controls.Add(lblId)

            ' ბენეფიციარის სახელი და გვარი - ბოლდით და მთავრული ასოებით
            ' მთავრული ქართული შრიფტი
            Dim georgianFonts As String() = {"Sylfaen", "BPG Glaho", "Arial Unicode MS", "DejaVu Sans", "FreeSerif", "Tahoma"}
            Dim georgianFont As String = "Sylfaen" ' ნაგულისხმევი ფონტი

            ' ვიპოვოთ რომელიმე ქართული ფონტი
            For Each fontName In georgianFonts
                Try
                    Using testFont As New Font(fontName, 10)
                        georgianFont = fontName
                        Debug.WriteLine($"ShowCurrentPageCards: მოიძებნა ქართული ფონტი: {fontName}")
                        Exit For
                    End Using
                Catch
                    ' ეს ფონტი ვერ მოიძებნა, შემდეგს ვცდით
                    Debug.WriteLine($"ShowCurrentPageCards: ვერ მოიძებნა ფონტი: {fontName}")
                End Try
            Next

            Dim lblBeneficiary As New Label()
            ' გარდავქმნათ ტექსტი დიდ ასოებად
            Dim beneficiaryText As String = $"{session.BeneficiaryName.ToUpper()} {session.BeneficiarySurname.ToUpper()}"
            lblBeneficiary.Text = beneficiaryText
            lblBeneficiary.Location = New Point(8, HEADER_HEIGHT + 8)
            lblBeneficiary.Size = New Size(CARD_WIDTH - 16, 24) ' ოდნავ უფრო მაღალი
            lblBeneficiary.Font = New Font(georgianFont, 10, FontStyle.Bold)
            card.Controls.Add(lblBeneficiary)

            ' თერაპევტის სახელი
            Dim lblTherapist As New Label()
            lblTherapist.Text = session.TherapistName
            lblTherapist.Location = New Point(8, HEADER_HEIGHT + 36)
            lblTherapist.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
            card.Controls.Add(lblTherapist)

            ' თერაპიის ტიპი
            Dim lblTherapyType As New Label()
            lblTherapyType.Text = session.TherapyType
            lblTherapyType.Location = New Point(8, HEADER_HEIGHT + 60)
            lblTherapyType.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
            card.Controls.Add(lblTherapyType)

            ' სივრცე
            Dim lblSpace As New Label()
            lblSpace.Text = session.Space
            lblSpace.Location = New Point(8, HEADER_HEIGHT + 84)
            lblSpace.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
            card.Controls.Add(lblSpace)

            ' დაფინანსება
            Dim lblFunding As New Label()
            lblFunding.Text = session.Funding
            lblFunding.Location = New Point(8, HEADER_HEIGHT + 108)
            lblFunding.Size = New Size(CARD_WIDTH - 16, 20) ' ფიქსირებული სიგანე
            card.Controls.Add(lblFunding)

            ' ვადაგადაცილების დღეები
            Dim overdueText As String = "ვადაგადაცილება: "
            Dim daysOverdue As Integer = Math.Abs((DateTime.Today - session.DateTime.Date).Days)
            overdueText += $"{daysOverdue} დღე"

            Dim lblOverdue As New Label()
            lblOverdue.Text = overdueText
            lblOverdue.Location = New Point(8, HEADER_HEIGHT + 132) ' ბოლო რიგი
            lblOverdue.AutoSize = True
            lblOverdue.ForeColor = Color.DarkRed
            card.Controls.Add(lblOverdue)

            ' რედაქტირების ღილაკი - მხოლოდ თუ ავტორიზებულია
            If IsAuthorizedUser AndAlso (UserRoleValue = "1" OrElse UserRoleValue = "2" OrElse UserRoleValue = "3") Then
                Dim btnEdit As New Button()
                btnEdit.Text = "✎" ' ფანქრის სიმბოლო
                btnEdit.Font = New Font("Segoe UI Symbol", 10, FontStyle.Bold)
                btnEdit.ForeColor = Color.White
                btnEdit.BackColor = Color.FromArgb(220, 0, 0)
                btnEdit.Size = New Size(30, 30) ' მრგვალი ზომა
                btnEdit.Location = New Point(CARD_WIDTH - 40, HEADER_HEIGHT + 123) ' აწეული 5 პიქსელით
                btnEdit.FlatStyle = FlatStyle.Flat
                btnEdit.FlatAppearance.BorderSize = 0
                btnEdit.Tag = session.Id ' შევინახოთ სესიის ID ღილაკის Tag-ში

                ' მრგვალი ფორმის ღილაკი
                Dim btnPath As New Drawing2D.GraphicsPath()
                btnPath.AddEllipse(0, 0, btnEdit.Width, btnEdit.Height)
                btnEdit.Region = New Region(btnPath)

                AddHandler btnEdit.Click, AddressOf BtnEditSession_Click
                card.Controls.Add(btnEdit)

                Debug.WriteLine($"ShowCurrentPageCards: ღილაკი დაემატა ბარათზე, პოზიცია: X={btnEdit.Location.X}, Y={btnEdit.Location.Y}")
            End If

            ' დავამატოთ ბარათი GroupBox-ზე
            GBRedTasks.Controls.Add(card)

            ' განვაახლოთ ბარათის პოზიცია შემდეგი ბარათისთვის
            cardCount += 1
            If cardCount Mod cardsPerRow = 0 Then
                ' ახალი მწკრივი
                xPos = CARD_MARGIN
                yPos += CARD_HEIGHT + CARD_MARGIN
            Else
                ' იგივე მწკრივი, შემდეგი სვეტი
                xPos += CARD_WIDTH + horizontalSpacing
            End If
        Next
    End Sub

    ''' <summary>
    ''' განაახლებს პაგინაციის კონტროლებს
    ''' </summary>
    Private Sub UpdatePaginationControls()
        ' LPage-ში ვაჩვენებთ მიმდინარე გვერდს და საერთო რაოდენობას
        LPage.Text = $"{CurrentPage + 1}/{TotalPages}"

        ' BtnPrev (წინა გვერდი) გამორთულია თუ პირველ გვერდზე ვართ
        BtnPrev.Enabled = (CurrentPage > 0)

        ' BtnNext (შემდეგი გვერდი) გამორთულია თუ ბოლო გვერდზე ვართ
        BtnNext.Enabled = (CurrentPage < TotalPages - 1)

        Debug.WriteLine($"UpdatePaginationControls: გვერდი {CurrentPage + 1}/{TotalPages}, " &
                        $"BtnPrev.Enabled={BtnPrev.Enabled}, BtnNext.Enabled={BtnNext.Enabled}")
    End Sub

    ''' <summary>
    ''' წინა გვერდის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnPrev_Click(sender As Object, e As EventArgs) Handles BtnPrev.Click
        If CurrentPage > 0 Then
            CurrentPage -= 1
            ShowCurrentPageCards()
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' შემდეგი გვერდის ღილაკზე დაჭერის დამმუშავებელი
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        If CurrentPage < TotalPages - 1 Then
            CurrentPage += 1
            ShowCurrentPageCards()
            UpdatePaginationControls()
        End If
    End Sub

    ''' <summary>
    ''' ვადაგადაცილებული სესიების ბარათების შექმნა და GBRedTasks-ში განთავსება
    ''' </summary>
    Public Sub PopulateOverdueSessions(sessions As List(Of SessionModel), isAuthorized As Boolean, userRole As String)
        Try
            Debug.WriteLine($"UC_Home.PopulateOverdueSessions: დაიწყო, სესიების რაოდენობა: {If(sessions Is Nothing, 0, sessions.Count)}")

            ' შევინახოთ მომხმარებლის მდგომარეობა
            IsAuthorizedUser = isAuthorized
            UserRoleValue = userRole

            ' გავასუფთავოთ ძველი სესიები და ჩავსვათ ახალი
            AllSessions = New List(Of SessionModel)()

            ' გადავატაროთ ყველა სესია, მოვიპოვოთ ცხრილიდან არსებული raw data და შევასწოროთ ველები
            If sessions IsNot Nothing Then
                For Each session In sessions
                    AllSessions.Add(session)
                Next
            End If

            ' თუ სესიები არ არის, დავამატოთ შეტყობინება
            If AllSessions Is Nothing OrElse AllSessions.Count = 0 Then
                ' გასუფთავება ძველი ბარათებისგან
                GBRedTasks.Controls.Clear()

                Dim lblNoSessions As New Label()
                lblNoSessions.Text = "ვადაგადაცილებული სესიები არ არის"
                lblNoSessions.AutoSize = True
                lblNoSessions.Location = New Point(10, 20)
                GBRedTasks.Controls.Add(lblNoSessions)

                ' პაგინაციის კონტროლები გამოვრთოთ
                BtnPrev.Enabled = False
                BtnNext.Enabled = False
                LPage.Text = ""

                Debug.WriteLine("UC_Home.PopulateOverdueSessions: დაემატა შეტყობინება 'ვადაგადაცილებული სესიები არ არის'")
                Return
            End If

            ' დავადგინოთ რამდენი ბარათი ეტევა თითო გვერდზე
            CalculateCardsPerPage()

            ' დავრწმუნდეთ რომ მიმდინარე გვერდი ვალიდურია
            CurrentPage = Math.Min(CurrentPage, Math.Max(0, TotalPages - 1))

            ' ვაჩვენოთ ბარათები
            ShowCurrentPageCards()

            ' განვაახლოთ პაგინაციის კონტროლები
            UpdatePaginationControls()

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