' ===========================================
' 📄 Services/CalendarGridManager.vb
' -------------------------------------------
' კალენდრის გრიდის მართვის სერვისი - შესწორებული ვერსია
' ===========================================
Imports System.Drawing
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Controls
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Services
    ''' <summary>
    ''' კალენდრის გრიდის მართვის სერვისი - პასუხისმგებელია გრიდის შექმნასა და განახლებაზე
    ''' </summary>
    Public Class CalendarGridManager
        ' კონსტანტები
        Public Const BASE_HEADER_HEIGHT As Integer = 40 ' სათაურების საბაზისო სიმაღლე - არ იცვლება
        Public Const BASE_TIME_COLUMN_WIDTH As Integer = 60 ' დროის სვეტის საბაზისო სიგანე - არ იცვლება
        Public Const BASE_ROW_HEIGHT As Integer = 40 ' მწკრივის საბაზისო სიმაღლე - იცვლება vScale-ით
        Public Const BASE_SPACE_COLUMN_WIDTH As Integer = 150 ' სივრცის სვეტის საბაზისო სიგანე - იცვლება hScale-ით
        Public Const BASE_DATE_COLUMN_WIDTH As Integer = 40 ' თარიღის ზოლის საბაზისო სიგანე - არ იცვლება მასშტაბით

        ' მთავარი პანელი
        Private ReadOnly calendarPanel As Panel

        ' გრიდის უჯრედები
        Public gridCells(,) As Panel

        ' ძირითადი პანელები
        Private mainGridPanel As Panel
        Private dateColumnPanel As Panel
        Private timeColumnPanel As Panel
        Private dateHeaderPanel As Panel
        Private timeHeaderPanel As Panel
        Private spacesHeaderPanel As Panel
        Private therapistsHeaderPanel As Panel
        Private beneficiaryHeaderPanel As Panel

        ' ხედვის ტიპი (გამოიყენება შესაბამისი ჰედერ პანელის მოსაძებნად)
        Private viewType As String = "space" ' შესაძლო მნიშვნელობები: "space", "therapist", "beneficiary"

        ' მასშტაბის ფაქტორები
        Private hScale As Double
        Private vScale As Double

        ' მონაცემები
        Private spaces As List(Of String)
        Private therapists As List(Of String)
        Private timeIntervals As List(Of DateTime)

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        Public Sub New(calendarPanel As Panel, spaces As List(Of String), therapists As List(Of String), timeIntervals As List(Of DateTime), hScale As Double, vScale As Double)
            Me.calendarPanel = calendarPanel

            ' ნალის შემოწმება - ცარიელი სიების ინიციალიზაცია საჭიროების შემთხვევაში
            Me.spaces = If(spaces, New List(Of String))
            Me.therapists = If(therapists, New List(Of String))
            Me.timeIntervals = If(timeIntervals, New List(Of DateTime))

            ' ვაყენებთ ხედვის ტიპს spaces/therapists პარამეტრების მიხედვით
            If therapists IsNot Nothing AndAlso therapists.Count > 0 Then
                viewType = "therapist"
            Else
                viewType = "space"
            End If

            Me.hScale = hScale
            Me.vScale = vScale

            Debug.WriteLine($"CalendarGridManager კონსტრუქტორი: viewType={viewType}, " &
                           $"spaces={If(spaces?.Count, 0)}, therapists={If(therapists?.Count, 0)}")
        End Sub

        ''' <summary>
        ''' ინიციალიზაცია დღის ხედისთვის
        ''' </summary>
        Public Sub InitializeDayViewPanels()
            Try
                Debug.WriteLine("InitializeDayViewPanels: დაიწყო პანელების ინიციალიზაცია")

                ' გავასუფთავოთ კალენდრის პანელი
                calendarPanel.Controls.Clear()

                ' მნიშვნელოვანი: calendarPanel-ს არ ჰქონდეს AutoScroll!
                calendarPanel.AutoScroll = False

                ' წავკითხოთ მშობელი პანელის ზომები
                Dim totalWidth As Integer = calendarPanel.ClientSize.Width
                Dim totalHeight As Integer = calendarPanel.ClientSize.Height

                Debug.WriteLine($"InitializeDayViewPanels: კონტეინერის ზომები - Width={totalWidth}, Height={totalHeight}")

                ' კონსტანტები, რომლებიც არ იცვლება მასშტაბირებისას
                Dim HEADER_HEIGHT As Integer = BASE_HEADER_HEIGHT
                Dim TIME_COLUMN_WIDTH As Integer = BASE_TIME_COLUMN_WIDTH
                Dim DATE_COLUMN_WIDTH As Integer = BASE_DATE_COLUMN_WIDTH ' ეს არ იმასშტაბირდება!

                ' გამოვთვალოთ დროისა და თარიღის პანელების ზომები
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)
                Dim totalRowsHeight As Integer = ROW_HEIGHT * timeIntervals.Count

                Debug.WriteLine($"InitializeDayViewPanels: ROW_HEIGHT={ROW_HEIGHT}, totalRowsHeight={totalRowsHeight}, vScale={vScale}")

                ' ======= 1. შევქმნათ თარიღის ზოლის პანელი (ყველაზე მარცხნივ) =======
                dateColumnPanel = New Panel()
                dateColumnPanel.Name = "dateColumnPanel"
                dateColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                dateColumnPanel.AutoScroll = False
                dateColumnPanel.Size = New Size(DATE_COLUMN_WIDTH, totalRowsHeight)
                dateColumnPanel.Location = New Point(0, HEADER_HEIGHT)
                dateColumnPanel.BackColor = Color.FromArgb(60, 80, 150)
                dateColumnPanel.BorderStyle = BorderStyle.FixedSingle
                calendarPanel.Controls.Add(dateColumnPanel)

                ' ======= 2. შევქმნათ საათების პანელი (თარიღის ზოლის მარჯვნივ) =======
                timeColumnPanel = New Panel()
                timeColumnPanel.Name = "timeColumnPanel"
                timeColumnPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                timeColumnPanel.AutoScroll = False
                timeColumnPanel.Size = New Size(TIME_COLUMN_WIDTH, totalRowsHeight)
                timeColumnPanel.Location = New Point(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
                timeColumnPanel.BackColor = Color.FromArgb(240, 240, 245)
                timeColumnPanel.BorderStyle = BorderStyle.FixedSingle
                calendarPanel.Controls.Add(timeColumnPanel)

                ' ======= 3. შევქმნათ თარიღის სათაურის პანელი (მარცხენა ზედა კუთხეში) =======
                dateHeaderPanel = New Panel()
                dateHeaderPanel.Name = "dateHeaderPanel"
                dateHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                dateHeaderPanel.Size = New Size(DATE_COLUMN_WIDTH, HEADER_HEIGHT)
                dateHeaderPanel.Location = New Point(0, 0)
                dateHeaderPanel.BackColor = Color.FromArgb(40, 60, 120)
                dateHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                calendarPanel.Controls.Add(dateHeaderPanel)

                ' თარიღის სათაურის ლეიბლი
                Dim dateHeaderLabel As New Label()
                dateHeaderLabel.Size = New Size(DATE_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
                dateHeaderLabel.Location = New Point(1, 1)
                dateHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
                dateHeaderLabel.Text = "თარიღი"
                dateHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
                dateHeaderLabel.ForeColor = Color.White
                dateHeaderPanel.Controls.Add(dateHeaderLabel)

                ' ======= 4. შევქმნათ დროის სათაურის პანელი (თარიღის სათაურის მარჯვნივ) =======
                timeHeaderPanel = New Panel()
                timeHeaderPanel.Name = "timeHeaderPanel"
                timeHeaderPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                timeHeaderPanel.Size = New Size(TIME_COLUMN_WIDTH, HEADER_HEIGHT)
                timeHeaderPanel.Location = New Point(DATE_COLUMN_WIDTH, 0)
                timeHeaderPanel.BackColor = Color.FromArgb(180, 180, 220)
                timeHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                calendarPanel.Controls.Add(timeHeaderPanel)

                ' დროის სათაურის ლეიბლი
                Dim timeHeaderLabel As New Label()
                timeHeaderLabel.Size = New Size(TIME_COLUMN_WIDTH - 2, HEADER_HEIGHT - 2)
                timeHeaderLabel.Location = New Point(1, 1)
                timeHeaderLabel.TextAlign = ContentAlignment.MiddleCenter
                timeHeaderLabel.Text = "დრო"
                timeHeaderLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)
                timeHeaderPanel.Controls.Add(timeHeaderLabel)

                ' ======= 5. შევქმნათ სათაურების პანელი ხედვის ტიპის მიხედვით =======
                Dim COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
                Dim itemsCount As Integer = 0

                ' ვადგენთ რომელი ხედვის ტიპი უნდა ვაჩვენოთ
                If viewType = "therapist" Then
                    ' თერაპევტების პანელი
                    therapistsHeaderPanel = New Panel()
                    therapistsHeaderPanel.Name = "therapistsHeaderPanel"
                    therapistsHeaderPanel.AutoScroll = False
                    itemsCount = Math.Max(1, therapists.Count) ' სულ მცირე 1 სვეტი
                    therapistsHeaderPanel.Size = New Size(COLUMN_WIDTH * itemsCount, HEADER_HEIGHT)
                    therapistsHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)
                    therapistsHeaderPanel.BackColor = Color.FromArgb(220, 220, 240)
                    therapistsHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                    calendarPanel.Controls.Add(therapistsHeaderPanel)
                Else
                    ' სივრცეების პანელი
                    spacesHeaderPanel = New Panel()
                    spacesHeaderPanel.Name = "spacesHeaderPanel"
                    spacesHeaderPanel.AutoScroll = False
                    itemsCount = Math.Max(1, spaces.Count) ' სულ მცირე 1 სვეტი
                    spacesHeaderPanel.Size = New Size(COLUMN_WIDTH * itemsCount, HEADER_HEIGHT)
                    spacesHeaderPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, 0)
                    spacesHeaderPanel.BackColor = Color.FromArgb(220, 220, 240)
                    spacesHeaderPanel.BorderStyle = BorderStyle.FixedSingle
                    calendarPanel.Controls.Add(spacesHeaderPanel)
                End If

                ' ======= 6. შევქმნათ მთავარი გრიდის პანელი =======
                mainGridPanel = New Panel()
                mainGridPanel.Name = "mainGridPanel"
                mainGridPanel.Size = New Size(totalWidth - TIME_COLUMN_WIDTH - DATE_COLUMN_WIDTH, totalHeight - HEADER_HEIGHT)
                mainGridPanel.Location = New Point(TIME_COLUMN_WIDTH + DATE_COLUMN_WIDTH, HEADER_HEIGHT)
                mainGridPanel.BackColor = Color.White
                mainGridPanel.BorderStyle = BorderStyle.FixedSingle

                ' მნიშვნელოვანი: მხოლოდ ეს პანელი არის სკროლებადი!
                mainGridPanel.AutoScroll = True
                calendarPanel.Controls.Add(mainGridPanel)

                Debug.WriteLine("InitializeDayViewPanels: პანელების ინიციალიზაცია დასრულდა")

            Catch ex As Exception
                Debug.WriteLine($"InitializeDayViewPanels: შეცდომა - {ex.Message}")
                Debug.WriteLine($"InitializeDayViewPanels: StackTrace - {ex.StackTrace}")
                MessageBox.Show($"პანელების ინიციალიზაციის შეცდომა: {ex.Message}", "შეცდომა", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' თარიღის ზოლის პანელის შევსება არჩეული თარიღის ინფორმაციით
        ''' </summary>
        Public Sub FillDateColumnPanel(selectedDate As DateTime)
            Try
                Debug.WriteLine("FillDateColumnPanel: დაიწყო თარიღის ზოლის შევსება")

                ' გავასუფთავოთ
                dateColumnPanel.Controls.Clear()

                ' ქართული კულტურა თარიღის ფორმატირებისთვის
                Dim georgianCulture As New Globalization.CultureInfo("ka-GE")

                ' კვირის დღე, რიცხვი, თვე და წელი
                Dim weekDay As String = selectedDate.ToString("dddd", georgianCulture)
                Dim dayOfMonth As String = selectedDate.ToString("dd", georgianCulture)
                Dim month As String = selectedDate.ToString("MMMM", georgianCulture)
                Dim year As String = selectedDate.ToString("yyyy", georgianCulture)

                ' შემობრუნებული ლეიბლის შექმნა
                Dim rotatedDateLabel As New Controls.RotatedLabel()
                rotatedDateLabel.Size = New Size(dateColumnPanel.Width - 4, dateColumnPanel.Height - 4)
                rotatedDateLabel.Location = New Point(2, 2)
                rotatedDateLabel.BackColor = Color.FromArgb(60, 80, 150)
                rotatedDateLabel.ForeColor = Color.White
                rotatedDateLabel.Font = New Font("Sylfaen", 10, FontStyle.Bold)

                ' ვერტიკალური ტექსტის ფორმირება
                Dim dateText As New System.Text.StringBuilder()
                dateText.Append(weekDay)
                dateText.Append("  ")
                dateText.Append(dayOfMonth)
                dateText.Append("  ")
                dateText.Append(month)
                dateText.Append("  ")
                dateText.Append(year)

                rotatedDateLabel.Text = dateText.ToString()
                rotatedDateLabel.RotationAngle = 270 ' შემობრუნება 270 გრადუსით

                ' დავამატოთ ლეიბლი პანელზე
                dateColumnPanel.Controls.Add(rotatedDateLabel)

                Debug.WriteLine($"FillDateColumnPanel: თარიღის ზოლი შევსებულია, პანელის ზომა: {dateColumnPanel.Size}")

            Catch ex As Exception
                Debug.WriteLine($"FillDateColumnPanel: შეცდომა - {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' დროის პანელის შევსება ინტერვალებით
        ''' </summary>
        Public Sub FillTimeColumnPanel()
            Try
                Debug.WriteLine("FillTimeColumnPanel: დაიწყო დროის პანელის შევსება")

                ' გავასუფთავოთ
                timeColumnPanel.Controls.Clear()

                ' მწკრივის სიმაღლე (იმასშტაბირდება)
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

                ' განვაახლოთ დროის პანელის სიმაღლე
                timeColumnPanel.Height = ROW_HEIGHT * timeIntervals.Count

                Debug.WriteLine($"FillTimeColumnPanel: დროის პანელის სიმაღლე: {timeColumnPanel.Height}, ROW_HEIGHT: {ROW_HEIGHT}, ინტერვალები: {timeIntervals.Count}")

                ' დროის ეტიკეტები
                For i As Integer = 0 To timeIntervals.Count - 1
                    Dim timeLabel As New Label()
                    timeLabel.Size = New Size(timeColumnPanel.Width - 5, ROW_HEIGHT)
                    timeLabel.Location = New Point(0, i * ROW_HEIGHT)
                    timeLabel.Text = timeIntervals(i).ToString("HH:mm")
                    timeLabel.TextAlign = ContentAlignment.MiddleRight
                    timeLabel.Padding = New Padding(0, 0, 5, 0)
                    timeLabel.Font = New Font("Segoe UI", 8, FontStyle.Regular)

                    ' შევინახოთ თავდაყრო Y პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                    timeLabel.Tag = i * ROW_HEIGHT

                    ' ალტერნატიული ფერები მწკრივებისთვის
                    If i Mod 2 = 0 Then
                        timeLabel.BackColor = Color.FromArgb(245, 245, 250)
                    Else
                        timeLabel.BackColor = Color.FromArgb(250, 250, 255)
                    End If

                    ' ჩარჩოს დამატება
                    timeLabel.BorderStyle = BorderStyle.FixedSingle

                    timeColumnPanel.Controls.Add(timeLabel)
                Next

                Debug.WriteLine($"FillTimeColumnPanel: დასრულდა - დაემატა {timeIntervals.Count} დროის ეტიკეტი")

            Catch ex As Exception
                Debug.WriteLine($"FillTimeColumnPanel: შეცდომა - {ex.Message}")
                Debug.WriteLine($"FillTimeColumnPanel: StackTrace: {ex.StackTrace}")
            End Try
        End Sub

        ''' <summary>
        ''' სივრცეების სათაურების პანელის შევსება
        ''' </summary>
        Public Sub FillSpacesHeaderPanel()
            Try
                Debug.WriteLine("FillSpacesHeaderPanel: დაიწყო სივრცეების სათაურების შევსება")

                ' გავასუფთავოთ
                If spacesHeaderPanel Is Nothing Then
                    Debug.WriteLine("FillSpacesHeaderPanel: spacesHeaderPanel არის Nothing, პანელის შევსებას ვწყვეტთ")
                    Return
                End If

                spacesHeaderPanel.Controls.Clear()

                ' სივრცის სვეტის სიგანე (იმასშტაბირდება)
                Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

                ' გამოვთვალოთ სათაურების პანელის სრული სიგანე
                Dim totalHeaderWidth As Integer = SPACE_COLUMN_WIDTH * Math.Max(1, spaces.Count)

                ' შევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
                spacesHeaderPanel.Width = totalHeaderWidth

                Debug.WriteLine($"FillSpacesHeaderPanel: პანელის სიგანე: {spacesHeaderPanel.Width}, SPACE_COLUMN_WIDTH: {SPACE_COLUMN_WIDTH}")

                ' მთავრული ფონტი
                Dim mtavruliFont As New Font("Sylfaen", 9, FontStyle.Bold)
                Try
                    If FontFamily.Families.Any(Function(f) f.Name = "BPG_Nino_Mtavruli") Then
                        mtavruliFont = New Font("BPG_Nino_Mtavruli", 9, FontStyle.Bold)
                    ElseIf FontFamily.Families.Any(Function(f) f.Name = "ALK_Tall_Mtavruli") Then
                        mtavruliFont = New Font("ALK_Tall_Mtavruli", 9, FontStyle.Bold)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"FillSpacesHeaderPanel: ფონტის შეცდომა - {ex.Message}")
                End Try

                ' სივრცეების სათაურების შექმნა
                For i As Integer = 0 To spaces.Count - 1
                    Dim spaceHeader As New Label()
                    spaceHeader.Size = New Size(SPACE_COLUMN_WIDTH - 1, spacesHeaderPanel.Height - 2)
                    spaceHeader.Location = New Point(i * SPACE_COLUMN_WIDTH, 1)
                    spaceHeader.BackColor = Color.FromArgb(60, 80, 150)
                    spaceHeader.ForeColor = Color.White
                    spaceHeader.TextAlign = ContentAlignment.MiddleCenter
                    spaceHeader.Text = spaces(i).ToUpper()
                    spaceHeader.Font = mtavruliFont
                    spaceHeader.BorderStyle = BorderStyle.FixedSingle

                    ' შევინახოთ X პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                    spaceHeader.Tag = i * SPACE_COLUMN_WIDTH

                    spacesHeaderPanel.Controls.Add(spaceHeader)
                Next

                Debug.WriteLine($"FillSpacesHeaderPanel: დაემატა {spaces.Count} სივრცის სათაური, მასშტაბი: {hScale}")

            Catch ex As Exception
                Debug.WriteLine($"FillSpacesHeaderPanel: შეცდომა - {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თერაპევტების სათაურების პანელის შევსება
        ''' </summary>
        Public Sub FillTherapistsHeaderPanel()
            Try
                Debug.WriteLine("FillTherapistsHeaderPanel: დაიწყო თერაპევტების სათაურების შევსება")

                ' გავასუფთავოთ
                If therapistsHeaderPanel Is Nothing Then
                    Debug.WriteLine("FillTherapistsHeaderPanel: therapistsHeaderPanel არის Nothing, პანელის შევსებას ვწყვეტთ")
                    Return
                End If

                therapistsHeaderPanel.Controls.Clear()

                ' თერაპევტის სვეტის სიგანე (იმასშტაბირდება)
                Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)

                ' გამოვთვალოთ სათაურების პანელის სრული სიგანე
                Dim totalHeaderWidth As Integer = THERAPIST_COLUMN_WIDTH * Math.Max(1, therapists.Count)

                ' შევცვალოთ პანელის სიგანე მასშტაბის მიხედვით
                therapistsHeaderPanel.Width = totalHeaderWidth

                Debug.WriteLine($"FillTherapistsHeaderPanel: პანელის სიგანე: {therapistsHeaderPanel.Width}, THERAPIST_COLUMN_WIDTH: {THERAPIST_COLUMN_WIDTH}")

                ' ფონტი თერაპევტებისთვის
                Dim therapistFont As New Font("Sylfaen", 9, FontStyle.Bold)

                ' თერაპევტების სათაურების შექმნა
                For i As Integer = 0 To therapists.Count - 1
                    Dim therapistHeader As New Label()
                    therapistHeader.Size = New Size(THERAPIST_COLUMN_WIDTH - 1, therapistsHeaderPanel.Height - 2)
                    therapistHeader.Location = New Point(i * THERAPIST_COLUMN_WIDTH, 1)
                    therapistHeader.BackColor = Color.FromArgb(80, 120, 80) ' მუქი მწვანე
                    therapistHeader.ForeColor = Color.White
                    therapistHeader.TextAlign = ContentAlignment.MiddleCenter
                    therapistHeader.Text = therapists(i)
                    therapistHeader.Font = therapistFont
                    therapistHeader.BorderStyle = BorderStyle.FixedSingle

                    ' შევინახოთ X პოზიცია Tag-ში სქროლის სინქრონიზაციისთვის
                    therapistHeader.Tag = i * THERAPIST_COLUMN_WIDTH

                    therapistsHeaderPanel.Controls.Add(therapistHeader)
                Next

                Debug.WriteLine($"FillTherapistsHeaderPanel: დაემატა {therapists.Count} თერაპევტის სათაური, მასშტაბი: {hScale}")

            Catch ex As Exception
                Debug.WriteLine($"FillTherapistsHeaderPanel: შეცდომა - {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მთავარი გრიდის პანელის შევსება
        ''' </summary>
        Public Sub FillMainGridPanel()
            Try
                Debug.WriteLine("FillMainGridPanel: დაიწყო მთავარი გრიდის შევსება")

                ' გავასუფთავოთ მთავარი პანელი
                mainGridPanel.Controls.Clear()

                ' პარამეტრები მასშტაბირების გამოყენებით
                Dim SPACE_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

                ' გრიდის სრული სიგანე (სულ მცირე 1 სვეტი)
                Dim gridWidth As Integer = SPACE_COLUMN_WIDTH * Math.Max(1, spaces.Count)

                ' გრიდის სრული სიმაღლე
                Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

                ' დავაყენოთ AutoScrollMinSize
                mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

                Debug.WriteLine($"FillMainGridPanel: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}, vScale={vScale}")

                ' უჯრედების მასივი
                Dim columnsCount As Integer = Math.Max(1, spaces.Count)
                Dim rowsCount As Integer = timeIntervals.Count
                ReDim gridCells(columnsCount - 1, rowsCount - 1)

                ' უჯრედების შექმნა
                For col As Integer = 0 To columnsCount - 1
                    For row As Integer = 0 To rowsCount - 1
                        Dim cell As New Panel()
                        cell.Size = New Size(SPACE_COLUMN_WIDTH, ROW_HEIGHT)
                        cell.Location = New Point(col * SPACE_COLUMN_WIDTH, row * ROW_HEIGHT)

                        ' ალტერნატიული ფერები
                        If row Mod 2 = 0 Then
                            cell.BackColor = Color.FromArgb(250, 250, 250)
                        Else
                            cell.BackColor = Color.FromArgb(245, 245, 245)
                        End If

                        ' უჯრედის ჩარჩო
                        AddHandler cell.Paint, AddressOf Cell_Paint

                        ' ვინახავთ უჯრედის ობიექტს მასივში
                        gridCells(col, row) = cell

                        ' მნიშვნელოვანი: ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
                        mainGridPanel.Controls.Add(cell)
                    Next
                Next

                Debug.WriteLine($"FillMainGridPanel: შეიქმნა {columnsCount * rowsCount} უჯრედი")

            Catch ex As Exception
                Debug.WriteLine($"FillMainGridPanel: შეცდომა - {ex.Message}")
                Debug.WriteLine($"FillMainGridPanel: StackTrace - {ex.StackTrace}")
            End Try
        End Sub

        ''' <summary>
        ''' მთავარი გრიდის პანელის შევსება თერაპევტებისთვის
        ''' </summary>
        Public Sub FillMainGridPanelForTherapists()
            Try
                Debug.WriteLine("FillMainGridPanelForTherapists: დაიწყო მთავარი გრიდის შევსება თერაპევტებისთვის")

                ' გავასუფთავოთ მთავარი პანელი
                mainGridPanel.Controls.Clear()

                ' პარამეტრები მასშტაბირების გამოყენებით
                Dim THERAPIST_COLUMN_WIDTH As Integer = CInt(BASE_SPACE_COLUMN_WIDTH * hScale)
                Dim ROW_HEIGHT As Integer = CInt(BASE_ROW_HEIGHT * vScale)

                ' გრიდის სრული სიგანე (სულ მცირე 1 სვეტი)
                Dim gridWidth As Integer = THERAPIST_COLUMN_WIDTH * Math.Max(1, therapists.Count)

                ' გრიდის სრული სიმაღლე
                Dim gridHeight As Integer = ROW_HEIGHT * timeIntervals.Count

                ' დავაყენოთ AutoScrollMinSize
                mainGridPanel.AutoScrollMinSize = New Size(gridWidth, gridHeight)

                Debug.WriteLine($"FillMainGridPanelForTherapists: გრიდის ზომები - Width={gridWidth}, Height={gridHeight}, hScale={hScale}, vScale={vScale}")

                ' უჯრედების მასივი
                Dim columnsCount As Integer = Math.Max(1, therapists.Count)
                Dim rowsCount As Integer = timeIntervals.Count
                ReDim gridCells(columnsCount - 1, rowsCount - 1)

                ' უჯრედების შექმნა
                For col As Integer = 0 To columnsCount - 1
                    For row As Integer = 0 To rowsCount - 1
                        Dim cell As New Panel()
                        cell.Size = New Size(THERAPIST_COLUMN_WIDTH, ROW_HEIGHT)
                        cell.Location = New Point(col * THERAPIST_COLUMN_WIDTH, row * ROW_HEIGHT)

                        ' ალტერნატიული ფერები (მწვანე ელფერით თერაპევტებისთვის)
                        If row Mod 2 = 0 Then
                            cell.BackColor = Color.FromArgb(250, 255, 250)
                        Else
                            cell.BackColor = Color.FromArgb(245, 250, 245)
                        End If

                        ' უჯრედის ჩარჩო
                        AddHandler cell.Paint, AddressOf Cell_Paint

                        ' ვინახავთ უჯრედის ობიექტს მასივში
                        gridCells(col, row) = cell

                        ' მნიშვნელოვანი: ვამატებთ უჯრედს პირდაპირ mainGridPanel-ზე
                        mainGridPanel.Controls.Add(cell)
                    Next
                Next

                Debug.WriteLine($"FillMainGridPanelForTherapists: შეიქმნა {columnsCount * rowsCount} უჯრედი თერაპევტებისთვის")

            Catch ex As Exception
                Debug.WriteLine($"FillMainGridPanelForTherapists: შეცდომა - {ex.Message}")
                Debug.WriteLine($"FillMainGridPanelForTherapists: StackTrace - {ex.StackTrace}")
            End Try
        End Sub

        ''' <summary>
        ''' უჯრედების მიღება
        ''' </summary>
        Public Function GetGridCells() As Panel(,)
            Return gridCells
        End Function

        ''' <summary>
        ''' უჯრედის ჩარჩოს დახატვა
        ''' </summary>
        Private Sub Cell_Paint(sender As Object, e As PaintEventArgs)
            Dim cell As Panel = DirectCast(sender, Panel)

            Using pen As New Pen(Color.FromArgb(200, 200, 200), 1)
                ' მხოლოდ მარჯვენა ბორდერი
                e.Graphics.DrawLine(pen, cell.Width - 1, 0, cell.Width - 1, cell.Height - 1)
            End Using
        End Sub

        ''' <summary>
        ''' სკროლის სინქრონიზაცია
        ''' </summary>
        Public Sub SynchronizeScroll(scrollPosition As Point)
            Try
                ' ვერტიკალური სქროლი
                Dim verticalOffset As Integer = scrollPosition.Y
                If timeColumnPanel IsNot Nothing Then
                    timeColumnPanel.Top = BASE_HEADER_HEIGHT + verticalOffset
                End If
                If dateColumnPanel IsNot Nothing Then
                    dateColumnPanel.Top = BASE_HEADER_HEIGHT + verticalOffset
                End If

                ' ჰორიზონტალური სქროლი
                Dim horizontalOffset As Integer = scrollPosition.X
                Dim fixedLeftPosition As Integer = BASE_TIME_COLUMN_WIDTH + BASE_DATE_COLUMN_WIDTH

                If spacesHeaderPanel IsNot Nothing AndAlso spacesHeaderPanel.Visible Then
                    spacesHeaderPanel.Left = fixedLeftPosition + horizontalOffset
                End If

                If therapistsHeaderPanel IsNot Nothing AndAlso therapistsHeaderPanel.Visible Then
                    therapistsHeaderPanel.Left = fixedLeftPosition + horizontalOffset
                End If

                If beneficiaryHeaderPanel IsNot Nothing AndAlso beneficiaryHeaderPanel.Visible Then
                    beneficiaryHeaderPanel.Left = fixedLeftPosition + horizontalOffset
                End If
            Catch ex As Exception
                Debug.WriteLine($"SynchronizeScroll: შეცდომა - {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მასშტაბის პარამეტრების განახლება
        ''' </summary>
        Public Sub UpdateScaleFactors(newHScale As Double, newVScale As Double)
            hScale = newHScale
            vScale = newVScale
        End Sub

        ''' <summary>
        ''' მასშტაბის ფაქტორების მიღება
        ''' </summary>
        Public Function GetScaleFactors() As (hScale As Double, vScale As Double)
            Return (hScale, vScale)
        End Function

        ''' <summary>
        ''' პანელების მიღება
        ''' </summary>
        Public Function GetPanels() As (mainPanel As Panel, datePanel As Panel, timePanel As Panel)
            Return (mainGridPanel, dateColumnPanel, timeColumnPanel)
        End Function

        ''' <summary>
        ''' ხედვის ტიპის მიღება
        ''' </summary>
        Public Function GetViewType() As String
            Return viewType
        End Function

        ''' <summary>
        ''' ხედვის ტიპის დაყენება
        ''' </summary>
        Public Sub SetViewType(newViewType As String)
            viewType = newViewType
        End Sub
    End Class
End Namespace