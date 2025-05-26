' ===========================================
' 📄 Utils/MenuManager.vb (განახლებული ვერსია)
' -------------------------------------------
' მართავს MenuStrip–ს როლების მიხედვით:
'   1: სრულად
'   2: ყველაფერი გარდა ადმინისტრირებისა
'   3: ყველაფერი გარდა ფინანსებისა და ადმინისტრირებისა
'   4: everything except documents, finances, admin
'   5: მხოლოდ საწყისი და კალენდარი
'   6: მხოლოდ საწყისი
' 
' ✨ ახალი ფუნქციონალი: დოკუმენტების ქვემენიუები
'   - ბენეფიციარის ანგარიში (როლები: 1, 2, 3)
'   - თერაპევტის ანგარიში (როლები: 1, 2)
' ===========================================
Imports System.Windows.Forms

Public Class MenuManager

    Private ReadOnly menu As MenuStrip

    ' მთავარი მენიუს ელემენტები
    Private mnuHome As ToolStripMenuItem
    Private mnuCalendar As ToolStripMenuItem
    Private mnuDatabases As ToolStripMenuItem
    Private mnuGraphs As ToolStripMenuItem
    Private mnuDocuments As ToolStripMenuItem
    Private mnuFinances As ToolStripMenuItem
    Private mnuAdmin As ToolStripMenuItem

    ' ბაზების ქვემენიუს ელემენტები
    Private mnuSchedule As ToolStripMenuItem
    Private mnuBeneficiaries As ToolStripMenuItem
    Private mnuTherapists As ToolStripMenuItem
    Private mnuTherapies As ToolStripMenuItem
    Private mnuFunding As ToolStripMenuItem

    ' ✨ დოკუმენტების ქვემენიუს ელემენტები - ახალი!
    Private mnuBeneficiaryReport As ToolStripMenuItem
    Private mnuTherapistReport As ToolStripMenuItem

    ' ადმინისტრირების ქვემენიუს ელემენტები
    Private mnuUserRegistration As ToolStripMenuItem

    ' ივენთები განრიგის მენიუს არჩევისას
    Public Event ScheduleMenuSelected As EventHandler
    ' ✨ ახალი ივენთები - დოკუმენტების მენიუებისთვის
    Public Event BeneficiaryReportSelected As EventHandler
    Public Event TherapistReportSelected As EventHandler

    ''' <summary>
    ''' განრიგის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' </summary>
    Private Sub OnScheduleMenuClick(sender As Object, e As EventArgs)
        RaiseEvent ScheduleMenuSelected(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' ✨ ბენეფიციარის ანგარიშის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' </summary>
    Private Sub OnBeneficiaryReportMenuClick(sender As Object, e As EventArgs)
        RaiseEvent BeneficiaryReportSelected(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' ✨ თერაპევტის ანგარიშის მენიუს არჩევისას გამოძახებული მეთოდი
    ''' </summary>
    Private Sub OnTherapistReportMenuClick(sender As Object, e As EventArgs)
        RaiseEvent TherapistReportSelected(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' სხვა ქვემენიუს ელემენტების დამმუშავებელი
    ''' </summary>
    Private Sub OnSubmenuClick(sender As Object, e As EventArgs)
        ' აქ შეგვიძლია დავამატოთ სხვა ქვემენიუს ელემენტების დამუშავება საჭიროებისამებრ
    End Sub

    ''' <summary>
    ''' კონსტრუქტორი - ინიციალიზებს მენიუს MenuStrip-ით
    ''' </summary>
    ''' <param name="menuStrip">MenuStrip კონტროლი</param>
    Public Sub New(menuStrip As MenuStrip)
        menu = menuStrip
        InitializeMenu()
    End Sub

    ''' <summary>
    ''' მენიუს ინიციალიზაცია - ყველა ელემენტის შექმნა და კონფიგურაცია
    ''' </summary>
    Private Sub InitializeMenu()
        menu.Items.Clear()

        ' მთავარი მენიუს ელემენტების შექმნა
        mnuHome = New ToolStripMenuItem("საწყისი")
        mnuCalendar = New ToolStripMenuItem("კალენდარი")
        mnuDatabases = New ToolStripMenuItem("ბაზები")
        mnuGraphs = New ToolStripMenuItem("გრაფიკები")
        mnuDocuments = New ToolStripMenuItem("დოკუმენტები")
        mnuFinances = New ToolStripMenuItem("ფინანსები")
        mnuAdmin = New ToolStripMenuItem("ადმინისტრირება")

        ' ბაზების ქვემენიუს ელემენტების შექმნა
        mnuSchedule = New ToolStripMenuItem("განრიგი")
        mnuBeneficiaries = New ToolStripMenuItem("ბენეფიციარები")
        mnuTherapists = New ToolStripMenuItem("თერაპევტები")
        mnuTherapies = New ToolStripMenuItem("თერაპიები")
        mnuFunding = New ToolStripMenuItem("დაფინანსება")

        ' ✨ დოკუმენტების ქვემენიუს ელემენტების შექმნა - ახალი!
        mnuBeneficiaryReport = New ToolStripMenuItem("ბენეფიციარის ანგარიში")
        mnuTherapistReport = New ToolStripMenuItem("თერაპევტის ანგარიში")

        ' ადმინისტრირების ქვემენიუს ელემენტების შექმნა
        mnuUserRegistration = New ToolStripMenuItem("მომხმარებელთა რეგისტრაცია")

        ' ივენთის ჰენდლერების დამატება ბაზების ქვემენიუსთვის
        AddHandler mnuSchedule.Click, AddressOf OnScheduleMenuClick
        AddHandler mnuBeneficiaries.Click, AddressOf OnSubmenuClick
        AddHandler mnuTherapists.Click, AddressOf OnSubmenuClick
        AddHandler mnuTherapies.Click, AddressOf OnSubmenuClick
        AddHandler mnuFunding.Click, AddressOf OnSubmenuClick

        ' ✨ ივენთის ჰენდლერების დამატება დოკუმენტების ქვემენიუსთვის - ახალი!
        AddHandler mnuBeneficiaryReport.Click, AddressOf OnBeneficiaryReportMenuClick
        AddHandler mnuTherapistReport.Click, AddressOf OnTherapistReportMenuClick

        ' ივენთის ჰენდლერების დამატება ადმინისტრირების ქვემენიუსთვის
        AddHandler mnuUserRegistration.Click, AddressOf OnSubmenuClick

        ' ქვემენიუს ელემენტების დამატება მთავარ მენიუებზე
        mnuDatabases.DropDownItems.AddRange({mnuSchedule, mnuBeneficiaries, mnuTherapists, mnuTherapies, mnuFunding})

        ' ✨ დოკუმენტების ქვემენიუს დამატება - ახალი!
        mnuDocuments.DropDownItems.AddRange({mnuBeneficiaryReport, mnuTherapistReport})

        mnuAdmin.DropDownItems.Add(mnuUserRegistration)

        ' მთავარი მენიუს ელემენტების დამატება MenuStrip-ზე
        menu.Items.AddRange({mnuHome, mnuCalendar, mnuDatabases, mnuGraphs, mnuDocuments, mnuFinances, mnuAdmin})

        ' საწყისი მდგომარეობა - მხოლოდ საწყისი მენიუ
        ShowOnlyHomeMenu()
    End Sub

    ''' <summary>
    ''' ადმინისტრატორის მენიუ (როლი 1) - ყველაფერი ხილვადია
    ''' </summary>
    Public Sub ShowAdminMenu()
        SetMainMenuVisibility(True, True, True, True, True, True, True)
        SetDocumentsSubMenuVisibility(True, True) ' ✨ ორივე ანგარიში ხილვადია
    End Sub

    ''' <summary>
    ''' მენეჯერის მენიუ (როლი 2) - ყველაფერი გარდა ადმინისტრირებისა
    ''' </summary>
    Public Sub ShowManagerMenu()
        SetMainMenuVisibility(True, True, True, True, True, True, False)
        SetDocumentsSubMenuVisibility(True, True) ' ✨ ორივე ანგარიში ხილვადია
    End Sub

    ''' <summary>
    ''' როლი 3 - ყველაფერი გარდა ფინანსებისა და ადმინისტრირებისა
    ''' </summary>
    Public Sub ShowRole3Menu()
        SetMainMenuVisibility(True, True, True, True, True, False, False)
        SetDocumentsSubMenuVisibility(True, False) ' ✨ მხოლოდ ბენეფიციარის ანგარიში
    End Sub

    ''' <summary>
    ''' როლი 4 - ყველაფერი გარდა დოკუმენტებისა, ფინანსებისა და ადმინისტრირებისა
    ''' </summary>
    Public Sub ShowRole4Menu()
        SetMainMenuVisibility(True, True, True, True, False, False, False)
        SetDocumentsSubMenuVisibility(False, False) ' ✨ არცერთი ანგარიში არ ჩანს
    End Sub

    ''' <summary>
    ''' როლი 5 - მხოლოდ საწყისი და კალენდარი
    ''' </summary>
    Public Sub ShowRole5Menu()
        SetMainMenuVisibility(True, True, False, False, False, False, False)
        SetDocumentsSubMenuVisibility(False, False) ' ✨ არცერთი ანგარიში არ ჩანს
    End Sub

    ''' <summary>
    ''' როლი 6 - მხოლოდ საწყისი
    ''' </summary>
    Public Sub ShowOnlyHomeMenu()
        SetMainMenuVisibility(True, False, False, False, False, False, False)
        SetDocumentsSubMenuVisibility(False, False) ' ✨ არცერთი ანგარიში არ ჩანს
    End Sub

    ''' <summary>
    ''' მთავარი მენიუს ელემენტების ხილვადობის მართვა
    ''' </summary>
    ''' <param name="home">საწყისი</param>
    ''' <param name="cal">კალენდარი</param>
    ''' <param name="db">ბაზები</param>
    ''' <param name="graphs">გრაფიკები</param>
    ''' <param name="docs">დოკუმენტები</param>
    ''' <param name="finances">ფინანსები</param>
    ''' <param name="admin">ადმინისტრირება</param>
    Private Sub SetMainMenuVisibility(home As Boolean, cal As Boolean, db As Boolean,
                                     graphs As Boolean, docs As Boolean, finances As Boolean, admin As Boolean)
        ' მთავარი მენიუთიემების ხილვადობის მართვა
        mnuHome.Visible = home
        mnuCalendar.Visible = cal
        mnuDatabases.Visible = db
        mnuGraphs.Visible = graphs
        mnuDocuments.Visible = docs
        mnuFinances.Visible = finances
        mnuAdmin.Visible = admin

        ' ქვემენიუს ელემენტების მართვა ბაზებისთვის
        For Each itm As ToolStripMenuItem In mnuDatabases.DropDownItems
            itm.Visible = db
        Next

        ' ქვემენიუს ელემენტების მართვა ადმინისტრირებისთვის
        For Each itm As ToolStripMenuItem In mnuAdmin.DropDownItems
            itm.Visible = admin
        Next
    End Sub

    ''' <summary>
    ''' ✨ დოკუმენტების ქვემენიუს ხილვადობის მართვა - ახალი მეთოდი!
    ''' </summary>
    ''' <param name="beneficiaryReport">ბენეფიციარის ანგარიშის ხილვადობა</param>
    ''' <param name="therapistReport">თერაპევტის ანგარიშის ხილვადობა</param>
    Private Sub SetDocumentsSubMenuVisibility(beneficiaryReport As Boolean, therapistReport As Boolean)
        mnuBeneficiaryReport.Visible = beneficiaryReport
        mnuTherapistReport.Visible = therapistReport

        ' დებაგირებისთვის
        Debug.WriteLine($"SetDocumentsSubMenuVisibility: ბენეფიციარის ანგარიში={beneficiaryReport}, თერაპევტის ანგარიში={therapistReport}")
    End Sub

    ''' <summary>
    ''' მენიუს ჩვენება როლის მიხედვით - განახლებული მეთოდი
    ''' </summary>
    ''' <param name="role">მომხმარებლის როლი</param>
    Public Sub ShowMenuByRole(role As String)
        Debug.WriteLine($"ShowMenuByRole: როლი '{role}' - მენიუს კონფიგურაცია იწყება")

        Select Case role.Trim()
            Case "1"
                ShowAdminMenu()
                Debug.WriteLine("ShowMenuByRole: ადმინისტრატორის მენიუ (როლი 1) - ყველაფერი ხილვადია")

            Case "2"
                ShowManagerMenu()
                Debug.WriteLine("ShowMenuByRole: მენეჯერის მენიუ (როლი 2) - ადმინისტრირების გარდა")

            Case "3"
                ShowRole3Menu()
                Debug.WriteLine("ShowMenuByRole: როლი 3 - ფინანსებისა და ადმინისტრირების გარდა")

            Case "4"
                ShowRole4Menu()
                Debug.WriteLine("ShowMenuByRole: როლი 4 - დოკუმენტების, ფინანსებისა და ადმინისტრირების გარდა")

            Case "5"
                ShowRole5Menu()
                Debug.WriteLine("ShowMenuByRole: როლი 5 - მხოლოდ საწყისი და კალენდარი")

            Case "6"
                ShowOnlyHomeMenu()
                Debug.WriteLine("ShowMenuByRole: როლი 6 - მხოლოდ საწყისი")

            Case Else
                ShowOnlyHomeMenu()
                Debug.WriteLine($"ShowMenuByRole: უცნობი როლი '{role}' - ნაგულისხმევი (მხოლოდ საწყისი)")
        End Select
    End Sub

End Class