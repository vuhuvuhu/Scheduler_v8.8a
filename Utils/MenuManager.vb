' ===========================================
' 📄 Utils/MenuManager.vb
' -------------------------------------------
' მართავს MenuStrip–ს როლების მიხედვით:
'   1: სრულად
'   2: ყველაფერი გარდა ადმინისტრირებისა
'   3: ყველაფერი გარდა ფინანსებისა და ადმინისტრირებისა
'   4: everything except documents, finances, admin
'   5: მხოლოდ საწყისი და კალენდარი
'   6: მხოლოდ საწყისი
' ===========================================
Imports System.Windows.Forms

Public Class MenuManager

    Private ReadOnly menu As MenuStrip

    Private mnuHome As ToolStripMenuItem
    Private mnuCalendar As ToolStripMenuItem
    Private mnuDatabases As ToolStripMenuItem
    Private mnuGraphs As ToolStripMenuItem
    Private mnuDocuments As ToolStripMenuItem
    Private mnuFinances As ToolStripMenuItem
    Private mnuAdmin As ToolStripMenuItem

    Private mnuSchedule As ToolStripMenuItem
    Private mnuBeneficiaries As ToolStripMenuItem
    Private mnuTherapists As ToolStripMenuItem
    Private mnuTherapies As ToolStripMenuItem
    Private mnuFunding As ToolStripMenuItem

    Private mnuUserRegistration As ToolStripMenuItem

    Public Sub New(menuStrip As MenuStrip)
        menu = menuStrip
        InitializeMenu()
    End Sub

    Private Sub InitializeMenu()
        menu.Items.Clear()

        mnuHome = New ToolStripMenuItem("საწყისი")
        mnuCalendar = New ToolStripMenuItem("კალენდარი")
        mnuDatabases = New ToolStripMenuItem("ბაზები")
        mnuGraphs = New ToolStripMenuItem("გრაფიკები")
        mnuDocuments = New ToolStripMenuItem("დოკუმენტები")
        mnuFinances = New ToolStripMenuItem("ფინანსები")
        mnuAdmin = New ToolStripMenuItem("ადმინისტრირება")

        mnuSchedule = New ToolStripMenuItem("განრიგი")
        mnuBeneficiaries = New ToolStripMenuItem("ბენეფიციარები")
        mnuTherapists = New ToolStripMenuItem("თერაპევტები")
        mnuTherapies = New ToolStripMenuItem("თერაპიები")
        mnuFunding = New ToolStripMenuItem("დაფინანსება")

        mnuUserRegistration = New ToolStripMenuItem("მომხმარებელთა რეგისტრაცია")

        mnuDatabases.DropDownItems.AddRange({mnuSchedule, mnuBeneficiaries, mnuTherapists, mnuTherapies, mnuFunding})
        mnuAdmin.DropDownItems.Add(mnuUserRegistration)

        menu.Items.AddRange({mnuHome, mnuCalendar, mnuDatabases, mnuGraphs, mnuDocuments, mnuFinances, mnuAdmin})
        AddHandlers()
        ShowOnlyHomeMenu()
    End Sub

    Public Sub ShowAdminMenu()
        SetVisibility(True, True, True, True, True, True, True)
    End Sub

    Public Sub ShowManagerMenu()
        SetVisibility(True, True, True, True, True, True, False)
    End Sub

    Public Sub ShowRole3Menu()
        SetVisibility(True, True, True, True, True, False, False)
    End Sub

    Public Sub ShowRole4Menu()
        SetVisibility(True, True, True, True, False, False, False)
    End Sub

    Public Sub ShowRole5Menu()
        SetVisibility(True, True, False, False, False, False, False)
    End Sub

    Public Sub ShowOnlyHomeMenu()
        SetVisibility(True, False, False, False, False, False, False)
    End Sub

    Private Sub SetVisibility(home As Boolean, cal As Boolean, db As Boolean, graphs As Boolean, docs As Boolean, finances As Boolean, admin As Boolean)
        ' მთავარი მენიუთიემების ხილვადობის მართვა
        mnuHome.Visible = home
        mnuCalendar.Visible = cal
        mnuDatabases.Visible = db
        mnuGraphs.Visible = graphs
        mnuDocuments.Visible = docs
        mnuFinances.Visible = finances
        mnuAdmin.Visible = admin

        ' ქვემენიუს ელემენტების მართვა: მხოლოდ მაშინ დაშვებულია, როცა მშობელი ჩანს
        For Each itm As ToolStripMenuItem In mnuDatabases.DropDownItems
            itm.Visible = db
        Next
        For Each itm As ToolStripMenuItem In mnuAdmin.DropDownItems
            itm.Visible = admin
        Next
    End Sub

    Public Sub ShowMenuByRole(role As String)
        Select Case role.Trim()
            Case "1"
                ShowAdminMenu()
            Case "2"
                ShowManagerMenu()
            Case "3"
                ShowRole3Menu()
            Case "4"
                ShowRole4Menu()
            Case "5"
                ShowRole5Menu()
            Case "6"
                ShowOnlyHomeMenu()
            Case Else
                ShowOnlyHomeMenu()
        End Select
    End Sub

    Private Sub AddHandlers()
        Dim items = {mnuHome, mnuCalendar, mnuDatabases, mnuGraphs, mnuDocuments, mnuFinances, mnuAdmin,
                     mnuSchedule, mnuBeneficiaries, mnuTherapists, mnuTherapies, mnuFunding, mnuUserRegistration}
        For Each it As ToolStripMenuItem In items
            AddHandler it.Click, AddressOf OnMenuItemClicked
        Next
    End Sub

    Private Sub OnMenuItemClicked(sender As Object, e As EventArgs)
        Dim itm = DirectCast(sender, ToolStripMenuItem)
        MessageBox.Show(itm.Text)
    End Sub

End Class
