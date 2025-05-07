' ===========================================
' 📄 UserControls/UC_Home.vb
' -------------------------------------------
' მთავარი გვერდი (მხოლოდ Transparent Greeting GroupBox)
' ===========================================
Imports System.Windows.Forms
Imports System.Drawing
Imports Scheduler_v8_8a.Models
Imports NamespaceOfYourControl    ' სადაც TransparentGroupBox ახსნილია
Imports System.Globalization

Public Class UC_Home
    Inherits UserControl

    Private ReadOnly viewModel As HomeViewModel

    ''' <summary>
    ''' კონსტრუქტორი: იღებს MainViewModel-ის Greeting პროპერტისათვის და გამჭვირვალე ფონს GroupBox-ს
    ''' </summary>
    Public Sub New(homeVm As HomeViewModel)
        viewModel = homeVm
        InitializeComponent()

        Timer1.Interval = 1000
        'AddHandler Timer1.Tick
        Timer1.Start()
        ' GroupBox გამჭვირვალედ დაყენება (50%)
        'GBGreeting.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        GBGreeting.BackColor = Color.FromArgb(220, Color.White)
        GBNow.BackColor = Color.FromArgb(220, Color.White)
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Dim now = DateTime.Now
        LTime.Text = now.ToString("HH:mm:ss")
        LDate.Text = $"{now.Day} {now.ToString("MMMM", New CultureInfo("ka"))} {now.Year} წელი"
        LWeekDay.Text = now.ToString("dddd", New CultureInfo("ka"))
    End Sub
End Class
