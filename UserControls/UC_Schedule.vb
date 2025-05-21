' ===========================================
' 📄 UserControls/UC_Schedule.vb
' -------------------------------------------
' განრიგის UserControl - აჩვენებს და მართავს სესიების განრიგს
' ===========================================
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

Public Class UC_Schedule
    Inherits UserControl

    ' მონაცემთა სერვისი
    Private dataService As IDataService = Nothing

    ''' <summary>
    ''' მონაცემთა სერვისის დაყენება
    ''' </summary>
    ''' <param name="service">მონაცემთა სერვისი (IDataService)</param>
    Public Sub SetDataService(service As IDataService)
        dataService = service
        Debug.WriteLine("UC_Schedule.SetDataService: მონაცემთა სერვისი დაყენებულია")

        ' როდესაც რეალურად დავიწყებთ მონაცემებთან მუშაობას,
        ' აქ შეგვიძლია დავამატოთ მონაცემების ჩატვირთვის ლოგიკა
    End Sub

    Private Sub UC_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pnlFilter.BackColor = Color.FromArgb(200, Color.White)
        GBSumInf.BackColor = Color.FromArgb(200, Color.White)
        GBSumFin.BackColor = Color.FromArgb(200, Color.White)
    End Sub
End Class
