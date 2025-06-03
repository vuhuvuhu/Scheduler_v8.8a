' ===========================================
' 📄 UserControls/UC_TherapistReport.vb
' -------------------------------------------
' თერაპევტის ანგარიშის UserControl
' შექმნილია ხელით, Designer.vb-ს გარეშე
' 
' 🎯 ხელმისაწვდომია მხოლოდ როლებისთვის: 1 (ადმინი), 2 (მენეჯერი)
' ===========================================
Imports System.ComponentModel
Imports Scheduler_v8_8a.Services
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models
Imports System.Text
Imports Scheduler_v8._8a.Scheduler_v8_8a.Services

''' <summary>
''' თერაპევტის ანგარიშის UserControl
''' საშუალებას აძლევს მენეჯერებს და ადმინისტრატორებს მოძებნონ თერაპევტი 
''' და ნახონ მისი სესიების დეტალური ანგარიში და შემოსავლების ანალიზი
''' </summary>
Public Class UC_TherapistReport


    Private Sub UC_TherapistReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class