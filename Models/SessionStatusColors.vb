' ===========================================
' 📄 Models/SessionStatusColors.vb
' -------------------------------------------
' სესიის სტატუსების ფერების კლასი
' ===========================================
Imports System.Drawing

Namespace Scheduler_v8_8a.Models
    ''' <summary>
    ''' სტატუსების ფერების კონფიგურაცია - ცენტრალური ადგილი ყველა სტატუსის ფერისთვის
    ''' გამოიყენება როგორც კალენდარში, ასევე სხვა ადგილებში სადაც სესიის ბარათები ჩანს
    ''' </summary>
    Public Class SessionStatusColors
        ''' <summary>
        ''' სტატუსის მიხედვით ფერის დაბრუნება
        ''' </summary>
        ''' <param name="status">სესიის სტატუსი</param>
        ''' <param name="sessionDateTime">სესიის თარიღი და დრო (ვადაგადაცილების შემოწმებისთვის)</param>
        ''' <returns>შესაბამისი ფერი</returns>
        Public Shared Function GetStatusColor(status As String, sessionDateTime As DateTime) As Color
            ' სტატუსის დანორმალიზება
            Dim normalizedStatus As String = status.Trim().ToLower()

            ' დაგეგმილი სესიებისთვის ვამოწმებთ ვადაგადაცილებას
            If normalizedStatus = "დაგეგმილი" Then
                ' თუ ვადა გავიდა - ვარდისფერი
                If sessionDateTime < DateTime.Now Then
                    Return OverdueColor
                Else
                    ' თუ ვადა ჯერ არ მოსულა - თეთრი
                    Return PlannedColor
                End If
            End If

            ' დანარჩენი სტატუსები
            Select Case normalizedStatus
                Case "შესრულებული"
                    Return CompletedColor
                Case "გაუქმებული", "გაუქმება"
                    Return CancelledColor
                Case "პროგრამით გატარება"
                    Return AutoProcessedColor
                Case "აღდგენა"
                    Return RestoredColor
                Case "გაცდენა არასაპატიო"
                    Return MissedUnexcusedColor
                Case "გაცდენა საპატიო"
                    Return MissedExcusedColor
                Case "შესრულების პროცესში"
                    Return InProgressColor
                Case Else
                    ' უცნობი სტატუსისთვის ნაგულისხმევი ფერი
                    Return DefaultColor
            End Select
        End Function

        ''' <summary>
        ''' ღია მწვანე - შესრულებული სესიებისთვის
        ''' </summary>
        Public Shared ReadOnly Property CompletedColor As Color
            Get
                Return Color.FromArgb(200, 255, 200) ' ღია მწვანე
            End Get
        End Property

        ''' <summary>
        ''' თეთრი - დაგეგმილი სესიებისთვის (ვადა ჯერ არ მოსულა)
        ''' </summary>
        Public Shared ReadOnly Property PlannedColor As Color
            Get
                Return Color.FromArgb(255, 255, 255) ' თეთრი
            End Get
        End Property

        ''' <summary>
        ''' ვარდისფერი - ვადაგადაცილებული დაგეგმილი სესიებისთვის
        ''' </summary>
        Public Shared ReadOnly Property OverdueColor As Color
            Get
                Return Color.FromArgb(255, 182, 193) ' ღია ვარდისფერი
            End Get
        End Property

        ''' <summary>
        ''' წითელი - გაუქმებული სესიებისთვის
        ''' </summary>
        Public Shared ReadOnly Property CancelledColor As Color
            Get
                Return Color.FromArgb(255, 150, 150) ' წითელი
            End Get
        End Property

        ''' <summary>
        ''' ნაცრისფერი - პროგრამით გატარებული სესიებისთვის
        ''' </summary>
        Public Shared ReadOnly Property AutoProcessedColor As Color
            Get
                Return Color.FromArgb(220, 220, 220) ' ნაცრისფერი
            End Get
        End Property

        ''' <summary>
        ''' უფრო ღია მწვანე - აღდგენილი სესიებისთვის
        ''' </summary>
        Public Shared ReadOnly Property RestoredColor As Color
            Get
                Return Color.FromArgb(170, 255, 170) ' უფრო ღია მწვანე
            End Get
        End Property

        ''' <summary>
        ''' ღია იასამნისფერი - არასაპატიო გაცდენისთვის
        ''' </summary>
        Public Shared ReadOnly Property MissedUnexcusedColor As Color
            Get
                Return Color.FromArgb(230, 230, 255) ' ღია იასამნისფერი
            End Get
        End Property

        ''' <summary>
        ''' ყვითელი - საპატიო გაცდენისთვის
        ''' </summary>
        Public Shared ReadOnly Property MissedExcusedColor As Color
            Get
                Return Color.FromArgb(255, 255, 200) ' ღია ყვითელი
            End Get
        End Property

        ''' <summary>
        ''' ღია ლურჯი - შესრულების პროცესშია
        ''' </summary>
        Public Shared ReadOnly Property InProgressColor As Color
            Get
                Return Color.FromArgb(200, 230, 255) ' ღია ლურჯი
            End Get
        End Property

        ''' <summary>
        ''' ნაგულისხმევი ფერი უცნობი სტატუსებისთვის
        ''' </summary>
        Public Shared ReadOnly Property DefaultColor As Color
            Get
                Return Color.FromArgb(240, 240, 240) ' ღია ნაცრისფერი
            End Get
        End Property

        ''' <summary>
        ''' ყველა შესაძლო სტატუსის სია მათი ფერებით
        ''' გამოიყენება ლეგენდისა და რეპორტებისთვის
        ''' </summary>
        Public Shared Function GetAllStatusColorsForLegend() As Dictionary(Of String, Color)
            Dim statusColors As New Dictionary(Of String, Color)

            statusColors.Add("შესრულებული", CompletedColor)
            statusColors.Add("დაგეგმილი (ვადაში)", PlannedColor)
            statusColors.Add("დაგეგმილი (ვადაგადაცილებული)", OverdueColor)
            statusColors.Add("გაუქმებული", CancelledColor)
            statusColors.Add("პროგრამით გატარება", AutoProcessedColor)
            statusColors.Add("აღდგენა", RestoredColor)
            statusColors.Add("გაცდენა არასაპატიო", MissedUnexcusedColor)
            statusColors.Add("გაცდენა საპატიო", MissedExcusedColor)
            statusColors.Add("შესრულების პროცესში", InProgressColor)

            Return statusColors
        End Function

        ''' <summary>
        ''' ბარათის ჩარჩოს ფერის დაბრუნება სტატუსის მიხედვით
        ''' ზოგიერთი ბარათისთვის შეიძლება სხვადასხვა ჩარჩო გვჭირდებოდეს
        ''' </summary>
        Public Shared Function GetStatusBorderColor(status As String, sessionDateTime As DateTime) As Color
            ' ძირითადი ფერი
            Dim mainColor As Color = GetStatusColor(status, sessionDateTime)

            ' ჩარჩო ოდნავ უფრო მუქია
            Return Color.FromArgb(
                Math.Max(0, mainColor.R - 30),
                Math.Max(0, mainColor.G - 30),
                Math.Max(0, mainColor.B - 30)
            )
        End Function
    End Class
End Namespace
