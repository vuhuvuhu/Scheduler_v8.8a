' ===========================================
' 📄 Utils/CalendarUtils.vb
' -------------------------------------------
' კალენდართან დაკავშირებული დამხმარე ფუნქციები
' ===========================================
Imports System.Drawing
Imports System.Drawing2D
Imports Scheduler_v8_8a.Models
Imports Scheduler_v8._8a.Scheduler_v8_8a.Models

Namespace Scheduler_v8_8a.Utils
    ''' <summary>
    ''' კალენდრის დამხმარე ფუნქციები
    ''' </summary>
    Public Class CalendarUtils
        ''' <summary>
        ''' გრიდის პოზიციის გამოთვლა pixel კოორდინატებიდან
        ''' </summary>
        Public Shared Function CalculateGridPosition(pixelLocation As Point, rowHeight As Integer, columnWidth As Integer, spacesCount As Integer, timeIntervalsCount As Integer) As (spaceIndex As Integer, timeIndex As Integer)
            Try
                ' გამოვთვალოთ ინდექსები (მარჯინების გათვალისწინებით)
                Dim spaceIndex As Integer = pixelLocation.X \ columnWidth
                Dim timeIndex As Integer = pixelLocation.Y \ rowHeight

                ' შევამოწმოთ საზღვრები
                If spaceIndex < 0 Then spaceIndex = 0
                If spaceIndex >= spacesCount Then spaceIndex = spacesCount - 1
                If timeIndex < 0 Then timeIndex = 0
                If timeIndex >= timeIntervalsCount Then timeIndex = timeIntervalsCount - 1

                Debug.WriteLine($"CalculateGridPosition: Pixel({pixelLocation.X}, {pixelLocation.Y}) -> Grid({spaceIndex}, {timeIndex})")
                Return (spaceIndex, timeIndex)

            Catch ex As Exception
                Debug.WriteLine($"CalculateGridPosition: შეცდომა - {ex.Message}")
                Return (0, 0)
            End Try
        End Function

        ''' <summary>
        ''' თერაპევტების გრიდის პოზიციის გამოთვლა pixel კოორდინატებიდან
        ''' </summary>
        Public Shared Function CalculateTherapistGridPosition(pixelLocation As Point, rowHeight As Integer, columnWidth As Integer, therapistsCount As Integer, timeIntervalsCount As Integer) As (therapistIndex As Integer, timeIndex As Integer)
            Try
                ' გამოვთვალოთ ინდექსები (მარჯინების გათვალისწინებით)
                Dim therapistIndex As Integer = pixelLocation.X \ columnWidth
                Dim timeIndex As Integer = pixelLocation.Y \ rowHeight

                ' შევამოწმოთ საზღვრები
                If therapistIndex < 0 Then therapistIndex = 0
                If therapistIndex >= therapistsCount Then therapistIndex = therapistsCount - 1
                If timeIndex < 0 Then timeIndex = 0
                If timeIndex >= timeIntervalsCount Then timeIndex = timeIntervalsCount - 1

                Debug.WriteLine($"CalculateTherapistGridPosition: Pixel({pixelLocation.X}, {pixelLocation.Y}) -> Grid({therapistIndex}, {timeIndex})")
                Return (therapistIndex, timeIndex)

            Catch ex As Exception
                Debug.WriteLine($"CalculateTherapistGridPosition: შეცდომა - {ex.Message}")
                Return (0, 0)
            End Try
        End Function

        ''' <summary>
        ''' უახლოესი დროის ინტერვალის ინდექსის გამოთვლა
        ''' </summary>
        Public Shared Function FindNearestTimeIndex(sessionTime As TimeSpan, timeIntervals As List(Of DateTime)) As Integer
            Try
                ' უახლოესი დროის ინტერვალის ძებნა
                Dim timeIndex As Integer = -1
                Dim minDifference As TimeSpan = TimeSpan.MaxValue

                For i As Integer = 0 To timeIntervals.Count - 1
                    Dim intervalTime As TimeSpan = timeIntervals(i).TimeOfDay
                    Dim difference As TimeSpan = If(sessionTime >= intervalTime,
                                          sessionTime - intervalTime,
                                          intervalTime - sessionTime)

                    If difference < minDifference Then
                        minDifference = difference
                        timeIndex = i
                    End If
                Next

                ' ალტერნატიული მეთოდი თუ პირველმა ვერ იპოვა
                If timeIndex < 0 OrElse minDifference.TotalMinutes > 15 Then
                    Dim startTime As TimeSpan = timeIntervals(0).TimeOfDay
                    Dim elapsedMinutes As Double = (sessionTime - startTime).TotalMinutes
                    timeIndex = CInt(Math.Round(elapsedMinutes / 30))
                    If timeIndex < 0 Then timeIndex = 0
                    If timeIndex >= timeIntervals.Count Then timeIndex = timeIntervals.Count - 1
                End If

                Return timeIndex
            Catch ex As Exception
                Debug.WriteLine($"FindNearestTimeIndex: შეცდომა - {ex.Message}")
                Return 0
            End Try
        End Function
    End Class
End Namespace
