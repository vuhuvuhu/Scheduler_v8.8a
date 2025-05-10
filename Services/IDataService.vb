' ===========================================
' 📄 Services/IDataService.vb
' -------------------------------------------
' მონაცემთა სერვისის ინტერფეისი
' ===========================================

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' IDataService - მონაცემთა სერვისების აბსტრაქტული ინტერფეისი
    ''' გამოიყენება Google Sheets-ისა და სხვა წყაროების გამოყენების უნიფიკაციისთვის
    ''' </summary>
    Public Interface IDataService

        ''' <summary>
        ''' წაიკითხავს მონაცემებს მითითებული დიაპაზონიდან
        ''' </summary>
        ''' <param name="range">მონაცემთა დიაპაზონი (მაგ: "DB-Users!B2:C10")</param>
        ''' <returns>მონაცემთა მასივი (მწკრივები და სვეტები)</returns>
        Function GetData(range As String) As IList(Of IList(Of Object))

        ''' <summary>
        ''' ჩაამატებს მონაცემებს მითითებულ დიაპაზონში
        ''' </summary>
        ''' <param name="range">მონაცემთა დიაპაზონი (მაგ: "DB-Users!B:C")</param>
        ''' <param name="values">მონაცემები დასამატებლად</param>
        Sub AppendData(range As String, values As IList(Of Object))

        ''' <summary>
        ''' განაახლებს მონაცემებს მითითებულ დიაპაზონში
        ''' </summary>
        ''' <param name="range">მონაცემთა დიაპაზონი (მაგ: "DB-Users!B2:C2")</param>
        ''' <param name="values">მონაცემები განახლებისთვის</param>
        Sub UpdateData(range As String, values As IList(Of Object))

        ''' <summary>
        ''' წაიკითხავს მომხმარებლის როლს ელ.ფოსტის მიხედვით
        ''' </summary>
        ''' <param name="email">მომხმარებლის ელ.ფოსტა</param>
        ''' <returns>მომხმარებლის როლი როგორც სტრიქონი</returns>
        Function GetUserRole(email As String) As String

        ''' <summary>
        ''' მიიღებს ან შექმნის მომხმარებლის ჩანაწერს
        ''' </summary>
        ''' <param name="email">მომხმარებლის ელ.ფოსტა</param>
        ''' <returns>მომხმარებლის როლი როგორც სტრიქონი</returns>
        Function GetOrCreateUserRole(email As String) As String
        ''' <summary>
        ''' წამოიღებს ყველა სესიას დღევანდელი სტატისტიკისთვის
        ''' </summary>
        ''' <returns>სესიების მოდელების სია</returns>
        Function GetTodaySessions() As List(Of Models.SessionModel)
        ''' <summary>
        ''' წამოიღებს მოახლოებულ დაბადების დღეებს (X დღეში)
        ''' </summary>
        ''' <param name="days">რამდენი დღის განმავლობაში (ნაგულისხმები: 7)</param>
        ''' <returns>დაბადების დღეების მოდელების სია</returns>
        Function GetUpcomingBirthdays(Optional days As Integer = 7) As List(Of Models.BirthdayModel)

        ''' <summary>
        ''' წამოიღებს მოლოდინში არსებულ სესიებს
        ''' </summary>
        ''' <returns>სესიების მოდელების სია</returns>
        Function GetPendingSessions() As List(Of Models.SessionModel)

        ''' <summary>
        ''' წამოიღებს აქტიურ დავალებებს
        ''' </summary>
        ''' <returns>დავალებების მოდელების სია</returns>
        Function GetActiveTasks() As List(Of Models.TaskModel)
        ''' <summary>
        ''' წამოიღებს ვადაგადაცილებულ სესიებს (სტატუსით "დაგეგმილი", მაგრამ თარიღი უკვე გასულია)
        ''' </summary>
        ''' <returns>ვადაგადაცილებული სესიების მოდელების სია</returns>
        Function GetOverdueSessions() As List(Of Models.SessionModel)
    End Interface

End Namespace