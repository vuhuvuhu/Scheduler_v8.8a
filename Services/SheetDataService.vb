' ===========================================
' 📄 Services/SheetDataService.vb
' -------------------------------------------
' პასუხისმგებელი Google Sheets–დან მონაცემთა წამოღებასა და ჩანაწერების დამატებაზე
' ===========================================
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Services
Imports Google.Apis.Auth.OAuth2

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' SheetDataService ახორციელებს მონაცემთა წაკითხვასა და განახლებას
    ''' Google Sheets–ის "DB-Users" ფურცელზე (B სვეტი = ელ.ფოსტა, C სვეტი = როლი).
    ''' </summary>
    Public Class SheetDataService

        Private ReadOnly sheetsService As SheetsService
        Private ReadOnly spreadsheetId As String
        Private Const usersRange As String = "DB-Users!B2:C"
        Private Const appendRange As String = "DB-Users!B:C"
        Private Const defaultRole As String = "6"

        ''' <summary>
        ''' კონსტრუქტორი: აგებს SheetsService–ს OAUTHCredential–ით და spreadsheetId–თ.
        ''' </summary>
        ''' <param name="credential">Google OAuth2 UserCredential</param>
        ''' <param name="spreadsheetId">Spreadsheet-ის ID</param>
        Public Sub New(credential As UserCredential, spreadsheetId As String)
            Me.sheetsService = New SheetsService(New BaseClientService.Initializer() With {
                .HttpClientInitializer = credential,
                .ApplicationName = "Scheduler_v8.8a"
            })
            Me.spreadsheetId = spreadsheetId
        End Sub

        ''' <summary>
        ''' აბრუნებს მომხმარებლის როლს, ან ხსნის ახალ ჩანაწერს თუ არ არსებობს (role=6).
        ''' </summary>
        ''' <param name="userEmail">მომხმარებლის ელ.ფოსტა</param>
        ''' <returns>C სვეტში მდებარე როლი როგორც string</returns>
        Public Function GetOrCreateUserRole(userEmail As String) As String
            ' 1) წაკითხვა
            Dim getReq = sheetsService.Spreadsheets.Values.Get(spreadsheetId, usersRange)
            Dim rows As IList(Of IList(Of Object)) = getReq.Execute().Values
            If rows IsNot Nothing Then
                For Each row As IList(Of Object) In rows
                    If row.Count >= 2 AndAlso String.Equals(row(0).ToString(), userEmail, StringComparison.OrdinalIgnoreCase) Then
                        Return row(1).ToString()
                    End If
                Next
            End If
            ' 2) თუ არ მოიძებნა, დავამატოთ ახალი role=6 და დავუბრუნოთ defaultRole
            Dim newRow As IList(Of Object) = New List(Of Object) From {userEmail, defaultRole}
            Dim valueRange As New ValueRange With {.Values = New List(Of IList(Of Object)) From {newRow}}
            Dim appendReq = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, appendRange)
            appendReq.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
            appendReq.Execute()
            Return defaultRole
        End Function

        ''' <summary>
        ''' Custom წაკითხვა ნებისმიერი დიაპაზონიდან (უნდა იყენებოდეს სხვა(sheet) ჟურნალებისთვის)
        ''' </summary>
        Public Function ReadRange(range As String) As IList(Of IList(Of Object))
            Dim req = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range)
            Return req.Execute().Values
        End Function

        ''' <summary>
        ''' ჩანაწერი დამატება ნებისმიერი დიაპაზონში
        ''' </summary>
        Public Sub AppendRow(range As String, values As IList(Of Object))
            Dim vr As New ValueRange With {.Values = New List(Of IList(Of Object)) From {values}}
            Dim appendReq = sheetsService.Spreadsheets.Values.Append(vr, spreadsheetId, range)
            appendReq.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
            appendReq.Execute()
        End Sub
    End Class
End Namespace
