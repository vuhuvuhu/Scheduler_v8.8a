' ===========================================
' 📄 Services/SimplePDFExportService.vb
' -------------------------------------------
' მარტივი PDF ექსპორტის სერვისი დამოკიდებულებების გარეშე
' CSV ფაილის შექმნისა და ყველა პლატფორმისთვის მუშაობის გარანტიით
' ===========================================
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' მარტივი ექსპორტის სერვისი - CSV და HTML ფორმატებში
    ''' PDF ალტერნატივა დამატებითი ბიბლიოთეკების გარეშე
    ''' </summary>
    Public Class SimplePDFExportService
        Implements IDisposable

        ' მთავარი DataGridView
        Private ReadOnly dataGridView As DataGridView

        ' მონიშნული სვეტების სია
        Private selectedColumns As List(Of DataGridViewColumn)

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dgv">ექსპორტისთვის DataGridView</param>
        Public Sub New(dgv As DataGridView)
            dataGridView = dgv
            selectedColumns = New List(Of DataGridViewColumn)
        End Sub

        ''' <summary>
        ''' 🔧 სვეტების მონიშვნის დიალოგის ჩვენება
        ''' </summary>
        Public Function ShowColumnSelectionDialog() As DialogResult
            Try
                Debug.WriteLine("SimplePDFExportService: სვეტების მონიშვნის დიალოგი")

                Using columnDialog As New ColumnSelectionForm(dataGridView)
                    Dim result As DialogResult = columnDialog.ShowDialog()

                    If result = DialogResult.OK Then
                        selectedColumns = columnDialog.GetSelectedColumns()
                        Debug.WriteLine($"SimplePDFExportService: მონიშნულია {selectedColumns.Count} სვეტი")
                    End If

                    Return result
                End Using

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ShowColumnSelectionDialog შეცდომა: {ex.Message}")
                Return DialogResult.Cancel
            End Try
        End Function

        ''' <summary>
        ''' 🔧 სრული ექსპორტის პროცესი - ფორმატის არჩევით
        ''' </summary>
        Public Sub ShowFullExportDialog()
            Try
                Debug.WriteLine("SimplePDFExportService: სრული ექსპორტის პროცესი")

                If dataGridView Is Nothing OrElse dataGridView.Rows.Count = 0 Then
                    MessageBox.Show("ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' ფორმატის არჩევა
                Dim formatResult As DialogResult = MessageBox.Show(
                    "აირჩიეთ ექსპორტის ფორმატი:" & Environment.NewLine & Environment.NewLine &
                    "დიახ - HTML ფაილი (ბრაუზერში სანახავად და ბეჭდვისთვის)" & Environment.NewLine &
                    "არა - CSV ფაილი (Excel-ისთვის)" & Environment.NewLine &
                    "გაუქმება - ოპერაციის შეწყვეტა",
                    "ექსპორტის ფორმატის არჩევა",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question)

                Select Case formatResult
                    Case DialogResult.Yes
                        ExportToHTML()
                    Case DialogResult.No
                        ExportToCSV()
                    Case DialogResult.Cancel
                        Debug.WriteLine("SimplePDFExportService: ექსპორტი გაუქმებულია")
                End Select

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ShowFullExportDialog შეცდომა: {ex.Message}")
                MessageBox.Show($"ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' HTML ფაილად ექსპორტი (ბეჭდვადი და ლამაზი)
        ''' </summary>
        Private Sub ExportToHTML()
            Try
                Debug.WriteLine("SimplePDFExportService: HTML ექსპორტი")

                ' 1. სვეტების მონიშვნა
                If ShowColumnSelectionDialog() <> DialogResult.OK Then
                    Debug.WriteLine("SimplePDFExportService: სვეტების მონიშვნა გაუქმებულია")
                    Return
                End If

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ექსპორტისთვის", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 2. ფაილის ადგილის არჩევა
                Using saveDialog As New SaveFileDialog()
                    saveDialog.Filter = "HTML ფაილები (*.html)|*.html"
                    saveDialog.Title = "HTML ფაილის შენახვა"
                    saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.html"

                    If saveDialog.ShowDialog() = DialogResult.OK Then
                        CreateHTMLFile(saveDialog.FileName)
                    Else
                        Debug.WriteLine("SimplePDFExportService: HTML ექსპორტი გაუქმებულია")
                    End If
                End Using

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ExportToHTML შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' HTML ფაილის შექმნა
        ''' </summary>
        Private Sub CreateHTMLFile(filePath As String)
            Try
                Debug.WriteLine($"SimplePDFExportService: HTML ფაილის შექმნა - {filePath}")

                Dim html As New StringBuilder()

                ' HTML დოკუმენტის საწყისი ნაწილი
                html.AppendLine("<!DOCTYPE html>")
                html.AppendLine("<html lang=""ka"">")
                html.AppendLine("<head>")
                html.AppendLine("    <meta charset=""UTF-8"">")
                html.AppendLine("    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">")
                html.AppendLine("    <title>განრიგის მონაცემები</title>")
                html.AppendLine("    <style>")
                html.AppendLine("        @media print { .no-print { display: none; } }")
                html.AppendLine("        body { font-family: 'Sylfaen', 'DejaVu Sans', sans-serif; margin: 20px; background-color: #f9f9f9; }")
                html.AppendLine("        .container { background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }")
                html.AppendLine("        h1 { text-align: center; color: #2c3e50; margin-bottom: 30px; }")
                html.AppendLine("        .info-bar { display: flex; justify-content: space-between; margin-bottom: 20px; font-size: 14px; color: #666; }")
                html.AppendLine("        table { width: 100%; border-collapse: collapse; margin-top: 20px; }")
                html.AppendLine("        th, td { padding: 12px 8px; text-align: left; border: 1px solid #ddd; }")
                html.AppendLine("        th { background-color: #f8f9fa; font-weight: bold; color: #2c3e50; }")
                html.AppendLine("        tr:nth-child(even) { background-color: #f8f9fa; }")
                html.AppendLine("        tr:hover { background-color: #e3f2fd; }")
                html.AppendLine("        .btn-container { text-align: center; margin: 20px 0; }")
                html.AppendLine("        .btn { background-color: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 0 10px; }")
                html.AppendLine("        .btn:hover { background-color: #0056b3; }")
                html.AppendLine("        .footer { text-align: center; margin-top: 30px; font-size: 12px; color: #666; }")
                html.AppendLine("    </style>")
                html.AppendLine("</head>")
                html.AppendLine("<body>")
                html.AppendLine("    <div class=""container"">")

                ' სათაური
                html.AppendLine("        <h1>განრიგის მონაცემები</h1>")

                ' ინფორმაციის ზოლი
                html.AppendLine("        <div class=""info-bar"">")
                html.AppendLine($"            <span>თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}</span>")
                html.AppendLine($"            <span>სვეტები: {selectedColumns.Count}</span>")
                html.AppendLine($"            <span>ჩანაწერები: {dataGridView.Rows.Count}</span>")
                html.AppendLine("        </div>")

                ' ღილაკები (მხოლოდ ეკრანზე ჩვენებისთვის)
                html.AppendLine("        <div class=""btn-container no-print"">")
                html.AppendLine("            <button class=""btn"" onclick=""window.print()"">🖨️ ბეჭდვა</button>")
                html.AppendLine("            <button class=""btn"" onclick=""window.close()"">❌ დახურვა</button>")
                html.AppendLine("        </div>")

                ' ცხრილი
                html.AppendLine("        <table>")

                ' სათაურები
                html.AppendLine("            <thead>")
                html.AppendLine("                <tr>")
                For Each column In selectedColumns
                    html.AppendLine($"                    <th>{column.HeaderText}</th>")
                Next
                html.AppendLine("                </tr>")
                html.AppendLine("            </thead>")

                ' მონაცემები
                html.AppendLine("            <tbody>")
                For rowIndex As Integer = 0 To dataGridView.Rows.Count - 1
                    Dim row As DataGridViewRow = dataGridView.Rows(rowIndex)
                    html.AppendLine("                <tr>")

                    For Each column In selectedColumns
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        ' HTML სპეციალური სიმბოლოების escape - მარტივი მეთოდით
                        cellValue = EscapeHtmlCharacters(cellValue)
                        html.AppendLine($"                    <td>{cellValue}</td>")
                    Next

                    html.AppendLine("                </tr>")
                Next
                html.AppendLine("            </tbody>")
                html.AppendLine("        </table>")

                ' ქვედა ინფორმაცია
                html.AppendLine("        <div class=""footer"">")
                html.AppendLine($"            <p>შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm} | სულ ჩანაწერები: {dataGridView.Rows.Count} | მონიშნული სვეტები: {selectedColumns.Count}</p>")
                html.AppendLine("            <p>Scheduler v8.8a - განრიგის მართვის სისტემა</p>")
                html.AppendLine("        </div>")

                html.AppendLine("    </div>")
                html.AppendLine("</body>")
                html.AppendLine("</html>")

                ' ფაილის ჩაწერა
                File.WriteAllText(filePath, html.ToString(), Encoding.UTF8)

                Debug.WriteLine("SimplePDFExportService: HTML ფაილი წარმატებით შეიქმნა")
                MessageBox.Show($"HTML ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}", "წარმატება",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' ფაილის გახსნის შეთავაზება
                Dim openResult As DialogResult = MessageBox.Show(
                    "გსურთ HTML ფაილის გახსნა ბრაუზერში?",
                    "ფაილის გახსნა",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If openResult = DialogResult.Yes Then
                    Process.Start(filePath)
                End If

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: CreateHTMLFile შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' CSV ფაილად ექსპორტი
        ''' </summary>
        Private Sub ExportToCSV()
            Try
                Debug.WriteLine("SimplePDFExportService: CSV ექსპორტი")

                ' 1. სვეტების მონიშვნა
                If ShowColumnSelectionDialog() <> DialogResult.OK Then
                    Debug.WriteLine("SimplePDFExportService: სვეტების მონიშვნა გაუქმებულია")
                    Return
                End If

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ექსპორტისთვის", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 2. ფაილის ადგილის არჩევა
                Using saveDialog As New SaveFileDialog()
                    saveDialog.Filter = "CSV ფაილები (*.csv)|*.csv"
                    saveDialog.Title = "CSV ფაილის შენახვა"
                    saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.csv"

                    If saveDialog.ShowDialog() = DialogResult.OK Then
                        CreateCSVFile(saveDialog.FileName)
                    Else
                        Debug.WriteLine("SimplePDFExportService: CSV ექსპორტი გაუქმებულია")
                    End If
                End Using

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ExportToCSV შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' CSV ფაილის შექმნა
        ''' </summary>
        Private Sub CreateCSVFile(filePath As String)
            Try
                Debug.WriteLine($"SimplePDFExportService: CSV ფაილის შექმნა - {filePath}")

                Dim csv As New StringBuilder()

                ' BOM (Byte Order Mark) UTF-8-სთვის
                ' ეს გაუმჯობესებს ქართული ტექსტის ჩვენებას Excel-ში
                Dim utf8WithBom As New UTF8Encoding(True)

                ' სათაურები
                Dim headers As New List(Of String)
                For Each column In selectedColumns
                    headers.Add(EscapeCSVField(column.HeaderText))
                Next
                csv.AppendLine(String.Join(",", headers))

                ' მონაცემები
                For rowIndex As Integer = 0 To dataGridView.Rows.Count - 1
                    Dim row As DataGridViewRow = dataGridView.Rows(rowIndex)
                    Dim rowData As New List(Of String)

                    For Each column In selectedColumns
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        rowData.Add(EscapeCSVField(cellValue))
                    Next

                    csv.AppendLine(String.Join(",", rowData))
                Next

                ' ფაილის ჩაწერა
                File.WriteAllText(filePath, csv.ToString(), utf8WithBom)

                Debug.WriteLine("SimplePDFExportService: CSV ფაილი წარმატებით შეიქმნა")
                MessageBox.Show($"CSV ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}" & Environment.NewLine & Environment.NewLine &
                               "ფაილი შეგიძლიათ გახსნათ Microsoft Excel-ში ან LibreOffice Calc-ში", "წარმატება",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' ფაილის გახსნის შეთავაზება
                Dim openResult As DialogResult = MessageBox.Show(
                    "გსურთ CSV ფაილის გახსნა?",
                    "ფაილის გახსნა",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If openResult = DialogResult.Yes Then
                    Process.Start(filePath)
                End If

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: CreateCSVFile შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' HTML სპეციალური სიმბოლოების escape
        ''' </summary>
        Private Function EscapeHtmlCharacters(text As String) As String
            Try
                If String.IsNullOrEmpty(text) Then
                    Return ""
                End If

                ' HTML სპეციალური სიმბოლოების ჩანაცვლება
                text = text.Replace("&", "&amp;")   ' & ყველაზე პირველ რიგში
                text = text.Replace("<", "&lt;")    ' <
                text = text.Replace(">", "&gt;")    ' >
                text = text.Replace("""", "&quot;") ' "
                text = text.Replace("'", "&#39;")   ' '

                Return text

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: EscapeHtmlCharacters შეცდომა: {ex.Message}")
                Return text
            End Try
        End Function

        ''' <summary>
        ''' CSV ველის escape (კომას, ახალ ხაზს და ციტატებს ეხმიანება)
        ''' </summary>
        Private Function EscapeCSVField(field As String) As String
            Try
                If String.IsNullOrEmpty(field) Then
                    Return ""
                End If

                ' თუ შეიცავს კომას, ახალ ხაზს ან ციტატებს
                If field.Contains(",") OrElse field.Contains(vbCrLf) OrElse field.Contains(vbLf) OrElse field.Contains("""") Then
                    ' ციტატების გადვოება
                    field = field.Replace("""", """""")
                    ' მთლიანი ველის ციტატებში ჩასმა
                    field = """" & field & """"
                End If

                Return field

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: EscapeCSVField შეცდომა: {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' სწრაფი HTML ექსპორტი - ყველა ხილული სვეტით
        ''' </summary>
        Public Sub ExportSimpleHTML(filePath As String)
            Try
                Debug.WriteLine("SimplePDFExportService: სწრაფი HTML ექსპორტი")

                ' ყველა ხილული სვეტის მონიშვნა (რედაქტირების ღილაკის გარდა)
                selectedColumns.Clear()
                For Each column As DataGridViewColumn In dataGridView.Columns
                    If column.Visible AndAlso column.Name <> "Edit" Then
                        selectedColumns.Add(column)
                    End If
                Next

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("ექსპორტისთვის ხილული სვეტები არ მოიძებნა", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' HTML ფაილის შექმნა
                CreateHTMLFile(filePath)

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ExportSimpleHTML შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' სწრაფი CSV ექსპორტი - ყველა ხილული სვეტით
        ''' </summary>
        Public Sub ExportSimpleCSV(filePath As String)
            Try
                Debug.WriteLine("SimplePDFExportService: სწრაფი CSV ექსპორტი")

                ' ყველა ხილული სვეტის მონიშვნა (რედაქტირების ღილაკის გარდა)
                selectedColumns.Clear()
                For Each column As DataGridViewColumn In dataGridView.Columns
                    If column.Visible AndAlso column.Name <> "Edit" Then
                        selectedColumns.Add(column)
                    End If
                Next

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("ექსპორტისთვის ხილული სვეტები არ მოიძებნა", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' CSV ფაილის შექმნა
                CreateCSVFile(filePath)

            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: ExportSimpleCSV შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' რესურსების განთავისუფლება - IDisposable იმპლემენტაცია
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Debug.WriteLine("SimplePDFExportService: რესურსების განთავისუფლება")
                selectedColumns?.Clear()
            Catch ex As Exception
                Debug.WriteLine($"SimplePDFExportService: Dispose შეცდომა: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace