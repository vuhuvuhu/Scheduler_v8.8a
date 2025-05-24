' ===========================================
' 📄 Services/DataGridViewPDFExportService.vb
' -------------------------------------------
' DataGridView-ის მონაცემების PDF-ად ექსპორტის სერვისი
' iTextSharp ბიბლიოთეკის გამოყენებით
' ქართული ფონტებისა და სვეტების მონიშვნის მხარდაჭერით
' ===========================================
Imports System.IO
Imports System.Windows.Forms
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' DataGridView-ის PDF ექსპორტის სერვისი
    ''' უზრუნველყოფს ქართული ტექსტის სწორ PDF-ად ექსპორტს
    ''' </summary>
    Public Class DataGridViewPDFExportService
        Implements IDisposable

        ' მთავარი DataGridView
        Private ReadOnly dataGridView As DataGridView

        ' PDF მონაცემები
        Private document As Document
        Private writer As PdfWriter
        Private selectedColumns As List(Of DataGridViewColumn)

        ' ფონტები PDF-სთვის
        Private titleFont As iTextSharp.text.Font
        Private headerFont As iTextSharp.text.Font
        Private dataFont As iTextSharp.text.Font
        Private footerFont As iTextSharp.text.Font

        ' ფერები
        Private headerColor As BaseColor
        Private alternateRowColor As BaseColor

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dgv">ექსპორტისთვის DataGridView</param>
        Public Sub New(dgv As DataGridView)
            dataGridView = dgv
            selectedColumns = New List(Of DataGridViewColumn)
            InitializeFonts()
        End Sub

        ''' <summary>
        ''' ფონტების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeFonts()
            Try
                ' PDF ფონტების შექმნა (Arial Unicode MS ქართული ტექსტისთვის)
                ' თუ Arial Unicode MS არ არის ხელმისაწვდომი, გამოვიყენოთ ნაგულისხმევი
                Dim baseFont As BaseFont = Nothing

                Try
                    ' ვცდილობთ ქართული ფონტის ჩატვირთვას
                    baseFont = baseFont.CreateFont("c:/windows/fonts/arialuni.ttf", baseFont.IDENTITY_H, baseFont.EMBEDDED)
                Catch
                    Try
                        ' ალტერნატივა - Sylfaen
                        baseFont = baseFont.CreateFont("c:/windows/fonts/sylfaen.ttf", baseFont.IDENTITY_H, baseFont.EMBEDDED)
                    Catch
                        ' უკანასკნელი ალტერნატივა - ნაგულისხმევი
                        baseFont = baseFont.CreateFont(baseFont.HELVETICA, baseFont.CP1252, baseFont.NOT_EMBEDDED)
                    End Try
                End Try

                ' ფონტების შექმნა
                titleFont = New iTextSharp.text.Font(baseFont, 16, iTextSharp.text.Font.BOLD)
                headerFont = New iTextSharp.text.Font(baseFont, 11, iTextSharp.text.Font.BOLD)
                dataFont = New iTextSharp.text.Font(baseFont, 9, iTextSharp.text.Font.NORMAL)
                footerFont = New iTextSharp.text.Font(baseFont, 8, iTextSharp.text.Font.ITALIC)

                ' ფერების ინიციალიზაცია
                headerColor = New BaseColor(220, 220, 220)
                alternateRowColor = New BaseColor(245, 245, 250)

                Debug.WriteLine("DataGridViewPDFExportService: ფონტები ინიციალიზებულია")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: ფონტების ინიციალიზაციის შეცდომა: {ex.Message}")
                ' ნაგულისხმევი ფონტები
                titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16)
                headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11)
                dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 9)
                footerFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8)
            End Try
        End Sub

        ''' <summary>
        ''' 🔧 სვეტების მონიშვნის დიალოგის ჩვენება
        ''' </summary>
        Public Function ShowColumnSelectionDialog() As DialogResult
            Try
                Debug.WriteLine("DataGridViewPDFExportService: სვეტების მონიშვნის დიალოგი")

                Using columnDialog As New ColumnSelectionForm(dataGridView)
                    Dim result As DialogResult = columnDialog.ShowDialog()

                    If result = DialogResult.OK Then
                        selectedColumns = columnDialog.GetSelectedColumns()
                        Debug.WriteLine($"DataGridViewPDFExportService: მონიშნულია {selectedColumns.Count} სვეტი")
                    End If

                    Return result
                End Using

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: ShowColumnSelectionDialog შეცდომა: {ex.Message}")
                Return DialogResult.Cancel
            End Try
        End Function

        ''' <summary>
        ''' 🔧 სრული PDF ექსპორტის პროცესი - სვეტების მონიშვნა + ფაილის ადგილის არჩევა + PDF შექმნა
        ''' </summary>
        Public Sub ShowFullExportDialog()
            Try
                Debug.WriteLine("DataGridViewPDFExportService: სრული PDF ექსპორტის პროცესი")

                If dataGridView Is Nothing OrElse dataGridView.Rows.Count = 0 Then
                    MessageBox.Show("ექსპორტისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 1. სვეტების მონიშვნა
                If ShowColumnSelectionDialog() <> DialogResult.OK Then
                    Debug.WriteLine("DataGridViewPDFExportService: სვეტების მონიშვნა გაუქმებულია")
                    Return
                End If

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ექსპორტისთვის", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 2. ფაილის ადგილის არჩევა
                Using saveDialog As New SaveFileDialog()
                    saveDialog.Filter = "PDF ფაილები (*.pdf)|*.pdf"
                    saveDialog.Title = "PDF ფაილის შენახვა"
                    saveDialog.FileName = $"განრიგი_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"

                    If saveDialog.ShowDialog() = DialogResult.OK Then
                        ' 3. PDF-ის შექმნა
                        CreatePDF(saveDialog.FileName)
                    Else
                        Debug.WriteLine("DataGridViewPDFExportService: ფაილის შენახვის დიალოგი გაუქმებულია")
                    End If
                End Using

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: ShowFullExportDialog შეცდომა: {ex.Message}")
                MessageBox.Show($"PDF ექსპორტის შეცდომა: {ex.Message}", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' PDF ფაილის შექმნა
        ''' </summary>
        ''' <param name="filePath">ფაილის მისამართი</param>
        Private Sub CreatePDF(filePath As String)
            Try
                Debug.WriteLine($"DataGridViewPDFExportService: PDF-ის შექმნა - {filePath}")

                ' A4 ლანდშაფტ ორიენტაცია
                document = New Document(PageSize.A4.Rotate(), 30, 30, 30, 30)
                writer = PdfWriter.GetInstance(document, New FileStream(filePath, FileMode.Create))

                document.Open()

                ' სათაურის დამატება
                AddTitle()

                ' თარიღისა და ინფორმაციის დამატება
                AddDateAndInfo()

                ' ცხრილის შექმნა და მონაცემების დამატება
                CreateTable()

                ' ქვედა ინფორმაციის დამატება
                AddFooter()

                document.Close()

                Debug.WriteLine("DataGridViewPDFExportService: PDF წარმატებით შეიქმნა")
                MessageBox.Show($"PDF ფაილი წარმატებით შეიქმნა:{Environment.NewLine}{filePath}", "წარმატება",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' ფაილის გახსნის შეთავაზება
                Dim openResult As DialogResult = MessageBox.Show(
                    "გსურთ PDF ფაილის გახსნა?",
                    "ფაილის გახსნა",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If openResult = DialogResult.Yes Then
                    Process.Start(filePath)
                End If

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: CreatePDF შეცდომა: {ex.Message}")
                document?.Close()
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' სათაურის დამატება
        ''' </summary>
        Private Sub AddTitle()
            Try
                Dim title As New Paragraph("განრიგის მონაცემები", titleFont)
                title.Alignment = Element.ALIGN_CENTER
                title.SpacingAfter = 20
                document.Add(title)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: AddTitle შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' თარიღისა და ინფორმაციის დამატება
        ''' </summary>
        Private Sub AddDateAndInfo()
            Try
                Dim infoTable As New PdfPTable(3)
                infoTable.WidthPercentage = 100
                infoTable.SetWidths({1, 1, 1})

                ' მარცხენა - თარიღი
                Dim dateCell As New PdfPCell(New Phrase($"თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}", dataFont))
                dateCell.Border = Rectangle.NO_BORDER
                dateCell.HorizontalAlignment = Element.ALIGN_LEFT
                infoTable.AddCell(dateCell)

                ' შუა - მონიშნული სვეტების რაოდენობა
                Dim columnsCell As New PdfPCell(New Phrase($"სვეტები: {selectedColumns.Count}", dataFont))
                columnsCell.Border = Rectangle.NO_BORDER
                columnsCell.HorizontalAlignment = Element.ALIGN_CENTER
                infoTable.AddCell(columnsCell)

                ' მარჯვენა - ჩანაწერების რაოდენობა
                Dim recordsCell As New PdfPCell(New Phrase($"ჩანაწერები: {dataGridView.Rows.Count}", dataFont))
                recordsCell.Border = Rectangle.NO_BORDER
                recordsCell.HorizontalAlignment = Element.ALIGN_RIGHT
                infoTable.AddCell(recordsCell)

                infoTable.SpacingAfter = 15
                document.Add(infoTable)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: AddDateAndInfo შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მონაცემების ცხრილის შექმნა
        ''' </summary>
        Private Sub CreateTable()
            Try
                If selectedColumns.Count = 0 Then Return

                ' ცხრილის შექმნა
                Dim table As New PdfPTable(selectedColumns.Count)
                table.WidthPercentage = 100
                table.SpacingBefore = 10

                ' სვეტების სიგანეების გამოთვლა
                Dim columnWidths(selectedColumns.Count - 1) As Single
                For i As Integer = 0 To selectedColumns.Count - 1
                    columnWidths(i) = selectedColumns(i).Width
                Next
                table.SetWidths(columnWidths)

                ' სათაურების დამატება
                AddTableHeaders(table)

                ' მონაცემების დამატება
                AddTableData(table)

                document.Add(table)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: CreateTable შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ცხრილის სათაურების დამატება
        ''' </summary>
        Private Sub AddTableHeaders(table As PdfPTable)
            Try
                For Each column In selectedColumns
                    Dim headerCell As New PdfPCell(New Phrase(column.HeaderText, headerFont))
                    headerCell.BackgroundColor = headerColor
                    headerCell.HorizontalAlignment = Element.ALIGN_CENTER
                    headerCell.VerticalAlignment = Element.ALIGN_MIDDLE
                    headerCell.Padding = 5
                    table.AddCell(headerCell)
                Next

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: AddTableHeaders შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ცხრილის მონაცემების დამატება
        ''' </summary>
        Private Sub AddTableData(table As PdfPTable)
            Try
                For rowIndex As Integer = 0 To dataGridView.Rows.Count - 1
                    Dim row As DataGridViewRow = dataGridView.Rows(rowIndex)

                    ' ალტერნაციული ფონის ფერი
                    Dim backgroundColor As BaseColor = If(rowIndex Mod 2 = 1, alternateRowColor, BaseColor.WHITE)

                    For Each column In selectedColumns
                        ' უჯრის მნიშვნელობის მიღება
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(column.Name).Value IsNot Nothing Then
                                cellValue = row.Cells(column.Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        ' PDF უჯრის შექმნა
                        Dim dataCell As New PdfPCell(New Phrase(cellValue, dataFont))
                        dataCell.BackgroundColor = backgroundColor
                        dataCell.HorizontalAlignment = Element.ALIGN_LEFT
                        dataCell.VerticalAlignment = Element.ALIGN_MIDDLE
                        dataCell.Padding = 3

                        ' რიცხვითი მონაცემებისთვის მარჯვნივ სწორება
                        If IsNumeric(cellValue) Then
                            dataCell.HorizontalAlignment = Element.ALIGN_RIGHT
                        End If

                        table.AddCell(dataCell)
                    Next
                Next

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: AddTableData შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ქვედა ინფორმაციის დამატება
        ''' </summary>
        Private Sub AddFooter()
            Try
                Dim footerText As String = $"სულ ჩანაწერები: {dataGridView.Rows.Count} | მონიშნული სვეტები: {selectedColumns.Count} | შექმნილია: {DateTime.Now:dd.MM.yyyy HH:mm}"
                Dim footer As New Paragraph(footerText, footerFont)
                footer.Alignment = Element.ALIGN_CENTER
                footer.SpacingBefore = 20
                document.Add(footer)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: AddFooter შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' მარტივი PDF ექსპორტი - ყველა სვეტით
        ''' </summary>
        ''' <param name="filePath">ფაილის მისამართი</param>
        Public Sub ExportSimplePDF(filePath As String)
            Try
                Debug.WriteLine("DataGridViewPDFExportService: მარტივი PDF ექსპორტი")

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

                ' PDF-ის შექმნა
                CreatePDF(filePath)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: ExportSimplePDF შეცდომა: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' რესურსების განთავისუფლება - IDisposable იმპლემენტაცია
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Debug.WriteLine("DataGridViewPDFExportService: რესურსების განთავისუფლება")

                document?.Close()
                writer?.Close()

                selectedColumns?.Clear()

                Debug.WriteLine("DataGridViewPDFExportService: ყველა რესურსი განთავისუფლებულია")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPDFExportService: Dispose შეცდომა: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace