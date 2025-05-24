' ===========================================
' 📄 Services/AdvancedDataGridViewPrintService.vb
' -------------------------------------------
' გაუმჯობესებული DataGridView ბეჭდვის სერვისი
' სვეტების მონიშვნა, პრინტერის არჩევა, ლანდშაფტი (სიგრძეზე ბეჭდვა)
' ===========================================
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' გაუმჯობესებული DataGridView ბეჭდვის სერვისი
    ''' მოიცავს სვეტების მონიშვნას, პრინტერის არჩევას და ლანდშაფტ ორიენტაციას
    ''' </summary>
    Public Class AdvancedDataGridViewPrintService
        Implements IDisposable

        ' მთავარი DataGridView
        Private ReadOnly dataGridView As DataGridView

        ' ბეჭდვის კომპონენტები
        Private printDocument As PrintDocument
        Private printDialog As PrintDialog
        Private printPreviewDialog As PrintPreviewDialog
        Private pageSetupDialog As PageSetupDialog

        ' ფონტები და სტილები
        Private titleFont As Font
        Private headerFont As Font
        Private dataFont As Font
        Private footerFont As Font

        ' ფერები
        Private headerBrush As SolidBrush
        Private dataBrush As SolidBrush
        Private titleBrush As SolidBrush

        ' ბეჭდვის პარამეტრები
        Private currentRowIndex As Integer = 0
        Private currentPageNumber As Integer = 1
        Private selectedColumns As List(Of DataGridViewColumn)

        ' გვერდის პარამეტრები (ლანდშაფტისთვის)
        Private Const marginsSize As Integer = 30
        Private Const rowHeight As Integer = 20
        Private Const headerHeight As Integer = 30
        Private Const titleHeight As Integer = 40

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dgv">ბეჭდვისთვის DataGridView</param>
        Public Sub New(dgv As DataGridView)
            dataGridView = dgv
            selectedColumns = New List(Of DataGridViewColumn)
            InitializePrintComponents()
            InitializeFontsAndColors()
        End Sub

        ''' <summary>
        ''' ბეჭდვის კომპონენტების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializePrintComponents()
            Try
                ' PrintDocument-ის შექმნა
                printDocument = New PrintDocument()
                AddHandler printDocument.PrintPage, AddressOf OnPrintPage
                AddHandler printDocument.BeginPrint, AddressOf OnBeginPrint

                ' 🔧 ლანდშაფტ ორიენტაცია (სიგრძეზე ბეჭდვა)
                printDocument.DefaultPageSettings.Landscape = True

                ' PrintDialog-ის შექმნა
                printDialog = New PrintDialog()
                printDialog.Document = printDocument
                printDialog.UseEXDialog = True

                ' PrintPreviewDialog-ის შექმნა
                printPreviewDialog = New PrintPreviewDialog()
                printPreviewDialog.Document = printDocument
                printPreviewDialog.WindowState = FormWindowState.Maximized

                ' PageSetupDialog-ის შექმნა
                pageSetupDialog = New PageSetupDialog()
                pageSetupDialog.Document = printDocument

                Debug.WriteLine("AdvancedDataGridViewPrintService: ბეჭდვის კომპონენტები ინიციალიზებულია (ლანდშაფტ რეჟიმი)")

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: კომპონენტების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ფონტებისა და ფერების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeFontsAndColors()
            Try
                ' ქართული ფონტები (ლანდშაფტისთვის ოპტიმიზირებული)
                titleFont = New Font("Sylfaen", 14, FontStyle.Bold)
                headerFont = New Font("Sylfaen", 10, FontStyle.Bold)
                dataFont = New Font("Sylfaen", 8, FontStyle.Regular)
                footerFont = New Font("Sylfaen", 7, FontStyle.Italic)

                ' ფერები
                titleBrush = New SolidBrush(Color.Black)
                headerBrush = New SolidBrush(Color.DarkBlue)
                dataBrush = New SolidBrush(Color.Black)

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: ფონტების ინიციალიზაციის შეცდომა: {ex.Message}")
                ' ნაგულისხმევი ფონტები
                titleFont = New Font("Arial", 14, FontStyle.Bold)
                headerFont = New Font("Arial", 10, FontStyle.Bold)
                dataFont = New Font("Arial", 8, FontStyle.Regular)
                footerFont = New Font("Arial", 7, FontStyle.Italic)
            End Try
        End Sub

        ''' <summary>
        ''' 🔧 სვეტების მონიშვნის დიალოგის ჩვენება
        ''' </summary>
        Public Function ShowColumnSelectionDialog() As DialogResult
            Try
                Debug.WriteLine("AdvancedDataGridViewPrintService: სვეტების მონიშვნის დიალოგი")

                Using columnDialog As New ColumnSelectionForm(dataGridView)
                    Dim result As DialogResult = columnDialog.ShowDialog()

                    If result = DialogResult.OK Then
                        selectedColumns = columnDialog.GetSelectedColumns()
                        Debug.WriteLine($"AdvancedDataGridViewPrintService: მონიშნულია {selectedColumns.Count} სვეტი")
                    End If

                    Return result
                End Using

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: ShowColumnSelectionDialog შეცდომა: {ex.Message}")
                Return DialogResult.Cancel
            End Try
        End Function

        ''' <summary>
        ''' 🔧 პრინტერის მოწყობის დიალოგის ჩვენება
        ''' </summary>
        Public Function ShowPrinterSetupDialog() As DialogResult
            Try
                Debug.WriteLine("AdvancedDataGridViewPrintService: პრინტერის მოწყობის დიალოგი")

                ' PageSetupDialog - ზომები, ორიენტაცია, მარჯინები
                Dim pageResult As DialogResult = pageSetupDialog.ShowDialog()

                If pageResult = DialogResult.OK Then
                    ' PrintDialog - პრინტერის არჩევა
                    Dim printResult As DialogResult = printDialog.ShowDialog()
                    Return printResult
                End If

                Return pageResult

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: ShowPrinterSetupDialog შეცდომა: {ex.Message}")
                Return DialogResult.Cancel
            End Try
        End Function

        ''' <summary>
        ''' 🔧 სრული ბეჭდვის პროცესი - სვეტების მონიშვნა + პრინტერის არჩევა + ბეჭდვა
        ''' </summary>
        Public Sub ShowFullPrintDialog()
            Try
                Debug.WriteLine("AdvancedDataGridViewPrintService: სრული ბეჭდვის პროცესი")

                If dataGridView Is Nothing OrElse dataGridView.Rows.Count = 0 Then
                    MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 1. სვეტების მონიშვნა
                If ShowColumnSelectionDialog() <> DialogResult.OK Then
                    Debug.WriteLine("AdvancedDataGridViewPrintService: სვეტების მონიშვნა გაუქმებულია")
                    Return
                End If

                If selectedColumns.Count = 0 Then
                    MessageBox.Show("გთხოვთ მონიშნოთ მინიმუმ ერთი სვეტი ბეჭდვისთვის", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 2. პრინტერის მოწყობა
                If ShowPrinterSetupDialog() <> DialogResult.OK Then
                    Debug.WriteLine("AdvancedDataGridViewPrintService: პრინტერის მოწყობა გაუქმებულია")
                    Return
                End If

                ' 3. ბეჭდვის პრევიუ ან პირდაპირ ბეჭდვა
                Dim previewResult As DialogResult = MessageBox.Show(
                    "გსურთ ჯერ ნახოთ ბეჭდვის პრევიუ?",
                    "ბეჭდვის პრევიუ",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                ResetPrintParameters()

                If previewResult = DialogResult.Yes Then
                    printPreviewDialog.ShowDialog()
                Else
                    printDocument.Print()
                End If

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: ShowFullPrintDialog შეცდომა: {ex.Message}")
                MessageBox.Show($"ბეჭდვის შეცდომა: {ex.Message}", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' ბეჭდვის პარამეტრების რესეტი
        ''' </summary>
        Private Sub ResetPrintParameters()
            currentRowIndex = 0
            currentPageNumber = 1
        End Sub

        ''' <summary>
        ''' ბეჭდვის დაწყების ივენთი
        ''' </summary>
        Private Sub OnBeginPrint(sender As Object, e As PrintEventArgs)
            Try
                Debug.WriteLine("AdvancedDataGridViewPrintService: ბეჭდვა იწყება")
                ResetPrintParameters()
            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: OnBeginPrint შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' გვერდის ბეჭდვის ივენთი - ლანდშაფტ ორიენტაციით
        ''' </summary>
        Private Sub OnPrintPage(sender As Object, e As PrintPageEventArgs)
            Try
                Debug.WriteLine($"AdvancedDataGridViewPrintService: გვერდის ბეჭდვა - გვერდი {currentPageNumber} (ლანდშაფტი)")

                ' ლანდშაფტ გვერდის პარამეტრები (სიგრძე უფრო ფართოა)
                Dim graphics As Graphics = e.Graphics
                Dim pageWidth As Integer = e.PageBounds.Width - (marginsSize * 2)
                Dim pageHeight As Integer = e.PageBounds.Height - (marginsSize * 2)
                Dim yPosition As Single = marginsSize

                ' სათაურის ბეჭდვა
                yPosition += PrintTitle(graphics, pageWidth, yPosition)

                ' თარიღისა და გვერდის ნომრის ბეჭდვა
                yPosition += PrintDateAndPageInfo(graphics, pageWidth, yPosition)

                ' მონიშნული სვეტების სათაურების ბეჭდვა
                Dim columnWidths() As Integer = CalculateSelectedColumnWidths(pageWidth)
                yPosition += PrintSelectedHeaders(graphics, columnWidths, yPosition)

                ' მონაცემების ბეჭდვა
                Dim remainingHeight As Integer = CInt(pageHeight - (yPosition - marginsSize))
                Dim rowsPrinted As Integer = PrintSelectedData(graphics, columnWidths, yPosition, remainingHeight)

                ' ქვედა ინფორმაციის ბეჭდვა
                PrintFooter(graphics, pageWidth, pageHeight + marginsSize - 25)

                ' შემდეგი გვერდის საჭიროების შემოწმება
                e.HasMorePages = (currentRowIndex < dataGridView.Rows.Count)

                If e.HasMorePages Then
                    currentPageNumber += 1
                End If

                Debug.WriteLine($"AdvancedDataGridViewPrintService: გვერდი {currentPageNumber} დაბეჭდილია (ლანდშაფტი)")

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: OnPrintPage შეცდომა: {ex.Message}")
                e.HasMorePages = False
            End Try
        End Sub

        ''' <summary>
        ''' სათაურის ბეჭდვა
        ''' </summary>
        Private Function PrintTitle(graphics As Graphics, pageWidth As Integer, yPos As Single) As Single
            Try
                Dim title As String = "განრიგის მონაცემები"
                Dim titleSize As SizeF = graphics.MeasureString(title, titleFont)
                Dim xPos As Single = (pageWidth - titleSize.Width) / 2 + marginsSize

                graphics.DrawString(title, titleFont, titleBrush, xPos, yPos)
                Return titleHeight

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: PrintTitle შეცდომა: {ex.Message}")
                Return titleHeight
            End Try
        End Function

        ''' <summary>
        ''' თარიღისა და გვერდის ინფორმაციის ბეჭდვა
        ''' </summary>
        Private Function PrintDateAndPageInfo(graphics As Graphics, pageWidth As Integer, yPos As Single) As Single
            Try
                ' მარცხენა მხარეს - თარიღი
                Dim dateText As String = $"თარიღი: {DateTime.Now:dd.MM.yyyy HH:mm}"
                graphics.DrawString(dateText, dataFont, dataBrush, marginsSize, yPos)

                ' მარჯვენა მხარეს - გვერდის ნომერი
                Dim pageText As String = $"გვერდი: {currentPageNumber}"
                Dim pageSize As SizeF = graphics.MeasureString(pageText, dataFont)
                graphics.DrawString(pageText, dataFont, dataBrush,
                                  marginsSize + pageWidth - pageSize.Width, yPos)

                ' შუაში - მონიშნული სვეტების რაოდენობა
                Dim columnsText As String = $"სვეტები: {selectedColumns.Count}"
                Dim columnsSize As SizeF = graphics.MeasureString(columnsText, dataFont)
                graphics.DrawString(columnsText, dataFont, dataBrush,
                                  marginsSize + (pageWidth - columnsSize.Width) / 2, yPos)

                Return 25

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: PrintDateAndPageInfo შეცდომა: {ex.Message}")
                Return 25
            End Try
        End Function

        ''' <summary>
        ''' მონიშნული სვეტების სიგანეების გამოთვლა
        ''' </summary>
        Private Function CalculateSelectedColumnWidths(pageWidth As Integer) As Integer()
            Try
                If selectedColumns.Count = 0 Then
                    Return {pageWidth}
                End If

                ' მონიშნული სვეტების სიგანეების გამოთვლა პროპორციულად
                Dim totalOriginalWidth As Integer = selectedColumns.Sum(Function(c) c.Width)
                Dim columnWidths(selectedColumns.Count - 1) As Integer

                For i As Integer = 0 To selectedColumns.Count - 1
                    If totalOriginalWidth > 0 Then
                        columnWidths(i) = CInt((selectedColumns(i).Width / totalOriginalWidth) * pageWidth)
                    Else
                        columnWidths(i) = pageWidth \ selectedColumns.Count
                    End If

                    ' მინიმალური სიგანე
                    If columnWidths(i) < 50 Then columnWidths(i) = 50
                Next

                Return columnWidths

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: CalculateSelectedColumnWidths შეცდომა: {ex.Message}")
                Return {pageWidth}
            End Try
        End Function

        ''' <summary>
        ''' მონიშნული სვეტების სათაურების ბეჭდვა
        ''' </summary>
        Private Function PrintSelectedHeaders(graphics As Graphics, columnWidths() As Integer, yPos As Single) As Single
            Try
                Dim xPos As Single = marginsSize
                Dim headerRect As New Rectangle

                ' მონიშნული სვეტების სათაურები
                For i As Integer = 0 To Math.Min(selectedColumns.Count - 1, columnWidths.Length - 1)
                    headerRect = New Rectangle(CInt(xPos), CInt(yPos), columnWidths(i), headerHeight)

                    ' ფონის ღია ფერი
                    graphics.FillRectangle(New SolidBrush(Color.LightGray), headerRect)
                    graphics.DrawRectangle(Pens.Black, headerRect)

                    ' სათაურის ტექსტი
                    Dim headerText As String = selectedColumns(i).HeaderText
                    Dim textRect As New RectangleF(xPos + 2, yPos + 5, columnWidths(i) - 4, headerHeight - 10)

                    graphics.DrawString(headerText, headerFont, headerBrush, textRect,
                                      New StringFormat() With {.Alignment = StringAlignment.Near,
                                                              .LineAlignment = StringAlignment.Center})

                    xPos += columnWidths(i)
                Next

                Return headerHeight

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: PrintSelectedHeaders შეცდომა: {ex.Message}")
                Return headerHeight
            End Try
        End Function

        ''' <summary>
        ''' მონიშნული სვეტების მონაცემების ბეჭდვა
        ''' </summary>
        Private Function PrintSelectedData(graphics As Graphics, columnWidths() As Integer, yPos As Single, availableHeight As Integer) As Integer
            Try
                Dim rowsPrinted As Integer = 0
                Dim maxRowsPerPage As Integer = (availableHeight \ rowHeight) - 1

                ' მწკრივების ბეჭდვა
                While currentRowIndex < dataGridView.Rows.Count AndAlso rowsPrinted < maxRowsPerPage
                    Dim row As DataGridViewRow = dataGridView.Rows(currentRowIndex)
                    Dim xPos As Single = marginsSize

                    ' მწკრივის ფონის ფერი (ალტერნაცია)
                    Dim rowRect As New Rectangle(marginsSize, CInt(yPos), columnWidths.Sum(), rowHeight)
                    If currentRowIndex Mod 2 = 1 Then
                        graphics.FillRectangle(New SolidBrush(Color.FromArgb(245, 245, 250)), rowRect)
                    End If

                    ' მონიშნული სვეტების მონაცემები
                    For i As Integer = 0 To Math.Min(selectedColumns.Count - 1, columnWidths.Length - 1)
                        Dim cellRect As New Rectangle(CInt(xPos), CInt(yPos), columnWidths(i), rowHeight)
                        graphics.DrawRectangle(Pens.LightGray, cellRect)

                        ' უჯრის მნიშვნელობა
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(selectedColumns(i).Name).Value IsNot Nothing Then
                                cellValue = row.Cells(selectedColumns(i).Name).Value.ToString()
                            End If
                        Catch
                            cellValue = ""
                        End Try

                        ' ტექსტის ბეჭდვა
                        If Not String.IsNullOrEmpty(cellValue) Then
                            Dim textRect As New RectangleF(xPos + 2, yPos + 2, columnWidths(i) - 4, rowHeight - 4)
                            graphics.DrawString(cellValue, dataFont, dataBrush, textRect,
                                              New StringFormat() With {.Alignment = StringAlignment.Near,
                                                                      .LineAlignment = StringAlignment.Center,
                                                                      .Trimming = StringTrimming.EllipsisCharacter})
                        End If

                        xPos += columnWidths(i)
                    Next

                    yPos += rowHeight
                    currentRowIndex += 1
                    rowsPrinted += 1
                End While

                Return rowsPrinted

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: PrintSelectedData შეცდომა: {ex.Message}")
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' ქვედა ინფორმაციის ბეჭდვა
        ''' </summary>
        Private Sub PrintFooter(graphics As Graphics, pageWidth As Integer, yPos As Single)
            Try
                Dim footerText As String = $"სულ მწკრივები: {dataGridView.Rows.Count} | მონიშნული სვეტები: {selectedColumns.Count} | დაბეჭდილია: {DateTime.Now:dd.MM.yyyy HH:mm}"
                Dim footerSize As SizeF = graphics.MeasureString(footerText, footerFont)
                Dim xPos As Single = (pageWidth - footerSize.Width) / 2 + marginsSize

                graphics.DrawString(footerText, footerFont, dataBrush, xPos, yPos)

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: PrintFooter შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' რესურსების განთავისუფლება - IDisposable იმპლემენტაცია
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Debug.WriteLine("AdvancedDataGridViewPrintService: რესურსების განთავისუფლება")

                ' ფონტების განთავისუფლება
                titleFont?.Dispose()
                headerFont?.Dispose()
                dataFont?.Dispose()
                footerFont?.Dispose()

                ' ფერების განთავისუფლება
                titleBrush?.Dispose()
                headerBrush?.Dispose()
                dataBrush?.Dispose()

                ' ბეჭდვის კომპონენტების განთავისუფლება
                printDocument?.Dispose()
                printDialog?.Dispose()
                printPreviewDialog?.Dispose()
                pageSetupDialog?.Dispose()

                ' სვეტების სიის გასუფთავება
                selectedColumns?.Clear()

            Catch ex As Exception
                Debug.WriteLine($"AdvancedDataGridViewPrintService: Dispose შეცდომა: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace