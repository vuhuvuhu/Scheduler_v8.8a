' ===========================================
' 📄 Services/DataGridViewPrintService.vb
' -------------------------------------------
' DataGridView-ის მონაცემების ბეჭდვის სერვისი
' ქართული ფონტებისა და ლეიაუტის მხარდაჭერით
' ===========================================
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' DataGridView-ის ბეჭდვის სერვისი
    ''' უზრუნველყოფს ქართული ტექსტის სწორ ბეჭდვას და ლეიაუტს
    ''' </summary>
    Public Class DataGridViewPrintService
        Implements IDisposable

        ' ბეჭდვის პარამეტრები
        Private ReadOnly dataGridView As DataGridView
        Private printDocument As PrintDocument
        Private printDialog As PrintDialog
        Private printPreviewDialog As PrintPreviewDialog

        ' ბეჭდვის მდგომარეობა
        Private currentRowIndex As Integer = 0
        Private currentPageNumber As Integer = 1

        ' ფონტები და სტილები
        Private titleFont As Font
        Private headerFont As Font
        Private dataFont As Font
        Private footerFont As Font

        ' ფერები
        Private headerBrush As SolidBrush
        Private dataBrush As SolidBrush
        Private titleBrush As SolidBrush

        ' გვერდის პარამეტრები
        Private Const marginsSize As Integer = 50
        Private Const rowHeight As Integer = 25
        Private Const headerHeight As Integer = 35
        Private Const titleHeight As Integer = 50

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="dgv">ბეჭდვისთვის DataGridView</param>
        Public Sub New(dgv As DataGridView)
            dataGridView = dgv
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

                ' PrintDialog-ის შექმნა
                printDialog = New PrintDialog()
                printDialog.Document = printDocument

                ' PrintPreviewDialog-ის შექმნა
                printPreviewDialog = New PrintPreviewDialog()
                printPreviewDialog.Document = printDocument
                printPreviewDialog.WindowState = FormWindowState.Maximized

                Debug.WriteLine("DataGridViewPrintService: ბეჭდვის კომპონენტები ინიციალიზებულია")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: კომპონენტების ინიციალიზაციის შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ფონტებისა და ფერების ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeFontsAndColors()
            Try
                ' ქართული ფონტები
                titleFont = New Font("Sylfaen", 16, FontStyle.Bold)
                headerFont = New Font("Sylfaen", 11, FontStyle.Bold)
                dataFont = New Font("Sylfaen", 9, FontStyle.Regular)
                footerFont = New Font("Sylfaen", 8, FontStyle.Italic)

                ' ფერები
                titleBrush = New SolidBrush(Color.Black)
                headerBrush = New SolidBrush(Color.DarkBlue)
                dataBrush = New SolidBrush(Color.Black)

                Debug.WriteLine("DataGridViewPrintService: ფონტები და ფერები ინიციალიზებულია")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: ფონტების ინიციალიზაციის შეცდომა: {ex.Message}")
                ' ნაგულისხმევი ფონტები
                titleFont = New Font("Arial", 16, FontStyle.Bold)
                headerFont = New Font("Arial", 11, FontStyle.Bold)
                dataFont = New Font("Arial", 9, FontStyle.Regular)
                footerFont = New Font("Arial", 8, FontStyle.Italic)
            End Try
        End Sub

        ''' <summary>
        ''' ბეჭდვის პრევიუს ჩვენება
        ''' </summary>
        Public Sub ShowPrintPreview()
            Try
                Debug.WriteLine("DataGridViewPrintService: ბეჭდვის პრევიუს ჩვენება")

                If dataGridView Is Nothing OrElse dataGridView.Rows.Count = 0 Then
                    MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' ბეჭდვის პარამეტრების რესეტი
                ResetPrintParameters()

                ' პრევიუს ჩვენება
                printPreviewDialog.ShowDialog()

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: ShowPrintPreview შეცდომა: {ex.Message}")
                MessageBox.Show($"ბეჭდვის პრევიუს შეცდომა: {ex.Message}", "შეცდომა",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' პირდაპირ ბეჭდვა
        ''' </summary>
        Public Sub Print()
            Try
                Debug.WriteLine("DataGridViewPrintService: პირდაპირ ბეჭდვა")

                If dataGridView Is Nothing OrElse dataGridView.Rows.Count = 0 Then
                    MessageBox.Show("ბეჭდვისთვის მონაცემები არ არის ხელმისაწვდომი", "ინფორმაცია",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' ბეჭდვის დიალოგის ჩვენება
                If printDialog.ShowDialog() = DialogResult.OK Then
                    ResetPrintParameters()
                    printDocument.Print()
                End If

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: Print შეცდომა: {ex.Message}")
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
                Debug.WriteLine("DataGridViewPrintService: ბეჭდვა იწყება")
                ResetPrintParameters()
            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: OnBeginPrint შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' გვერდის ბეჭდვის ივენთი
        ''' </summary>
        Private Sub OnPrintPage(sender As Object, e As PrintPageEventArgs)
            Try
                Debug.WriteLine($"DataGridViewPrintService: გვერდის ბეჭდვა - გვერდი {currentPageNumber}")

                ' გვერდის პარამეტრები
                Dim graphics As Graphics = e.Graphics
                Dim pageWidth As Integer = e.PageBounds.Width - (marginsSize * 2)
                Dim pageHeight As Integer = e.PageBounds.Height - (marginsSize * 2)
                Dim yPosition As Single = marginsSize

                ' სათაურის ბეჭდვა
                yPosition += PrintTitle(graphics, pageWidth, yPosition)

                ' თარიღისა და გვერდის ნომრის ბეჭდვა
                yPosition += PrintDateAndPageInfo(graphics, pageWidth, yPosition)

                ' სვეტების სათაურების ბეჭდვა
                Dim columnWidths() As Integer = CalculateColumnWidths(pageWidth)
                yPosition += PrintHeaders(graphics, columnWidths, yPosition)

                ' მონაცემების ბეჭდვა
                Dim remainingHeight As Integer = CInt(pageHeight - (yPosition - marginsSize))
                Dim rowsPrinted As Integer = PrintData(graphics, columnWidths, yPosition, remainingHeight)

                ' ქვედა ინფორმაციის ბეჭდვა
                PrintFooter(graphics, pageWidth, pageHeight + marginsSize - 30)

                ' შემდეგი გვერდის საჭიროების შემოწმება
                e.HasMorePages = (currentRowIndex < dataGridView.Rows.Count)

                If e.HasMorePages Then
                    currentPageNumber += 1
                End If

                Debug.WriteLine($"DataGridViewPrintService: გვერდი {currentPageNumber} დაბეჭდილია, დარჩა {dataGridView.Rows.Count - currentRowIndex} მწკრივი")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: OnPrintPage შეცდომა: {ex.Message}")
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
                Debug.WriteLine($"DataGridViewPrintService: PrintTitle შეცდომა: {ex.Message}")
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

                Return 30

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: PrintDateAndPageInfo შეცდომა: {ex.Message}")
                Return 30
            End Try
        End Function

        ''' <summary>
        ''' სვეტების სიგანეების გამოთვლა
        ''' </summary>
        Private Function CalculateColumnWidths(pageWidth As Integer) As Integer()
            Try
                ' მხოლოდ ხილული სვეტების რაოდენობა
                Dim visibleColumns As New List(Of DataGridViewColumn)
                For Each col As DataGridViewColumn In dataGridView.Columns
                    If col.Visible AndAlso col.Name <> "Edit" Then ' რედაქტირების ღილაკს გამოვრიცხავთ
                        visibleColumns.Add(col)
                    End If
                Next

                If visibleColumns.Count = 0 Then
                    Return {pageWidth}
                End If

                ' სვეტების სიგანეების გამოთვლა პროპორციულად
                Dim totalOriginalWidth As Integer = visibleColumns.Sum(Function(c) c.Width)
                Dim columnWidths(visibleColumns.Count - 1) As Integer

                For i As Integer = 0 To visibleColumns.Count - 1
                    If totalOriginalWidth > 0 Then
                        columnWidths(i) = CInt((visibleColumns(i).Width / totalOriginalWidth) * pageWidth)
                    Else
                        columnWidths(i) = pageWidth \ visibleColumns.Count
                    End If
                Next

                Return columnWidths

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: CalculateColumnWidths შეცდომა: {ex.Message}")
                Return {pageWidth}
            End Try
        End Function

        ''' <summary>
        ''' სვეტების სათაურების ბეჭდვა
        ''' </summary>
        Private Function PrintHeaders(graphics As Graphics, columnWidths() As Integer, yPos As Single) As Single
            Try
                Dim xPos As Single = marginsSize
                Dim headerRect As New Rectangle

                ' ხილული სვეტების სათაურები
                Dim visibleColumns As New List(Of DataGridViewColumn)
                For Each col As DataGridViewColumn In dataGridView.Columns
                    If col.Visible AndAlso col.Name <> "Edit" Then
                        visibleColumns.Add(col)
                    End If
                Next

                ' თითოეული სვეტისთვის
                For i As Integer = 0 To Math.Min(visibleColumns.Count - 1, columnWidths.Length - 1)
                    headerRect = New Rectangle(CInt(xPos), CInt(yPos), columnWidths(i), headerHeight)

                    ' ფონის ღია ფერი
                    graphics.FillRectangle(New SolidBrush(Color.LightGray), headerRect)
                    graphics.DrawRectangle(Pens.Black, headerRect)

                    ' სათაურის ტექსტი
                    Dim headerText As String = visibleColumns(i).HeaderText
                    Dim textRect As New RectangleF(xPos + 2, yPos + 5, columnWidths(i) - 4, headerHeight - 10)

                    graphics.DrawString(headerText, headerFont, headerBrush, textRect,
                                      New StringFormat() With {.Alignment = StringAlignment.Near,
                                                              .LineAlignment = StringAlignment.Center})

                    xPos += columnWidths(i)
                Next

                Return headerHeight

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: PrintHeaders შეცდომა: {ex.Message}")
                Return headerHeight
            End Try
        End Function

        ''' <summary>
        ''' მონაცემების ბეჭდვა
        ''' </summary>
        Private Function PrintData(graphics As Graphics, columnWidths() As Integer, yPos As Single, availableHeight As Integer) As Integer
            Try
                Dim rowsPrinted As Integer = 0
                Dim maxRowsPerPage As Integer = (availableHeight \ rowHeight) - 1

                ' ხილული სვეტები
                Dim visibleColumns As New List(Of DataGridViewColumn)
                For Each col As DataGridViewColumn In dataGridView.Columns
                    If col.Visible AndAlso col.Name <> "Edit" Then
                        visibleColumns.Add(col)
                    End If
                Next

                ' მწკრივების ბეჭდვა
                While currentRowIndex < dataGridView.Rows.Count AndAlso rowsPrinted < maxRowsPerPage
                    Dim row As DataGridViewRow = dataGridView.Rows(currentRowIndex)
                    Dim xPos As Single = marginsSize

                    ' მწკრივის ფონის ფერი (ალტერნაცია)
                    Dim rowRect As New Rectangle(marginsSize, CInt(yPos), columnWidths.Sum(), rowHeight)
                    If currentRowIndex Mod 2 = 1 Then
                        graphics.FillRectangle(New SolidBrush(Color.FromArgb(245, 245, 250)), rowRect)
                    End If

                    ' თითოეული უჯრისთვის
                    For i As Integer = 0 To Math.Min(visibleColumns.Count - 1, columnWidths.Length - 1)
                        Dim cellRect As New Rectangle(CInt(xPos), CInt(yPos), columnWidths(i), rowHeight)
                        graphics.DrawRectangle(Pens.LightGray, cellRect)

                        ' უჯრის მნიშვნელობა
                        Dim cellValue As String = ""
                        Try
                            If row.Cells(visibleColumns(i).Name).Value IsNot Nothing Then
                                cellValue = row.Cells(visibleColumns(i).Name).Value.ToString()
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
                Debug.WriteLine($"DataGridViewPrintService: PrintData შეცდომა: {ex.Message}")
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' ქვედა ინფორმაციის ბეჭდვა
        ''' </summary>
        Private Sub PrintFooter(graphics As Graphics, pageWidth As Integer, yPos As Single)
            Try
                Dim footerText As String = $"სულ მწკრივები: {dataGridView.Rows.Count} | დაბეჭდილია: {DateTime.Now:dd.MM.yyyy HH:mm}"
                Dim footerSize As SizeF = graphics.MeasureString(footerText, footerFont)
                Dim xPos As Single = (pageWidth - footerSize.Width) / 2 + marginsSize

                graphics.DrawString(footerText, footerFont, dataBrush, xPos, yPos)

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: PrintFooter შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' რესურსების განთავისუფლება - IDisposable იმპლემენტაცია
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Debug.WriteLine("DataGridViewPrintService: რესურსების განთავისუფლება")

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

                Debug.WriteLine("DataGridViewPrintService: ყველა რესურსი განთავისუფლებულია")

            Catch ex As Exception
                Debug.WriteLine($"DataGridViewPrintService: Dispose შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Finalizer - უსაფრთხოების მიზნით
        ''' </summary>
        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub

    End Class
End Namespace