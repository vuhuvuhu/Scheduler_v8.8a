' ===========================================
' 📄 Services/ScheduleFinancialAnalysisService.vb
' -------------------------------------------
' განრიგის ფინანსური ანალიზის სერვისი
' პასუხისმგებელია GBSumFin გრუპბოქსის მართვაზე
' მხოლოდ ადმინისა და მენეჯერისთვის (როლი 1 ან 2)
' ===========================================
Imports System.Drawing
Imports System.Windows.Forms
Imports Scheduler_v8_8a.Services
Imports System.Globalization

Namespace Scheduler_v8_8a.Services

    ''' <summary>
    ''' განრიგის ფინანსური ანალიზის სერვისი
    ''' მართავს GBSumFin პანელს - CheckBox-ების მიხედვით ფინანსური სტატისტიკის ჩვენება
    ''' მხოლოდ ადმინისა და მენეჯერისთვის ხილულია
    ''' </summary>
    Public Class ScheduleFinancialAnalysisService

        ' UI ელემენტები
        Private ReadOnly gbSumFin As GroupBox
        Private ReadOnly financialCheckBoxes As Dictionary(Of String, CheckBox)
        Private ReadOnly financialLabels As Dictionary(Of String, Label)

        ' ფინანსური მონაცემები ქეშირებისთვის
        Private cachedFinancialData As FinancialData = Nothing
        Private lastDataHash As String = ""

        ''' <summary>
        ''' ფინანსური მონაცემების მოდელი
        ''' </summary>
        Public Class FinancialData
            ''' <summary>შესრულებული სეანსების ფინანსები</summary>
            Public Property CompletedPrivate As Decimal = 0
            Public Property CompletedOther As Decimal = 0
            Public Property CompletedTotal As Decimal = 0

            ''' <summary>აღდგენილი სეანსების ფინანსები</summary>
            Public Property RestoredPrivate As Decimal = 0
            Public Property RestoredOther As Decimal = 0
            Public Property RestoredTotal As Decimal = 0

            ''' <summary>გაცდენა არასაპატიო ფინანსები</summary>
            Public Property MissedUnexcusedPrivate As Decimal = 0
            Public Property MissedUnexcusedOther As Decimal = 0
            Public Property MissedUnexcusedTotal As Decimal = 0

            ''' <summary>გაცდენა საპატიო ფინანსები</summary>
            Public Property MissedExcusedPrivate As Decimal = 0
            Public Property MissedExcusedOther As Decimal = 0
            Public Property MissedExcusedTotal As Decimal = 0

            ''' <summary>პროგრამით გატარება ფინანსები</summary>
            Public Property AutoProcessedPrivate As Decimal = 0
            Public Property AutoProcessedOther As Decimal = 0
            Public Property AutoProcessedTotal As Decimal = 0

            ''' <summary>გაუქმებული სეანსების ფინანსები</summary>
            Public Property CancelledPrivate As Decimal = 0
            Public Property CancelledOther As Decimal = 0
            Public Property CancelledTotal As Decimal = 0

            ''' <summary>დაგეგმილი მომავალი ფინანსები</summary>
            Public Property PlannedFuturePrivate As Decimal = 0
            Public Property PlannedFutureOther As Decimal = 0
            Public Property PlannedFutureTotal As Decimal = 0

            ''' <summary>დაგეგმილი ვადაგასული ფინანსები</summary>
            Public Property PlannedOverduePrivate As Decimal = 0
            Public Property PlannedOverdueOther As Decimal = 0
            Public Property PlannedOverdueTotal As Decimal = 0

            ''' <summary>
            ''' თანხის ფორმატირება ₾ სიმბოლოთ
            ''' </summary>
            Public Shared Function FormatAmount(amount As Decimal) As String
                Return $"{amount:F2}₾"
            End Function
        End Class

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        ''' <param name="groupBox">GBSumFin გრუპბოქსი</param>
        Public Sub New(groupBox As GroupBox)
            gbSumFin = groupBox
            financialCheckBoxes = New Dictionary(Of String, CheckBox)()
            financialLabels = New Dictionary(Of String, Label)()

            ' UI ელემენტების ინიციალიზაცია
            InitializeFinancialControls()
        End Sub

        ''' <summary>
        ''' ფინანსური კონტროლების ძებნა და ინიციალიზაცია
        ''' </summary>
        Private Sub InitializeFinancialControls()
            Try
                Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური კონტროლების ძებნა დაიწყო")

                If gbSumFin Is Nothing Then
                    Debug.WriteLine("ScheduleFinancialAnalysisService: GBSumFin გრუპბოქსი არ არის მოწოდებული")
                    Return
                End If

                ' CheckBox-ების სახელები
                Dim checkBoxNames As String() = {
                    "cbs",   ' შესრულებული
                    "cba",   ' აღდგენა
                    "cbga",  ' გაცდენა არასაპატიო
                    "cbgs",  ' გაცდენა საპატიო
                    "cbpg",  ' პროგრამით გატარება
                    "cbg",   ' გაუქმებული
                    "cbd",   ' დაგეგმილი მომავალში
                    "cbdv"   ' დაგეგმილი ვადაგასული
                }

                ' ლეიბლების სახელები
                Dim labelNames As String() = {
                    "lsk", "lsa", "lssum",
                    "lak", "laa", "lasum",
                    "lgak", "lgaa", "lgasum",
                    "lgsk", "lgsa", "lgssum",
                    "lpgk", "lpga", "lpgsum",
                    "lgk", "lga", "lgsum",
                    "ldk", "lda", "ldsum",
                    "ldvk", "ldva", "ldvsum",
                    "lsumk", "lsuma", "lsumsum"
                }

                ' CheckBox-ების ძებნა
                For Each cbName In checkBoxNames
                    Dim foundCheckBox = FindCheckBoxRecursive(gbSumFin, cbName)
                    If foundCheckBox IsNot Nothing Then
                        financialCheckBoxes(cbName) = foundCheckBox
                        Debug.WriteLine($"ScheduleFinancialAnalysisService: ნაპოვნია CheckBox '{cbName}'")

                        ' ივენთის მიბმა
                        AddHandler foundCheckBox.CheckedChanged, AddressOf OnCheckBoxChanged
                    Else
                        Debug.WriteLine($"ScheduleFinancialAnalysisService: ვერ მოიძებნა CheckBox '{cbName}'")
                    End If
                Next

                ' ლეიბლების ძებნა
                For Each labelName In labelNames
                    Dim foundLabel = FindLabelRecursive(gbSumFin, labelName)
                    If foundLabel IsNot Nothing Then
                        financialLabels(labelName) = foundLabel
                        Debug.WriteLine($"ScheduleFinancialAnalysisService: ნაპოვნია Label '{labelName}'")
                    Else
                        Debug.WriteLine($"ScheduleFinancialAnalysisService: ვერ მოიძებნა Label '{labelName}'")
                    End If
                Next

                Debug.WriteLine($"ScheduleFinancialAnalysisService: ნაპოვნია {financialCheckBoxes.Count} CheckBox და {financialLabels.Count} Label")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: InitializeFinancialControls შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' CheckBox-ის რეკურსიული ძებნა
        ''' </summary>
        Private Function FindCheckBoxRecursive(parent As Control, name As String) As CheckBox
            Try
                ' ჯერ პირდაპირ შვილებში ვეძებთ
                For Each ctrl As Control In parent.Controls
                    If TypeOf ctrl Is CheckBox AndAlso
                       String.Equals(ctrl.Name, name, StringComparison.OrdinalIgnoreCase) Then
                        Return DirectCast(ctrl, CheckBox)
                    End If
                Next

                ' მერე რეკურსიულად შვილების შვილებში
                For Each ctrl As Control In parent.Controls
                    Dim found = FindCheckBoxRecursive(ctrl, name)
                    If found IsNot Nothing Then
                        Return found
                    End If
                Next

                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' ლეიბლის რეკურსიული ძებნა
        ''' </summary>
        Private Function FindLabelRecursive(parent As Control, name As String) As Label
            Try
                ' ჯერ პირდაპირ შვილებში ვეძებთ
                For Each ctrl As Control In parent.Controls
                    If TypeOf ctrl Is Label AndAlso
                       String.Equals(ctrl.Name, name, StringComparison.OrdinalIgnoreCase) Then
                        Return DirectCast(ctrl, Label)
                    End If
                Next

                ' მერე რეკურსიულად შვილების შვილებში
                For Each ctrl As Control In parent.Controls
                    Dim found = FindLabelRecursive(ctrl, name)
                    If found IsNot Nothing Then
                        Return found
                    End If
                Next

                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' მომხმარებლის როლის მიხედვით ხილულობის კონტროლი
        ''' </summary>
        ''' <param name="userRole">მომხმარებლის როლი (1=ადმინი, 2=მენეჯერი, 6=მომხმარებელი)</param>
        Public Sub SetVisibilityByUserRole(userRole As Integer)
            Try
                Debug.WriteLine($"ScheduleFinancialAnalysisService: ხილულობის დაყენება როლისთვის: {userRole}")

                If gbSumFin Is Nothing Then Return

                ' ფინანსური პანელი ხილულია მხოლოდ ადმინისა და მენეჯერისთვის
                Dim isVisible As Boolean = (userRole = 1 OrElse userRole = 2)

                gbSumFin.Visible = isVisible

                If isVisible Then
                    Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური პანელი ხილულია")
                Else
                    Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური პანელი დამალულია")
                End If

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: SetVisibilityByUserRole შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ფინანსური მონაცემების გამოთვლა გაფილტრული მონაცემებისთვის
        ''' </summary>
        ''' <param name="filteredData">გაფილტრული მონაცემები</param>
        ''' <returns>ფინანსური მონაცემები</returns>
        Public Function CalculateFinancialData(filteredData As List(Of IList(Of Object))) As FinancialData
            Try
                Debug.WriteLine($"ScheduleFinancialAnalysisService: ფინანსური მონაცემების გამოთვლა {filteredData.Count} ჩანაწერისთვის")

                Dim financial As New FinancialData()
                Dim currentTime As DateTime = DateTime.Now

                ' ყოველი ჩანაწერის დამუშავება
                For Each row In filteredData
                    Try
                        ' ძირითადი ვალიდაცია
                        If row.Count < 14 Then Continue For

                        ' სტატუსის მიღება
                        Dim status As String = If(row.Count > 12, row(12).ToString().Trim().ToLower(), "")

                        ' ფასის მიღება
                        Dim price As Decimal = 0
                        If row.Count > 11 AndAlso IsNumeric(row(11)) Then
                            Decimal.TryParse(row(11).ToString().Replace(",", "."),
                                           NumberStyles.Any, CultureInfo.InvariantCulture, price)
                        End If

                        ' დაფინანსების ტიპის მიღება
                        Dim funding As String = If(row.Count > 13, row(13).ToString().Trim().ToLower(), "")
                        Dim isPrivate As Boolean = String.Equals(funding, "კერძო", StringComparison.OrdinalIgnoreCase)

                        ' სესიის თარიღის მიღება
                        Dim sessionDate As DateTime = DateTime.Now
                        If row.Count > 5 Then
                            DateTime.TryParse(row(5).ToString(), sessionDate)
                        End If

                        ' სტატუსის მიხედვით კატეგორიზაცია
                        Select Case status
                            Case "შესრულებული"
                                If isPrivate Then
                                    financial.CompletedPrivate += price
                                Else
                                    financial.CompletedOther += price
                                End If
                                financial.CompletedTotal += price

                            Case "აღდგენა"
                                If isPrivate Then
                                    financial.RestoredPrivate += price
                                Else
                                    financial.RestoredOther += price
                                End If
                                financial.RestoredTotal += price

                            Case "გაცდენა არასაპატიო"
                                If isPrivate Then
                                    financial.MissedUnexcusedPrivate += price
                                Else
                                    financial.MissedUnexcusedOther += price
                                End If
                                financial.MissedUnexcusedTotal += price

                            Case "გაცდენა საპატიო"
                                If isPrivate Then
                                    financial.MissedExcusedPrivate += price
                                Else
                                    financial.MissedExcusedOther += price
                                End If
                                financial.MissedExcusedTotal += price

                            Case "პროგრამით გატარება"
                                If isPrivate Then
                                    financial.AutoProcessedPrivate += price
                                Else
                                    financial.AutoProcessedOther += price
                                End If
                                financial.AutoProcessedTotal += price

                            Case "გაუქმება", "გაუქმებული"
                                If isPrivate Then
                                    financial.CancelledPrivate += price
                                Else
                                    financial.CancelledOther += price
                                End If
                                financial.CancelledTotal += price

                            Case "დაგეგმილი"
                                ' დაგეგმილი სეანსებისთვის ვამოწმებთ ვადაგადაცილებას
                                If sessionDate <= currentTime Then
                                    ' ვადაგასული
                                    If isPrivate Then
                                        financial.PlannedOverduePrivate += price
                                    Else
                                        financial.PlannedOverdueOther += price
                                    End If
                                    financial.PlannedOverdueTotal += price
                                Else
                                    ' მომავალი
                                    If isPrivate Then
                                        financial.PlannedFuturePrivate += price
                                    Else
                                        financial.PlannedFutureOther += price
                                    End If
                                    financial.PlannedFutureTotal += price
                                End If
                        End Select

                    Catch ex As Exception
                        Debug.WriteLine($"ScheduleFinancialAnalysisService: მწკრივის დამუშავების შეცდომა: {ex.Message}")
                        Continue For
                    End Try
                Next

                Debug.WriteLine($"ScheduleFinancialAnalysisService: ფინანსური მონაცემები გამოთვლილია - შესრულებული სულ: {financial.CompletedTotal:F2}")
                Return financial

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: CalculateFinancialData შეცდომა: {ex.Message}")
                Return New FinancialData()
            End Try
        End Function

        ''' <summary>
        ''' ფინანსური მონაცემების განახლება
        ''' </summary>
        ''' <param name="filteredData">გაფილტრული მონაცემები</param>
        Public Sub UpdateFinancialData(filteredData As List(Of IList(Of Object)))
            Try
                Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური მონაცემების განახლება")

                ' ფინანსური მონაცემების გამოთვლა
                cachedFinancialData = CalculateFinancialData(filteredData)

                ' UI-ის განახლება
                UpdateFinancialDisplay()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: UpdateFinancialData შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ფინანსური ინფორმაციის ჩვენება CheckBox-ების მდგომარეობის მიხედვით
        ''' </summary>
        Private Sub UpdateFinancialDisplay()
            Try
                If cachedFinancialData Is Nothing Then
                    Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური მონაცემები ცარიელია")
                    Return
                End If

                Debug.WriteLine("ScheduleFinancialAnalysisService: ფინანსური ჩვენების განახლება")

                ' ყოველი კატეგორიის განახლება
                UpdateCategoryDisplay("cbs", "lsk", "lsa", "lssum",
                                    cachedFinancialData.CompletedPrivate,
                                    cachedFinancialData.CompletedOther,
                                    cachedFinancialData.CompletedTotal)

                UpdateCategoryDisplay("cba", "lak", "laa", "lasum",
                                    cachedFinancialData.RestoredPrivate,
                                    cachedFinancialData.RestoredOther,
                                    cachedFinancialData.RestoredTotal)

                UpdateCategoryDisplay("cbga", "lgak", "lgaa", "lgasum",
                                    cachedFinancialData.MissedUnexcusedPrivate,
                                    cachedFinancialData.MissedUnexcusedOther,
                                    cachedFinancialData.MissedUnexcusedTotal)

                UpdateCategoryDisplay("cbgs", "lgsk", "lgsa", "lgssum",
                                    cachedFinancialData.MissedExcusedPrivate,
                                    cachedFinancialData.MissedExcusedOther,
                                    cachedFinancialData.MissedExcusedTotal)

                UpdateCategoryDisplay("cbpg", "lpgk", "lpga", "lpgsum",
                                    cachedFinancialData.AutoProcessedPrivate,
                                    cachedFinancialData.AutoProcessedOther,
                                    cachedFinancialData.AutoProcessedTotal)

                UpdateCategoryDisplay("cbg", "lgk", "lga", "lgsum",
                                    cachedFinancialData.CancelledPrivate,
                                    cachedFinancialData.CancelledOther,
                                    cachedFinancialData.CancelledTotal)

                UpdateCategoryDisplay("cbd", "ldk", "lda", "ldsum",
                                    cachedFinancialData.PlannedFuturePrivate,
                                    cachedFinancialData.PlannedFutureOther,
                                    cachedFinancialData.PlannedFutureTotal)

                UpdateCategoryDisplay("cbdv", "ldvk", "ldva", "ldvsum",
                                    cachedFinancialData.PlannedOverduePrivate,
                                    cachedFinancialData.PlannedOverdueOther,
                                    cachedFinancialData.PlannedOverdueTotal)

                ' საერთო ჯამების განახლება
                UpdateTotalSums()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: UpdateFinancialDisplay შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' კონკრეტული კატეგორიის ჩვენების განახლება
        ''' </summary>
        ''' <param name="checkBoxName">CheckBox-ის სახელი</param>
        ''' <param name="privateLabel">კერძო ლეიბლის სახელი</param>
        ''' <param name="otherLabel">სხვა ლეიბლის სახელი</param>
        ''' <param name="totalLabel">სულ ლეიბლის სახელი</param>
        ''' <param name="privateAmount">კერძო თანხა</param>
        ''' <param name="otherAmount">სხვა თანხა</param>
        ''' <param name="totalAmount">სულ თანხა</param>
        Private Sub UpdateCategoryDisplay(checkBoxName As String,
                                        privateLabel As String,
                                        otherLabel As String,
                                        totalLabel As String,
                                        privateAmount As Decimal,
                                        otherAmount As Decimal,
                                        totalAmount As Decimal)
            Try
                ' CheckBox-ის მდგომარეობის შემოწმება
                Dim isChecked As Boolean = False
                If financialCheckBoxes.ContainsKey(checkBoxName) AndAlso financialCheckBoxes(checkBoxName) IsNot Nothing Then
                    isChecked = financialCheckBoxes(checkBoxName).Checked
                End If

                ' ფერის განსაზღვრა
                Dim labelColor As Color = If(isChecked, Color.Black, Color.Gray)

                ' ლეიბლების განახლება
                UpdateFinancialLabel(privateLabel, FinancialData.FormatAmount(privateAmount), labelColor)
                UpdateFinancialLabel(otherLabel, FinancialData.FormatAmount(otherAmount), labelColor)
                UpdateFinancialLabel(totalLabel, FinancialData.FormatAmount(totalAmount), labelColor)

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: UpdateCategoryDisplay შეცდომა '{checkBoxName}': {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' საერთო ჯამების განახლება მონიშნული კატეგორიების მიხედვით
        ''' </summary>
        Private Sub UpdateTotalSums()
            Try
                Dim totalPrivate As Decimal = 0
                Dim totalOther As Decimal = 0

                ' ყოველი მონიშნული კატეგორიის ჯამი
                If IsCheckBoxChecked("cbs") Then
                    totalPrivate += cachedFinancialData.CompletedPrivate
                    totalOther += cachedFinancialData.CompletedOther
                End If

                If IsCheckBoxChecked("cba") Then
                    totalPrivate += cachedFinancialData.RestoredPrivate
                    totalOther += cachedFinancialData.RestoredOther
                End If

                If IsCheckBoxChecked("cbga") Then
                    totalPrivate += cachedFinancialData.MissedUnexcusedPrivate
                    totalOther += cachedFinancialData.MissedUnexcusedOther
                End If

                If IsCheckBoxChecked("cbgs") Then
                    totalPrivate += cachedFinancialData.MissedExcusedPrivate
                    totalOther += cachedFinancialData.MissedExcusedOther
                End If

                If IsCheckBoxChecked("cbpg") Then
                    totalPrivate += cachedFinancialData.AutoProcessedPrivate
                    totalOther += cachedFinancialData.AutoProcessedOther
                End If

                If IsCheckBoxChecked("cbg") Then
                    totalPrivate += cachedFinancialData.CancelledPrivate
                    totalOther += cachedFinancialData.CancelledOther
                End If

                If IsCheckBoxChecked("cbd") Then
                    totalPrivate += cachedFinancialData.PlannedFuturePrivate
                    totalOther += cachedFinancialData.PlannedFutureOther
                End If

                If IsCheckBoxChecked("cbdv") Then
                    totalPrivate += cachedFinancialData.PlannedOverduePrivate
                    totalOther += cachedFinancialData.PlannedOverdueOther
                End If

                Dim grandTotal As Decimal = totalPrivate + totalOther

                ' საერთო ჯამების ლეიბლების განახლება
                UpdateFinancialLabel("lsumk", FinancialData.FormatAmount(totalPrivate), Color.Black)
                UpdateFinancialLabel("lsuma", FinancialData.FormatAmount(totalOther), Color.Black)
                UpdateFinancialLabel("lsumsum", FinancialData.FormatAmount(grandTotal), Color.Black)

                Debug.WriteLine($"ScheduleFinancialAnalysisService: საერთო ჯამები - კერძო: {totalPrivate:F2}, სხვა: {totalOther:F2}, სულ: {grandTotal:F2}")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: UpdateTotalSums შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' CheckBox-ის მონიშნულობის შემოწმება
        ''' </summary>
        ''' <param name="checkBoxName">CheckBox-ის სახელი</param>
        ''' <returns>True თუ მონიშნულია</returns>
        Private Function IsCheckBoxChecked(checkBoxName As String) As Boolean
            Try
                If financialCheckBoxes.ContainsKey(checkBoxName) AndAlso financialCheckBoxes(checkBoxName) IsNot Nothing Then
                    Return financialCheckBoxes(checkBoxName).Checked
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' ფინანსური ლეიბლის განახლება
        ''' </summary>
        ''' <param name="labelName">ლეიბლის სახელი</param>
        ''' <param name="value">ახალი მნიშვნელობა</param>
        ''' <param name="color">ფერი</param>
        Private Sub UpdateFinancialLabel(labelName As String, value As String, color As Color)
            Try
                If financialLabels.ContainsKey(labelName) AndAlso financialLabels(labelName) IsNot Nothing Then
                    financialLabels(labelName).Text = value
                    financialLabels(labelName).ForeColor = color
                    Debug.WriteLine($"ScheduleFinancialAnalysisService: ლეიბლი '{labelName}' განახლდა: '{value}', ფერი: {color.Name}")
                Else
                    Debug.WriteLine($"ScheduleFinancialAnalysisService: ლეიბლი '{labelName}' ვერ მოიძებნა განსახლებისთვის")
                End If
            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: UpdateFinancialLabel შეცდომა '{labelName}': {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' CheckBox-ის შეცვლის ივენთ ჰენდლერი
        ''' </summary>
        Private Sub OnCheckBoxChanged(sender As Object, e As EventArgs)
            Try
                Dim checkBox As CheckBox = CType(sender, CheckBox)
                Debug.WriteLine($"ScheduleFinancialAnalysisService: CheckBox '{checkBox.Name}' შეიცვალა: {checkBox.Checked}")

                ' ფინანსური ჩვენების განახლება
                UpdateFinancialDisplay()

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: OnCheckBoxChanged შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ყველა ლეიბლის გასუფთავება
        ''' </summary>
        Public Sub ClearAllLabels()
            Try
                Debug.WriteLine("ScheduleFinancialAnalysisService: ყველა ლეიბლის გასუფთავება")

                For Each kvp In financialLabels
                    If kvp.Value IsNot Nothing Then
                        kvp.Value.Text = "0.00₾"
                        kvp.Value.ForeColor = Color.Gray
                    End If
                Next

                Debug.WriteLine("ScheduleFinancialAnalysisService: ყველა ლეიბლი გასუფთავებულია")

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: ClearAllLabels შეცდომა: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' ყველა CheckBox-ის მონიშვნა/განიშვნა
        ''' </summary>
        ''' <param name="checkAll">True - ყველას მონიშვნა, False - ყველას განიშვნა</param>
        Public Sub SetAllCheckBoxes(checkAll As Boolean)
            Try
                Debug.WriteLine($"ScheduleFinancialAnalysisService: ყველა CheckBox-ის დაყენება: {checkAll}")

                For Each kvp In financialCheckBoxes
                    If kvp.Value IsNot Nothing Then
                        kvp.Value.Checked = checkAll
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine($"ScheduleFinancialAnalysisService: SetAllCheckBoxes შეცდომა: {ex.Message}")
            End Try
        End Sub

    End Class
End Namespace