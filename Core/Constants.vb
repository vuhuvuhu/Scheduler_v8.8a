' ===========================================
' ?? Core/Constants.vb
' -------------------------------------------
' ?????????????? ??????????? (??????, ?????????, ????? ???????????, ??????? ???????)
' ===========================================
Imports System.Drawing

Public Module Roles
    Public Const Admin As String = "1"
    Public Const Manager As String = "2"
    Public Const Staff As String = "3"
    Public Const Basic As String = "6" ' ???????????? / ?????????????? ????????????
End Module

Public Module Statuses
    Public Const Planned As String = "?????????"
    Public Const Completed As String = "???????????"
    Public Const Cancelled As String = "??????????"
    Public Const Cancel As String = "????????"
    Public Const Restored As String = "???????"
    Public Const MissedUnexcused As String = "??????? ??????????"
    Public Const MissedExcused As String = "??????? ???????"
    Public Const AutoProcessed As String = "????????? ????????"
End Module

Public Module SheetRanges
    Public Const Users = "DB-Users!B2:C"
    Public Const Birthdays = "DB-Personal!B2:G"
    Public Const Sessions = "DB-Schedule!A2:O"
    Public Const Tasks = "DB-Tasks!B2:M"
End Module

Public Module ColorPalette
    Private ReadOnly NeutralHeader As Color = Color.FromArgb(60, 60, 60)

    Public Function GetSessionCardColor(status As String, Optional isOverdue As Boolean = False) As Color
        If String.IsNullOrWhiteSpace(status) Then Return Color.FromArgb(245, 245, 245)
        Dim s = status.Trim().ToLowerInvariant()
        If isOverdue AndAlso (s = Statuses.Planned OrElse s = Statuses.MissedUnexcused.ToLower() OrElse s = Statuses.MissedExcused.ToLower()) Then
            Return Color.FromArgb(255, 235, 235) ' ???????????????? - ??? ??????
        End If
        Select Case s
            Case Statuses.Completed.ToLower() : Return Color.FromArgb(225, 245, 225) ' ?????? ???
            Case Statuses.Planned.ToLower() : Return Color.FromArgb(235, 240, 255) ' ??? ??????-???????????
            Case Statuses.Cancelled.ToLower(), Statuses.Cancel.ToLower() : Return Color.FromArgb(240, 240, 240) ' ?????????? ??????????
            Case Statuses.Restored.ToLower() : Return Color.FromArgb(230, 250, 240) ' ??? ?????????
            Case Statuses.MissedUnexcused.ToLower() : Return Color.FromArgb(255, 230, 230)
            Case Statuses.MissedExcused.ToLower() : Return Color.FromArgb(255, 245, 230)
            Case Statuses.AutoProcessed.ToLower() : Return Color.FromArgb(230, 240, 255)
            Case Else : Return Color.FromArgb(245, 245, 245)
        End Select
    End Function

    Public Function GetSessionHeaderColor(status As String, Optional isOverdue As Boolean = False) As Color
        If String.IsNullOrWhiteSpace(status) Then Return NeutralHeader
        Dim s = status.Trim().ToLowerInvariant()
        If isOverdue AndAlso (s = Statuses.Planned OrElse s.StartsWith("???????")) Then
            Return Color.FromArgb(180, 0, 0) ' ???? ??????
        End If
        Select Case s
            Case Statuses.Completed.ToLower() : Return Color.FromArgb(0, 120, 0)
            Case Statuses.Planned.ToLower() : Return Color.FromArgb(0, 70, 140)
            Case Statuses.Cancelled.ToLower(), Statuses.Cancel.ToLower() : Return Color.FromArgb(90, 90, 90)
            Case Statuses.Restored.ToLower() : Return Color.FromArgb(0, 100, 80)
            Case Statuses.MissedUnexcused.ToLower() : Return Color.FromArgb(180, 0, 0)
            Case Statuses.MissedExcused.ToLower() : Return Color.FromArgb(200, 110, 0)
            Case Statuses.AutoProcessed.ToLower() : Return Color.FromArgb(0, 90, 160)
            Case Else : Return NeutralHeader
        End Select
    End Function
End Module
