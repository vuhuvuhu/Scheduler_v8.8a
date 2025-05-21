' ===========================================
' 📄 Controls/RotatedLabel.vb
' -------------------------------------------
' შემობრუნებული ტექსტის ლეიბლი (კალენდრის თარიღის ზოლისთვის)
' ===========================================
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace Scheduler_v8_8a.Controls
    ''' <summary>
    ''' შემობრუნებული ტექსტის ლეიბლი - საშუალებას გვაძლევს შევაბრუნოთ ტექსტი ნებისმიერი კუთხით
    ''' </summary>
    Public Class RotatedLabel
        Inherits Control

        Private _text As String = String.Empty
        Private _rotationAngle As Integer = 0

        ''' <summary>
        ''' შესაბრუნებელი კუთხე გრადუსებში (0-360)
        ''' </summary>
        Public Property RotationAngle As Integer
            Get
                Return _rotationAngle
            End Get
            Set(value As Integer)
                If _rotationAngle <> value Then
                    _rotationAngle = value
                    Me.Invalidate() ' გადავხატოთ კონტროლი
                End If
            End Set
        End Property

        ''' <summary>
        ''' ლეიბლის ტექსტი
        ''' </summary>
        Public Overrides Property Text As String
            Get
                Return _text
            End Get
            Set(value As String)
                If _text <> value Then
                    _text = value
                    Me.Invalidate() ' გადავხატოთ კონტროლი
                End If
            End Set
        End Property

        ''' <summary>
        ''' კონსტრუქტორი
        ''' </summary>
        Public Sub New()
            Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
            Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            Me.SetStyle(ControlStyles.UserPaint, True)
            Me.SetStyle(ControlStyles.ResizeRedraw, True)
        End Sub

        ''' <summary>
        ''' კონტროლის მოხატვის მეთოდი
        ''' </summary>
        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)

            If String.IsNullOrEmpty(_text) Then
                Return
            End If

            ' ფონის შევსება
            Using brush As New SolidBrush(Me.BackColor)
                e.Graphics.FillRectangle(brush, Me.ClientRectangle)
            End Using

            ' ანტიალიასინგის ჩართვა (უფრო გლუვი ტექსტისთვის)
            e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            ' დავამატოთ ტრანსფორმაცია (ბრუნვა)
            e.Graphics.TranslateTransform(Me.Width / 2, Me.Height / 2)
            e.Graphics.RotateTransform(_rotationAngle)

            ' გავზომოთ ტექსტის ზომა
            Dim textSize = e.Graphics.MeasureString(_text, Me.Font)

            ' გამოვთვალოთ ტექსტის პოზიცია - ცენტრირებული
            Dim x = -textSize.Width / 2
            Dim y = -textSize.Height / 2

            ' დავხატოთ ტექსტი
            Using brush As New SolidBrush(Me.ForeColor)
                e.Graphics.DrawString(_text, Me.Font, brush, x, y)
            End Using

            ' დავაბრუნოთ ტრანსფორმაცია საწყის მდგომარეობაში
            e.Graphics.ResetTransform()
        End Sub
    End Class
End Namespace