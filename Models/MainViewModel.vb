' ===========================================
' 📄 Models/MainViewModel.vb
' -------------------------------------------
' MVVM-ish ViewModel: ლოგიკა და UI-ის State (Email, Role, IsAuthorized)
' ===========================================
Imports System.ComponentModel

Namespace Scheduler_v8_8a.Models
    ''' <summary>
    ''' MainViewModel ახორციელებს INotifyPropertyChanged და ინახავს UI-ის მდგომარეობას
    ''' Properties: Email, Role, IsAuthorized
    ''' </summary>
    Public Class MainViewModel
        Implements INotifyPropertyChanged

        Private _email As String = String.Empty
        Private _role As String = String.Empty
        Private _isAuthorized As Boolean = False

        ''' <summary>მომხმარებლის Email</summary>
        Public Property Email As String
            Get
                Return _email
            End Get
            Set(value As String)
                If _email <> value Then
                    _email = value
                    OnPropertyChanged(NameOf(Email))
                End If
            End Set
        End Property

        ''' <summary>მომხმარებლის როლი</summary>
        Public Property Role As String
            Get
                Return _role
            End Get
            Set(value As String)
                If _role <> value Then
                    _role = value
                    OnPropertyChanged(NameOf(Role))
                End If
            End Set
        End Property

        ''' <summary>True თუ ავტორიზებულია</summary>
        Public Property IsAuthorized As Boolean
            Get
                Return _isAuthorized
            End Get
            Set(value As Boolean)
                If _isAuthorized <> value Then
                    _isAuthorized = value
                    OnPropertyChanged(NameOf(IsAuthorized))
                End If
            End Set
        End Property

        ''' <inheritdoc/>
        Public Event PropertyChanged As PropertyChangedEventHandler _
            Implements INotifyPropertyChanged.PropertyChanged

        ''' <summary>
        ''' Fires when property changes
        ''' </summary>
        Protected Sub OnPropertyChanged(propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub
    End Class
End Namespace

