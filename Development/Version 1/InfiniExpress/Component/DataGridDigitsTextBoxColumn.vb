Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.Windows.Forms

Namespace DataGridNumbersOnly
    Public Class DataGridDigitsTextBoxColumn
        Inherits DataGridTextBoxColumn
        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New(ByVal pd As PropertyDescriptor, ByVal format As String, ByVal b As Boolean)
            MyBase.New(pd, format, b)
            AddHandler Me.TextBox.KeyPress, New System.Windows.Forms.KeyPressEventHandler(AddressOf HandleKeyPress)

        End Sub
        
        Private Sub HandleKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)

            'ignore if not digit or control key
            'ignore if more than 3 digits
            If (Not (System.Char.IsDigit(e.KeyChar)) AndAlso Not (System.Char.IsControl(e.KeyChar))) Then
                e.Handled = True
            End If
            If ((Me.TextBox.Text.Length >= 3) _
                        AndAlso Not (System.Char.IsControl(e.KeyChar)) AndAlso Me.TextBox.SelectionLength = 0) Then
                e.Handled = True
            End If

        End Sub
    End Class
End Namespace
