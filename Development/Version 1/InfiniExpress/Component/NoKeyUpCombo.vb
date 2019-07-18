Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Windows.Forms

Namespace DataGridTextBoxCombo
    Public Class NoKeyUpCombo
        Inherits ComboBox
        Private WM_KEYUP As Integer = &H101
        Public Shared shouldBeExecuted As Integer = 0
        Public Shared maintainRNo As Integer = Global.rowNumber


        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            If m.Msg = WM_KEYUP Then
                'ignore keyup to avoid problem with tabbing & dropdownlist;
                Return
            End If
            MyBase.WndProc(m)
        End Sub 'WndProc

        Private Sub combobox_selectedindexchanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SelectedIndexChanged

            If shouldBeExecuted = 1 Or Global.rowNumber = maintainRNo Then
                If MyBase.SelectedIndex = -1 Then
                    Global.taxRatesArray(Global.rowNumber) = 1
                Else
                    Global.taxRatesArray(Global.rowNumber) = MyBase.SelectedValue
                End If
                maintainRNo = Global.rowNumber
                shouldBeExecuted = 0
            End If

        End Sub
    End Class 'NoKeyUpCombo
End Namespace
