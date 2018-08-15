Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Public Class Form1
 Public curWord
    Private Sub RichTextBox1_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseMove
         Label1.Text = e.X
        Label2.Text = e.Y
        GetWordUnderMouse(Me.RichTextBox1, e.X, e.Y)
        Timer1.Start()
    End Sub
    Public Function GetWordUnderMouse(ByRef Rtf As System.Windows.Forms.RichTextBox, ByVal X As Integer, ByVal Y As Integer) As String
        On Error Resume Next
        Dim POINT As System.Drawing.Point = New System.Drawing.Point()
        Dim Pos As Integer, i As Integer, lStart As Integer, lEnd As Integer
        Dim lLen As Integer, sTxt As String, sChr As String
        '
        POINT.X = X
        POINT.Y = Y
        GetWordUnderMouse = vbNullString
        '
        With Rtf
            lLen = Len(.Text)
            sTxt = .Text
            Pos = Rtf.GetCharIndexFromPosition(POINT)
            If Pos > 0 Then
                For i = Pos To 1 Step -1
                    sChr = Mid(sTxt, i, 1)
                    If sChr = " " Or sChr = Chr(13) Or i = 1 Then
                        'if the starting character is vbcrlf then
                        'we want to chop that off
                        If sChr = Chr(13) Then
                            lStart = (i + 2)
                        Else
                            lStart = i
                        End If
                        Exit For
                    End If
                Next i
                For i = Pos To lLen
                    If Mid(sTxt, i, 1) = " " Or Mid(sTxt, i, 1) = Chr(13) Or i = lLen Then
                        lEnd = i + 1
                        Exit For
                    End If
                Next i
                If lEnd >= lStart Then
                    GetWordUnderMouse = Trim(Mid(sTxt, lStart, lEnd - lStart))
                End If
            End If
        End With
    End Function
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim h As Integer = Label1.Text
        Dim h2 As Integer = Label2.Text
        '  MsgBox(GetWordUnderMouse(Me.RichTextBox1, h, h2))
        TextBox1.Text = GetWordUnderMouse(Me.RichTextBox1, h, h2)
        TextBox2.Text = GetWordUnderMouse(Me.RichTextBox1, h, h2)
        Dim sapi
        sapi = CreateObject("sapi.spvoice")
        sapi.Speak(GetWordUnderMouse(Me.RichTextBox1, h, h2))

        Dim a As String
        Dim b As String
        Dim index_a As Integer
        Dim index_b As Integer
        a = TextBox1.Text
        b = TextBox2.Text
        index_a = InStr(RichTextBox1.Text, a)
        index_b = InStr(RichTextBox1.Text, b)
        If index_a And index_b Then
            RichTextBox1.Focus()
            RichTextBox1.SelectionStart = index_a - 1
            RichTextBox1.SelectionLength = (index_b - index_a) + Len(b)
        End If
        Me.Timer1.Enabled = False
    End Sub
End Class
