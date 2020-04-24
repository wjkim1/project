Imports System.ComponentModel
Imports System.Threading

Public Class Template_Data_Loading_Form
    Private trd As Thread
    Public Delegate Sub ThreadDelegate(ByVal Val As Short)
    Public Trd_Delegate As ThreadDelegate


    Private Sub Template_Data_Loading_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.TopMost = True
        Me.ProgressBar1.Maximum = 100
        'Me.ProgressBar1.Value = 100

        '--------------------------------------------------------------------------------------------------- 선정 및 템플릿 생성시 비교 변수 초기화
        For i = 0 To 20
            get_parameter_first_value(i) = ""
            get_parameter_second_value(i) = ""
        Next i

        If TypeName(trd).ToString() = "Nothing" Then
            trd = New Thread(AddressOf ThreadTask)

        End If

        trd.IsBackground = True
        If trd.IsAlive = False Then
            trd.Start() '실행
            'trd.Abort() '중단
            'trd.Suspend() '일시 중단
            'trd.Resume()    '재실행
            'trd.Join()      '종료 또는 중단을 기다리는 상태
            'Thread.Sleep(100)
        End If

    End Sub


    Private Sub ThreadTask()
        Dim Val As Short
        Trd_Delegate = New ThreadDelegate(AddressOf ProgressBar1_Change)
        Val = 0
        Do
            Me.Invoke(Trd_Delegate, Val)
            Val = Val + 1
            If Val > 100 Then
                Val = 0
            End If
            Thread.Sleep(50)
        Loop
    End Sub

    Public Sub ProgressBar1_Change(ByVal Val As Short)
        Me.ProgressBar1.Value = Val
        Me.Refresh()
    End Sub

    Private Sub Template_Data_Loading_Form_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            trd.Abort()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try
    End Sub
End Class