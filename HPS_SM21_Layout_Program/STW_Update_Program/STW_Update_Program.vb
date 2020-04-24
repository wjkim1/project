Imports Scripting


Public Class STW_Update_Program
    Public Const oLayoutParogramName As String = "SM21_Layout_Program.exe"
    Public Const oTempFolderPath As String = "C:\SM21_Exe_File"
    Public Const oServerFolderPath As String = "Z:\WORK_REF\Layout_Program\SM21_Exe_File"


    Private Sub STW_Update_Program_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- Form 최상위
        Me.Top = 150
        Me.Left = 150
        Me.Width = 378
        Me.Height = 119
        Me.TopMost = True

        Me.Show()

        Label1.Refresh()
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0

        Try
            LBL_State.Text = "레이아웃 프로그램 업데이트 중..."
            LBL_State.Refresh()

            '--------------------------------------------------------------------------------------------------- 스무스한 진행률 표시를 위해..
            For i = 1 To 30
                ProgressBar1.Value = ProgressBar1.Value + 1
                ProgressBar1.Refresh()
                Threading.Thread.Sleep(50)
            Next

            Dim processes = GetObject("winmgmts:").execquery("select * from win32_process where name='" & oLayoutParogramName & "'")
            If Not processes.Count = 0 Then
                For Each process In processes
                    process.Terminate()
                Next
            End If

            '--------------------------------------------------------------------------------------------------- 스무스한 진행률 표시를 위해..
            For i = 1 To 30
                ProgressBar1.Value = ProgressBar1.Value + 1
                ProgressBar1.Refresh()
                Threading.Thread.Sleep(50)
            Next

            '--------------------------------------------------------------------------------------------------- TEMP 폴더로 프로그램 복사
            Dim fso As FileSystemObject
            fso = New FileSystemObject

            If fso.FolderExists(oTempFolderPath) = False Then
                '--------------------------------------------------------------------------------------------------- 폴더가 없으면 생성
                'fso.CopyFolder("D:\STW", oTempFolderPath)              ' 테스트 경로
                fso.CopyFolder(oServerFolderPath, oTempFolderPath)
                '--------------------------------------------------------------------------------------------------- 스무스한 진행률 표시를 위해..
                For i = 1 To 40
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    ProgressBar1.Refresh()
                    Threading.Thread.Sleep(50)
                Next
            Else
                '--------------------------------------------------------------------------------------------------- 폴더안에 파일 개수
                Dim FName = Dir(oTempFolderPath & "\*.*", vbNormal)  '목록을 생성해서 파일명을 가져옵니다.
                Dim iNum As Integer = 0
                Dim asda(40)

                Do While FName <> Nothing
                    If Not Strings.Left(FName, 10) = "STW_Update" Then
                        fso.GetFile(oTempFolderPath & "\" & FName).Delete()
                    End If

                    FName = Dir()
                Loop

                'fso.GetFolder(oTempFolderPath).Delete()
                '--------------------------------------------------------------------------------------------------- 스무스한 진행률 표시를 위해..
                For i = 1 To 20
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    ProgressBar1.Refresh()
                    Threading.Thread.Sleep(50)
                Next

                '--------------------------------------------------------------------------------------------------- 폴더가 없으면 생성
                'fso.CopyFolder("D:\STW", oTempFolderPath)              ' 테스트 경로
                fso.CopyFolder(oServerFolderPath, oTempFolderPath)
                '--------------------------------------------------------------------------------------------------- 스무스한 진행률 표시를 위해..
                For i = 1 To 20
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    ProgressBar1.Refresh()
                    Threading.Thread.Sleep(50)
                Next
            End If

            LBL_State.Text = "레이아웃 프로그램 업데이트 완료!"
            LBL_State.Refresh()

            Shell(oTempFolderPath & "\" & oLayoutParogramName, vbNormalFocus)
            '--------------------------------------------------------------------------------------------------- 프로그램 종료
            End
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            MessageBox.Show("프로그램 업데이트 중 오류발생!!", "ISPark Automation", MessageBoxButtons.OK, MessageBoxIcon.Error, _
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            '--------------------------------------------------------------------------------------------------- 프로그램 종료
            End
        End Try
    End Sub

End Class