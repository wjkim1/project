Public Class AA_Update_Form
    ' Const sUpdateProgramPath As String = "d:\stw\"         ' 업데이트 프로그램이 있는 서버 경로
    Public Const sUpdateProgramPath As String = "Z:\WORK_REF\Layout_Program\SM21_Exe_File\"         ' 업데이트 프로그램이 있는 서버 경로
    Public Const sUpdateProgramName As String = "STW_Update_Program.exe"                            ' 업데이트 프로그램 파일명

    Private Sub AA_Update_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        '--------------------------------------------------------------------------------------------------- Form 최상위
        Me.Top = 150
        Me.Left = 50
        Me.Width = 611
        Me.Height = 326
        Me.TopMost = True

        Try
            '--------------------------------------------------------------------------------------------------- 업데이트 프로세스 킬
            Dim wmi = GetObject("winmgmts:")
            Dim sQuery = "select * from win32_process where name='" & sUpdateProgramName & "'"
            Dim processes = wmi.execquery(sQuery)

            If Not processes.Count = 0 Then
                For Each process In processes
                    process.Terminate()
                Next

            End If

            Comp_Layout_Update_Count = Comp_Layout_Update_Count + 1
            '-------------------------------------------------------------------------------------------------- 프로그램 Version Check (컴파일)
            'Call Program_Version_Check()
            'Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString)    
            'Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString)   '위를 아래와 같이 변경

            '--------------------------------------------------------------------------------------------------- Update txt 파일 Value 가져오기
            Dim Txt_Path As String
            If Comp_Layout_Update_Count = 1 Then
                '--------------------------------------------------------------------------------------------------- 기 생성된 Layout_Program 제거
                Call KillProcess_Layout_Program()

                Txt_Path = "Z:\WORK_REF\Layout_Program\SM21_Exe_File\Update.txt"
            Else
                Txt_Path = "Z:\WORK_REF\Layout_Program\SM21_Exe_File\Update_All.txt"
            End If

            '--------------------------------------------------------------------------------------------------- 업데이트 내역 파일 재 검색
            If System.IO.File.Exists(Txt_Path) = False Then
                If System.IO.File.Exists(Application.StartupPath & "\Update.txt") = True Then
                    Txt_Path = Application.StartupPath & "\Update.txt"
                End If
            End If

            Dim fso As Scripting.FileSystemObject
            fso = New Scripting.FileSystemObject

            If System.IO.File.Exists(Txt_Path) = True Then
                Dim oUpdateText As Scripting.TextStream
                oUpdateText = fso.OpenTextFile(Txt_Path)

                Do Until oUpdateText.AtEndOfStream                       ' 파일을 끝까지 읽을 때까지 루프
                    Me.Update_List.Items.Add(oUpdateText.ReadLine)       ' 한줄 읽음
                Loop

                oUpdateText.Close()
            Else
                Me.Update_List.Items.Add(" - Update 내역 파일이 없습니다.")
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("AA_Update_Form_Load() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Updata_OK_Click(sender As Object, e As EventArgs) Handles Updata_OK.Click
        If Version_Check_Value = Nothing Then
            Version_Check_Value = "New"
        End If

        '--------------------------------------------------------------------------------------------------- 서버 파일과 지금 실행 파일의 최종 수정일이 다르면..
        If Version_Check_Value = "Old" Then
            If SHOW_MESSAGE_BOX("프로그램을 업데이트 하시겠습니까?", "", 2, 1) = vbYes Then
                If System.IO.File.Exists(sUpdateProgramPath & sUpdateProgramName) = False Then
                    SHOW_MESSAGE_BOX("프로그램 업데이트 파일 오류로 종료합니다." & vbCrLf & "관리자에게 문의하세요.", "", 1, 1)
                    End
                End If

                Call Shell(sUpdateProgramPath & sUpdateProgramName, vbNormalFocus)
                End
            Else
                End
            End If
        Else
            '--------------------------------------------------------------------------------------------------- 프로그램 시작시에만 사용
            If Comp_Layout_Update_Count = 1 Then
                A_Main_Form.Show()
                'A_Main_Form1 = New A_Main_Form()
                'A_Main_Form1.Show()
                Me.Close()
            End If
        End If
    End Sub

    Private Sub Update_List_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Update_List.SelectedIndexChanged

    End Sub
End Class