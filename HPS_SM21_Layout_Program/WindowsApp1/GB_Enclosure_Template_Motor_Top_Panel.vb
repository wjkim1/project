Public Class GB_Enclosure_Template_Motor_Top_Panel


    Private Sub GB_Enclosure_Template_Motor_Top_Panel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Top = 150
        Me.Left = 50
        Me.TopMost = True

        '--------------------------------------------------------------------------------------------------- 초기 Parameter 값 변경
        If TOP_MOTOR_PNL_OPTION.Value = "SPLIT" Then
            Enclosure_Remove_Option.Checked = True
        Else
            Enclosure_Extrusion_Option.Checked = True
        End If
    End Sub


    Private Sub Enclosure_Motor_Top_Panel_Update_Click(sender As Object, e As EventArgs) Handles Enclosure_Motor_Top_Panel_Update.Click
        Try
            Message_Label.Text = "Update 중....."
            '--------------------------------------------------------------------------------------------------- 흡음제 제거 선택
            If Enclosure_Remove_Option.Checked = True Then
                GA_Enclosure_Template_Create.Enclosure_Remove_Option.Checked = True
                TOP_MOTOR_PNL_OPTION.Value = "SPLIT"
            Else
                GA_Enclosure_Template_Create.Enclosure_Extrusion_Option.Checked = True
                TOP_MOTOR_PNL_OPTION.Value = "EXTRUDE"
            End If

            Call CATIA_BASIC02()
            CATIA.ActiveDocument.Product.Update()
            Message_Label.Text = "Update 완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)

        End Try
    End Sub

End Class