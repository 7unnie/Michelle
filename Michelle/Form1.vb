Public Class Form1
    ' PRELIM
    Dim q1_1, q2_1, q3_1, att_1, rec1_1, rec2_1, act1_1, act2_1, act3_1, exam1 As Double
    Dim tq1, tatt1, tre1, tact1, tex1, prelim As Double

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        q1_1 = Val(TextBox1.Text)
        q2_1 = Val(TextBox2.Text)
        q3_1 = Val(TextBox3.Text)
        att_1 = Val(TextBox4.Text)
        rec1_1 = Val(TextBox5.Text)
        rec2_1 = Val(TextBox6.Text)
        act1_1 = Val(TextBox7.Text)
        act2_1 = Val(TextBox8.Text)
        act3_1 = Val(TextBox9.Text)
        exam1 = Val(TextBox10.Text)

        tq1 = ((q1_1 + q2_1 + q3_1) / 3) * 0.25
        tatt1 = att_1 * 0.1
        tre1 = ((rec1_1 + rec2_1) / 2) * 0.1
        tact1 = ((act1_1 + act2_1 + act3_1) / 3) * 0.25
        tex1 = exam1 * 0.3

        prelim = (tq1 + tatt1 + tre1 + tact1 + tex1) * 0.3
        TextBox31.Text = prelim

    End Sub

    ' MIDTERM
    Dim q1_2, q2_2, q3_2, att_2, rec1_2, rec2_2, act1_2, act2_2, act3_2, exam2 As Double
    Dim tq2, tatt2, tre2, tact2, tex2, midterm As Double

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        q1_2 = Val(TextBox11.Text)
        q2_2 = Val(TextBox12.Text)
        q3_2 = Val(TextBox13.Text)
        att_2 = Val(TextBox14.Text)
        rec1_2 = Val(TextBox15.Text)
        rec2_2 = Val(TextBox16.Text)
        act1_2 = Val(TextBox17.Text)
        act2_2 = Val(TextBox18.Text)
        act3_2 = Val(TextBox19.Text)
        exam2 = Val(TextBox20.Text)

        tq2 = ((q1_2 + q2_2 + q3_2) / 3) * 0.25
        tatt2 = att_2 * 0.1
        tre2 = ((rec1_2 + rec2_2) / 2) * 0.1
        tact2 = ((act1_2 + act2_2 + act3_2) / 3) * 0.25
        tex2 = exam2 * 0.3

        midterm = (tq2 + tatt2 + tre2 + tact2 + tex2) * 0.3
        TextBox32.Text = midterm



    End Sub

    ' finals
    Dim q1_3, q2_3, q3_3, att_3, rec1_3, rec2_3, act1_3, act2_3, act3_3, exam3 As Double
    Dim tq3, tatt3, tre3, tact3, tex3, finals As Double

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        q1_3 = Val(TextBox21.Text)
        q2_3 = Val(TextBox22.Text)
        q3_3 = Val(TextBox23.Text)
        att_3 = Val(TextBox24.Text)
        rec1_3 = Val(TextBox25.Text)
        rec2_3 = Val(TextBox26.Text)
        act1_3 = Val(TextBox27.Text)
        act2_3 = Val(TextBox28.Text)
        act3_3 = Val(TextBox29.Text)
        exam3 = Val(TextBox30.Text)

        tq3 = ((q1_3 + q2_3 + q3_3) / 3) * 0.25
        tatt3 = att_3 * 0.1
        tre3 = ((rec1_3 + rec2_3) / 2) * 0.1
        tact3 = ((act1_3 + act2_3 + act3_3) / 3) * 0.25
        tex3 = exam3 * 0.3
        finals = (tq3 + tatt3 + tre3 + tact3 + tex3) * 0.4
        TextBox33.Text = finals


    End Sub
    ' total grade
    Dim tg As Double
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        tg = prelim + midterm + finals
        TextBox34.Text = tg
    End Sub
End Class
