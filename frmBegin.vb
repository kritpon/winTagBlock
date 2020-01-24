Public Class frmBegin
    Dim strLabel As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fmtListView()
        DBtools.openDB()
    End Sub
    Function chkPCactive(strDocNo As String) As Integer

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select DISTINCT Trh_Type,Trh_KeyType,Trh_Date,Trh_No,Dtl_idTrade,Dtl_n_trade,Dtl_Date_S,Dtl_Date_F,"
        txtSQL = txtSQL & "Trh_DateSection,Trh_chk_Print, "
        txtSQL = txtSQL & "Dtl_Num,Dtl_Num_2  "

        txtSQL = txtSQL & "From TranDataH  "
        txtSQL = txtSQL & "left Join TranDataD "
        txtSQL = txtSQL & "On (TranDataH.Trh_Type=TranDataD.Dtl_Type And TranDataH.Trh_NO=TranDataD.Dtl_No ) "
        'txtsql=txtsql & "On"
        txtSQL = txtSQL & "Where Trh_type='M' "
        'txtSQL = txtSQL & "And Trh_KeyType='" & txtFindSection.Text & "' "
        txtSQL = txtSQL & "And Trh_No='" & strDocNo & "' "  '  เป็นวันที่จาก  Clock_Update

        'txtSQL = txtSQL & "Order by trh_date desc "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "data")

        If subDS.Tables("data").Rows.Count > 0 Then
            lbDocNo.Text = strDocNo
            lbStkName.Text = subDS.Tables("data").Rows(0).Item("Dtl_N_Trade")
            lbQty.Text = subDS.Tables("data").Rows(0).Item("Dtl_Num")
        End If


    End Function
    Function getDocNo() As String

        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        Dim strAns As String = ""

        txtSQL = "Select * "
        txtSQL = txtSQL & "From BOMmastF "

        txtSQL = txtSQL & "Where year(BOM_RM_Update)=" & Year(Now) - 543 & " "
        txtSQL = txtSQL & "And BOM_RM_Scales='5' "
        txtSQL = txtSQL & "And month(BOM_RM_Update)=" & Month(Now) & " "

        '   --where (bom_RM_Update)='2019-02-25'
        txtSQL = txtSQL & "Order by BOM_RM_Updatetime desc"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dataList")

        If subDS.Tables("dataList").Rows.Count > 0 Then
            With subDS.Tables("dataList").Rows(0)

                strAns = .Item("BOM_No")

            End With
        End If
        Return strAns
    End Function
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
    Sub getstrDocNo()
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Dim anydata() As String
        Dim lvi As New ListViewItem

        Dim strDocNO As String
        Dim strstartdate As String
        Dim strEndDate As String
        Dim strTagNo As String

        strDocNO = getDocNo()
        lbDocNo.Text = strDocNO
        chkPCactive(strDocNO)

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TagTranH "
        txtSQL = txtSQL & "Where Tag_DocNo='" & strDocNO & "' "
        txtSQL = txtSQL & "Order by Tag_End_Date "

        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "dataList")
        lsvWork.Items.Clear()


        For i = 0 To subDs.Tables("dataList").Rows.Count - 1

            With subDs.Tables("dataList").Rows(i)
                strstartdate = Format(.Item("Tag_Start_Date"), "HH:MM:ss")
                strEndDate = Format(.Item("Tag_End_Date"), "HH:MM:ss")
                strDocNO = .Item("Tag_DocNO")
                strTagNo = .Item("Tag_Number")
                anydata = New String() {(i + 1).ToString, strTagNo, strstartdate, strEndDate}
                lvi = New ListViewItem(anydata)
                lsvWork.Items.Add(lvi)
            End With


        Next

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        lbClock.Text = Format(Now, "HH:mm:ss")
        Me.Refresh()
    End Sub

    Private Sub cmbExit_Click(sender As Object, e As EventArgs) Handles cmbExit.Click
        End
    End Sub

    Private Sub cmb1_Click(sender As Object, e As EventArgs) Handles cmb1.Click
        strLabel = strLabel & cmb1.Text
        lbTagNo.Text = strLabel
    End Sub

    Private Sub cmb2_Click(sender As Object, e As EventArgs) Handles cmb2.Click
        strLabel = strLabel & cmb2.Text
        lbTagNo.Text = strLabel
    End Sub

    Private Sub cmb3_Click(sender As Object, e As EventArgs) Handles cmb3.Click
        strLabel = strLabel & cmb3.Text
        lbTagNo.Text = strLabel

    End Sub

    Private Sub cmb4_Click(sender As Object, e As EventArgs) Handles cmb4.Click
        strLabel = strLabel & cmb4.Text
        lbTagNo.Text = strLabel

    End Sub

    Private Sub cmb5_Click(sender As Object, e As EventArgs) Handles cmb5.Click
        strLabel = strLabel & cmb5.Text
        lbTagNo.Text = strLabel

    End Sub

    Private Sub cmb6_Click(sender As Object, e As EventArgs) Handles cmb6.Click
        strLabel = strLabel & cmb6.Text
        lbTagNo.Text = strLabel

    End Sub

    Private Sub cmb7_Click(sender As Object, e As EventArgs) Handles cmb7.Click
        strLabel = strLabel & cmb7.Text
        lbTagNo.Text = strLabel

    End Sub

    Private Sub cmb8_Click(sender As Object, e As EventArgs) Handles cmb8.Click
        strLabel = strLabel & cmb8.Text
        lbTagNo.Text = strLabel
    End Sub

    Private Sub cmb9_Click(sender As Object, e As EventArgs) Handles cmb9.Click
        strLabel = strLabel & cmb9.Text
        lbTagNo.Text = strLabel
    End Sub

    Private Sub cmb0_Click(sender As Object, e As EventArgs) Handles cmb0.Click
        strLabel = strLabel & cmb0.Text
        lbTagNo.Text = strLabel
    End Sub

    Private Sub cmbStrDel_Click(sender As Object, e As EventArgs) Handles cmbStrDel.Click

        If strLabel = "" Then
            lbTagNo.Text = ""
        Else
            strLabel = Microsoft.VisualBasic.Left(strLabel, Len(strLabel) - 1)
            lbTagNo.Text = strLabel
        End If

    End Sub

    Private Sub cmbNewWork_Click(sender As Object, e As EventArgs) Handles cmbNewWork.Click

        Try
            getstrDocNo()
            lbNow.Text = Now

        Catch ex As Exception
            Me.Refresh()
            lbDocNo.Text = "--เกิดข้อผิดพลาด--"
        End Try

    End Sub

    Sub fmtListView()

        With lsvWork

            .Columns.Add("ลำดับ", 60, HorizontalAlignment.Center)

            .Columns.Add("เบอร์", 100, HorizontalAlignment.Center)
            .Columns.Add("เวลาเริ่ม", 150, HorizontalAlignment.Right) '1
            .Columns.Add("เวลาเสร็จ", 150, HorizontalAlignment.Right) '1
            .View = View.Details
            .GridLines = True

        End With

    End Sub

    Private Sub cmbStrOK_Click(sender As Object, e As EventArgs) Handles cmbStrOK.Click
        If lbTagNo.Text = "" Then
            MsgBox("หมายเลขตะแกงไม่ถูกต้อง", MsgBoxStyle.Critical, "แจ้งเตือน")
            Exit Sub
        End If
        Dim anydata() As String
            Dim lvi As ListViewItem

            Dim intRow As Integer = 0
            Dim strDocNo As String = lbDocNo.Text
            Dim strTag As String = lbTagNo.Text
            intRow = lsvWork.Items.Count + 1



        Dim strStartDate As String = lbNow.Text
        Dim strEndDate As String = Now

        anydata = New String() {intRow, strTag, strStartDate, strEndDate}
        lvi = New ListViewItem(anydata)
        lsvWork.Items.Add(lvi)

        txtSQL = "Insert into TagTranH(Tag_Start_Date,Tag_DocNo,Tag_End_Date,Tag_Number) "
            txtSQL = txtSQL & "Values('" & strToDate(strStartDate) & "','" & strDocNo & "','" & strToDate(strEndDate) & "','" & strTag & "')"
            'txtSQL = txtSQL & "'" & strToDate(docDate) & "','" & trhNo2 & "','" & running & "','0')"
            DBtools.dbSaveDATA(txtSQL, "")
            txtSQL = ""
            lbNow.Text = Now
            strLabel = ""
            lbTagNo.Text = ""

    End Sub

    Private Sub lsvWork_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsvWork.SelectedIndexChanged

    End Sub

    Private Sub lsvWork_Click(sender As Object, e As EventArgs) Handles lsvWork.Click
        Dim lvi As ListViewItem

        For i = 0 To lsvWork.SelectedItems.Count - 1


            lvi = lsvWork.SelectedItems(i)
            lbTagNo.Text = lsvWork.Items(lvi.Index).SubItems(2).Text


            lsvWork.Items.Remove(lvi)

        Next

    End Sub
End Class
