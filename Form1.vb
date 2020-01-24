Public Class frmBegin
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

        txtSQL = txtSQL & "Where Trh_type='M' "
        'txtSQL = txtSQL & "And Trh_KeyType='" & txtFindSection.Text & "' "
        txtSQL = txtSQL & "And Trh_No='" & strDocNo & "' "  '  เป็นวันที่จาก  Clock_Update
        'txtSQL = txtSQL & "Order by trh_date desc "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "data")

        If subDS.Tables("data").Rows.Count > 0 Then



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
        Dim strDocNO As String
        strDocNO = getDocNo()
        lbDocNo.Text = strDocNO
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        getstrDocNo()
    End Sub
End Class
