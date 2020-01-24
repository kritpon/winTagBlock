Public Class frmConnectDB
    Dim strConn1 As String = ""
    Dim strConn2 As String = ""
    Dim strConn3 As String = ""
    Dim strConn4 As String = ""

    Private Sub frmConnectDB_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dTB As New DataTable()
        Dim dRow As DataRow
        'Dim i As Integer
        strConn1 = DBConnString.strConn1
        strConn2 = DBConnString.strConn2
        strConn3 = DBConnString.strConn3
        strConn4 = DBConnString.strConn4

        dTB.Columns.Add(New DataColumn("DB_Name", GetType(String)))
        dTB.Columns.Add(New DataColumn("DB_Connect", GetType(String)))

        dRow = dTB.NewRow
        dRow(0) = "บริษัท แพนเอเซีย"
        dRow(1) = strConn1
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "บ.แพน เอเซีย"
        dRow(1) = strConn2
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "ผ่าน Internet"
        dRow(1) = strConn3
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "newZone"
        dRow(1) = strConn4
        dTB.Rows.Add(dRow)

        With cboDBlist
            .DataSource = dTB
            .DisplayMember = "DB_Name"
            .ValueMember = "DB_Connect"
        End With

        cboDBlist.SelectedIndex = 0
        strConn = cboDBlist.SelectedValue ' DBConnString.strConn2
        lbDBconnect.Text = strConn

        'With cboDBlist


        '    .Items.Add(strConn2)
        '    .Items.Add(strConn1)
        '    .SelectedItem = 1

        'End With

        lbTimer1.Text = 0
        Timer1.Enabled = True
    End Sub

    Private Sub cboDBlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDBlist.SelectedIndexChanged
        Try
            strConn = cboDBlist.SelectedValue ' DBConnString.strConn2
            lbDBconnect.Text = strConn
        Catch ex As Exception

        End Try




    End Sub
    Sub RunDB()
        Dim frmBegin As New frmBegin
        DBtools.openDB()
        With Conn
            If .State = ConnectionState.Open Then
                Me.Hide()
                frmBegin.Show()

            Else
                '  MsgBox("ไม่สามารถติดต่อฐานข้อมูลได้")
            End If

        End With

    End Sub
    Private Sub cmdConnect_Click(sender As Object, e As EventArgs) Handles cmdConnect.Click

        Call RunDB()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        lbTimer1.Text = lbTimer1.Text + 1
        If lbTimer1.Text = 1 Then
            'Call selectOK()
            Call RunDB()
            Me.Hide()
        End If

    End Sub
End Class