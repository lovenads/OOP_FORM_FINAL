Imports System.Data.OleDb
Public Class Form1
    Dim connStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\crudDB.mdb"
    Dim conn As New OleDbConnection(connStr)
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Sub LoadData()
        Try
            dt = New DataTable
            conn.Open()
            da = New OleDbDataAdapter("SELECT * FROM crud", conn)
            da.Fill(dt)
            DataGridView1.DataSource = dt
            conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles txtMI.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtLN.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles txtSTREET.TextChanged

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        LoadData()
    End Sub

    Private Sub btnCLEAR_Click(sender As Object, e As EventArgs) Handles btnCLEAR.Click
        txtLN.Clear()
        txtFN.Clear()
        txtMI.Clear()
        cmbSF.SelectedIndex = -1
        txtPB.Clear()
        txtAGE.Clear()
        rbFEMALE.Checked = False
        rbMALE.Checked = False
        txtCS.Clear()
        txtNAT.Clear()
        txtRELIGION.Clear()
        txtEA.Clear()
        txtCN.Clear()
        txtTN.Clear()
        txtHN.Clear()
        txtBLOCK.Clear()
        txtSTREET.Clear()
        txtBARANGAY.Clear()
        txtCITY.Clear()
        txtPROVINCE.Clear()
        txtPOSTAL.Clear()

    End Sub

    Private Sub addBtn_Click(sender As Object, e As EventArgs) Handles addBtn.Click
        Dim addressStr As String =
            txtHN.Text + ", " +
            txtBLOCK.Text + " " +
            txtSTREET.Text + " " +
            txtBARANGAY.Text + " " +
            txtCITY.Text + " " +
            txtPROVINCE.Text + " " +
            txtPOSTAL.Text + " " + "Phillipines"
        Dim genderStr As String = ""

        If rbMALE.Checked Then
            genderStr = "Male"

        ElseIf rbFEMALE.Checked Then
            genderStr = "Female"
        End If
        Try
            conn.Open()
            cmd = New OleDbCommand("INSERT INTO crud 
            ([Fname], [Lname], [MI], [Suffix], [Birth_place], [Birth_date], [Age], [Gender], [Civil_status], [Nationality], [Religion], [Email], [Contact_number], [Tel], [Address])
            VALUES (@Fname, @Lname, @MI, @Suffix, @Birth_place, @Birth_date, @Age, @Gender, @Civil_status, @Nationality, @Religion, @Email, @Contact_number, @Tel, @Address)", conn)
            cmd.Parameters.AddWithValue("@Fname", txtFN.Text)
            cmd.Parameters.AddWithValue("@Lname", txtLN.Text)
            cmd.Parameters.AddWithValue("@MI", txtMI.Text)
            cmd.Parameters.AddWithValue("@Suffix", cmbSF.Text)
            cmd.Parameters.AddWithValue("@Birth_place", txtPB.Text)
            cmd.Parameters.AddWithValue("@Birth_date", DateTime.Parse(brthDate.Text))
            cmd.Parameters.AddWithValue("@Age", Val(txtAGE.Text))
            cmd.Parameters.AddWithValue("@Gender", genderStr)
            cmd.Parameters.AddWithValue("@Civil_status", txtCS.Text)
            cmd.Parameters.AddWithValue("@Nationality", txtNAT.Text)
            cmd.Parameters.AddWithValue("@Religion", txtRELIGION.Text)
            cmd.Parameters.AddWithValue("@Email", txtEA.Text)
            cmd.Parameters.AddWithValue("@Contact_number", txtCN.Text)
            cmd.Parameters.AddWithValue("@Tel", txtTN.Text)
            cmd.Parameters.AddWithValue("@Address", addressStr)
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Record Inserted Successfully")
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub

    Private Sub btnEDIT_Click(sender As Object, e As EventArgs) Handles btnUPDATE.Click
        Try
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to update this record?",
                                                         "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            Dim addressStr As String = txtHN.Text + ", " +
            txtBLOCK.Text + " " +
            txtSTREET.Text + " " +
            txtBARANGAY.Text + " " +
            txtCITY.Text + " " +
            txtPROVINCE.Text + " " +
            txtPOSTAL.Text

            Dim genderStr As String = ""
            If rbMALE.Checked Then
                genderStr = "Male"
            ElseIf rbFEMALE.Checked Then
                genderStr = "Female"
            End If

            If result = DialogResult.Yes Then
                Dim id As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells(0).Value)
                conn.Open()
                cmd = New OleDbCommand("UPDATE crud set 
                        [Fname]=@Fname, 
                        [Lname]=@Lname, 
                        [MI]=@MI, 
                        [Suffix]=@Suffix, 
                        [Birth_place]=@Birth_place,
                        [Birth_date]=@Birth_date,
                        [Age]=@Age, 
                        [Gender]=@Gender, 
                        [Civil_status]=@Civil_status, 
                        [Nationality]=@Nationality, 
                        [Religion]=@Religion, 
                        [Email]=@Email, 
                        [Contact_number]=@Contact_number, 
                        [Tel]=@Tel, 
                        [Address]=@Address 
                        WHERE ID=@ID", conn)

                cmd.Parameters.AddWithValue("@Fname", txtFN.Text)
                cmd.Parameters.AddWithValue("@Lname", txtLN.Text)
                cmd.Parameters.AddWithValue("@MI", txtMI.Text)
                cmd.Parameters.AddWithValue("@Suffix", cmbSF.Text)
                cmd.Parameters.AddWithValue("@Birth_place", txtPB.Text)
                cmd.Parameters.AddWithValue("@Birth_date", DateTime.Parse(brthDate.Text))
                cmd.Parameters.AddWithValue("@Age", Val(txtAGE.Text))
                cmd.Parameters.AddWithValue("@Gender", genderStr)
                cmd.Parameters.AddWithValue("@Civil_status", txtCS.Text)
                cmd.Parameters.AddWithValue("@Nationality", txtNAT.Text)
                cmd.Parameters.AddWithValue("@Religion", txtRELIGION.Text)
                cmd.Parameters.AddWithValue("@Email", txtEA.Text)
                cmd.Parameters.AddWithValue("@Contact_number", txtCN.Text)
                cmd.Parameters.AddWithValue("@Tel", txtTN.Text)
                cmd.Parameters.AddWithValue("@Address", addressStr)
                cmd.Parameters.AddWithValue("@ID", id)
                cmd.ExecuteNonQuery()
                conn.Close()
                MessageBox.Show("Record Updated Successfully")
                LoadData()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub

    Private Sub btnDELETE_Click(sender As Object, e As EventArgs) Handles btnDELETE.Click
        Try
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this record?",
                                                         "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Dim id As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells(0).Value)
                conn.Open()
                cmd = New OleDbCommand("DELETE FROM crud WHERE ID=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", id)
                cmd.ExecuteNonQuery()
                conn.Close()
                MessageBox.Show("Record Deleted Successfully")
                LoadData()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub
End Class
