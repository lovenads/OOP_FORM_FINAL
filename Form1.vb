Imports System.Data.OleDb
Imports System.Globalization
Imports System.Text.RegularExpressions
Public Class Form1
    Dim connStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\crudDB.mdb"
    Dim conn As New OleDbConnection(connStr)
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Function validateInputs() As Boolean
        Dim letterOnlyPattern As String = "^[a-zA-z ]+$"
        Dim last_nameRegex As New Regex(letterOnlyPattern)
        Dim first_nameRegex As New Regex(letterOnlyPattern)
        Dim middleInitialRegex As New Regex(letterOnlyPattern)
        Dim place_birthRegex As New Regex(letterOnlyPattern)
        Dim nationalityRegex As New Regex(letterOnlyPattern)
        Dim religionRegex As New Regex(letterOnlyPattern)
        Dim cityRegex As New Regex(letterOnlyPattern)

        Dim middleInitialInputs As String = txtMI.Text.Trim()

        Dim contact_numberPattern As String = "^09\d{9}$"
        Dim contact_numberRegex As New Regex(contact_numberPattern)

        Dim numberPattern As String = "^[0-9]+$"
        Dim ageRegex As New Regex(numberPattern)

        Dim telnumberPattern As String = "^((\(?0\d{2,3}\)?[- ]?)?\d{3}[- ]?\d{4}$"
        Dim telnumberRegex As New Regex(telnumberPattern)
        Dim telNumberInputs As String = txtTN.Text.Trim()

        Dim emailPattern As String = "^[^@\s]+@[^@\s]+\.[^@\s]+$"
        Dim emailRegex As New Regex(emailPattern)

        Dim houseNumberPattern As String = "^[0-9]{4}$"
        Dim houseNumberRegex As New Regex(houseNumberPattern)

        Dim blockPattern As String = "^(Blk|Block)\s?\d+$"
        Dim blockRegex As New Regex(blockPattern)
        Dim blockInputs As String = txtBLOCK.Text.Trim()

        Dim streetPattern As String = "^[a-zA-z0-9 .]+$"
        Dim streetRegex As New Regex(streetPattern)
        Dim streetInputs As String = txtSTREET.Text.Trim()

        Dim brgyRegex As New Regex(numberPattern)

        Dim cityInputs As String = txtCITY.Text.Trim()

        Dim postalPattern As String = "^[0-9]{4}$"
        Dim postalRegex As New Regex(postalPattern)
        Dim postalInputs As String = txtPOSTAL.Text.Trim()

        If Not contact_numberRegex.IsMatch(txtCN.Text) Then
            MessageBox.Show("Invalid Contact Number. It must start with 09 and be 11 digits long.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        If telNumberInputs <> "" Then
            If Not telnumberRegex.IsMatch(txtTN.Text) Then
                MessageBox.Show("Invalid Telephone Number. It must be XXX-XXXX", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
        End If

        If Not last_nameRegex.IsMatch(txtLN.Text) Then
            MessageBox.Show("Invalid last name format. It must be letter only and cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtLN.Focus()
            Return False
        End If

        If Not first_nameRegex.IsMatch(txtFN.Text) Then
            MessageBox.Show("Invalid first name format. It must be letter only and cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFN.Focus()
            Return False
        End If

        If middleInitialInputs <> "" Then
            If Not middleInitialRegex.IsMatch(txtMI.Text) Then
                MessageBox.Show("Invalid first name format. It must be letter only and cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtMI.Focus()
                Return False
            End If
        End If

            If Not place_birthRegex.IsMatch(txtPB.Text) Then
            MessageBox.Show("Invalid place of birth format. It must be letter only and cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPB.Focus()
            Return False
        End If

        If Not ageRegex.IsMatch(txtAGE.Text) Then
            MessageBox.Show("Invalid age format. It must be number only and cannot be empty", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAGE.Focus()
            Return False
        End If


        If Not emailRegex.IsMatch(txtEA.Text.Trim()) Then
            MessageBox.Show("Invalid Email format. And also cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtEA.Focus()
            Return False
        End If

        If Not houseNumberRegex.IsMatch(txtHN.Text.Trim()) Then
            MessageBox.Show("Invalid House Number. It must be 1 to 4 digits only.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtHN.Focus()
            Return False
        End If

        If Not nationalityRegex.IsMatch(txtNAT.Text.Trim()) Then
            MessageBox.Show("Invalid Nationality format. It must be letter only and cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNAT.Focus()
            Return False
        End If

        If blockInputs <> "" Then
            If Not blockRegex.IsMatch(txtBLOCK.Text.Trim()) Then
                MessageBox.Show("Invalid Block format. Use block word first and block number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtBLOCK.Focus()
                Return False
            End If
        End If

        If streetInputs <> "" Then
            If Not streetRegex.IsMatch(txtSTREET.Text) Then
                MessageBox.Show("Invalid Street format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtBLOCK.Focus()
                Return False
            End If
        End If

        If Not brgyRegex.IsMatch(txtBARANGAY.Text) Then
            MessageBox.Show("Invalid barangay format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBARANGAY.Focus()
            Return False
        End If

        If cityInputs <> "" Then
            If Not cityRegex.IsMatch(txtCITY.Text) Then
                MessageBox.Show("Invalid city format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCITY.Focus()
                Return False
            End If
        End If

        If postalInputs <> "" Then
            If Not postalRegex.IsMatch(txtPOSTAL.Text) Then
                MessageBox.Show("Invalid postal format.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPOSTAL.Focus()
                Return False
            End If
        End If

        If rbFEMALE.Checked = False And rbMALE.Checked = False Then
            MessageBox.Show("Gender cannot be empty", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            rbMALE.Focus()
            rbFEMALE.Focus()
            Return False
        End If

        Return True
    End Function

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
        brthDate.MaxDate = Date.Today.AddDays(-1)
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
        'txtCS.Clear()
        csCbo.SelectedIndex = -1
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
        Dim birthDate As Date = brthDate.Value
        Dim today As Date = Date.Today

        If birthDate >= today Then
            MessageBox.Show("Birth Date cannot be today or a future date.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            brthDate.Focus()
            Exit Sub
        End If

        If Not validateInputs() Then
            Exit Sub
        End If

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
            cmd.Parameters.AddWithValue("@Civil_status", csCbo.Text)
            cmd.Parameters.AddWithValue("@Nationality", txtNAT.Text)
            cmd.Parameters.AddWithValue("@Religion", txtRELIGION.Text)
            cmd.Parameters.AddWithValue("@Email", txtEA.Text)
            cmd.Parameters.AddWithValue("@Contact_number", txtCN.Text)
            cmd.Parameters.AddWithValue("@Tel", txtTN.Text)
            cmd.Parameters.AddWithValue("@Address", addressStr)
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Record Inserted Successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub

    Private Sub btnEDIT_Click(sender As Object, e As EventArgs) Handles btnUPDATE.Click
        Dim birthDate As Date = brthDate.Value
        Dim today As Date = Date.Today

        If birthDate >= today Then
            MessageBox.Show("Birth Date cannot be today or a future date.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            brthDate.Focus()
            Exit Sub
        End If

        If Not validateInputs() Then
            Exit Sub
        End If

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
                cmd.Parameters.AddWithValue("@Civil_status", csCbo.Text)
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

    Private Sub resetBtn_Click(sender As Object, e As EventArgs) Handles resetBtn.Click
        Try
            conn.Open()
            cmd = New OleDbCommand("DELETE FROM crud", conn)
            cmd.ExecuteNonQuery()
            cmd = New OleDbCommand("ALTER TABLE crud ALTER COLUMN ID COUNTER (1,1)", conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Record Reset Successfully")
            LoadData()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub
End Class
