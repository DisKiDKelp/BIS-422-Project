Public Class Retrieve
    Private DB As New DBAccessClass

    Private Sub Retrieve_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DB.ExecuteQuery("SELECT DISTINCT Model FROM Vehicle WHERE AvailableYN = 'Y' ORDER BY Model")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        modelComboBox.DataSource = DB.DBDataTable
        modelComboBox.DisplayMember = "Model"
        modelComboBox.Text = ""

        DB.ExecuteQuery("SELECT DISTINCT Model FROM Vehicle WHERE AvailableYN = 'Y' ORDER BY Model")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        modelSearchComboBox.DataSource = DB.DBDataTable
        modelSearchComboBox.DisplayMember = "Model"
        modelSearchComboBox.Text = ""
        DB.ExecuteQuery("SELECT DISTINCT Year FROM Vehicle WHERE AvailableYN = 'Y' ORDER BY Year")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        yearComboBox.DataSource = DB.DBDataTable
        yearComboBox.DisplayMember = "Year"
        yearComboBox.Text = ""

        DB.ExecuteQuery("SELECT DISTINCT Year FROM Vehicle WHERE AvailableYN = 'Y' ORDER BY Year")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        yearSearchComboBox.DataSource = DB.DBDataTable
        yearSearchComboBox.DisplayMember = "Year"
        yearSearchComboBox.Text = ""


        DB.ExecuteQuery("SELECT * FROM Vehicle Order BY VehicleID")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        BuildingDataGridView.DataSource = DB.DBDataTable

        updateButton.Enabled = False
    End Sub

    Private Sub makeSearchTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles makeSearchTextBox.TextChanged
        DB.AddParam("@Make", makeSearchTextBox.Text & "%")
        DB.AddParam("@Model", modelSearchComboBox.Text & "%")
        DB.AddParam("@Year", yearSearchComboBox.Text & "%")

        DB.ExecuteQuery("SELECT * FROM Vehicle WHERE Make LIKE ? AND Model LIKE ? and Year LIKE ? ORDER BY VehicleID")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        BuildingDataGridView.DataSource = DB.DBDataTable

        IDTextBox.Text = ""
        makeTextBox.Text = ""
        modelComboBox.Text = ""
        modelSearchComboBox.Text = ""
        yearSearchComboBox.Text = ""
        yearComboBox.Text = ""

        updateButton.Enabled = False
    End Sub

    Private Sub modelSearchComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles modelSearchComboBox.SelectedIndexChanged
        DB.AddParam("@Make", makeSearchTextBox.Text & "%")
        DB.AddParam("@Model", modelSearchComboBox.Text & "%")
        DB.AddParam("@Year", yearSearchComboBox.Text & "%")

        DB.ExecuteQuery("SELECT * FROM Vehicle WHERE Make LIKE ? AND Model LIKE ? and Year LIKE ? ORDER BY VehicleID")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        BuildingDataGridView.DataSource = DB.DBDataTable

        IDTextBox.Text = ""
        makeTextBox.Text = ""
        modelComboBox.Text = ""
        modelSearchComboBox.Text = ""
        yearSearchComboBox.Text = ""
        yearComboBox.Text = ""

        updateButton.Enabled = False
    End Sub
    Private Sub yearSearchComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles yearSearchComboBox.SelectedIndexChanged
        DB.AddParam("@Make", makeSearchTextBox.Text & "%")
        DB.AddParam("@Model", modelSearchComboBox.Text & "%")
        DB.AddParam("@Year", yearSearchComboBox.Text & "%")

        DB.ExecuteQuery("SELECT * FROM Vehicle WHERE Make LIKE ? AND Model LIKE ? and Year LIKE ? ORDER BY VehicleID")

        If DB.ErrorFlag Then
            Exit Sub
        End If

        BuildingDataGridView.DataSource = DB.DBDataTable

        IDTextBox.Text = ""
        modelSearchComboBox.Text = ""
        yearSearchComboBox.Text = ""
        makeTextBox.Text = ""
        modelComboBox.Text = ""
        yearComboBox.Text = ""

        updateButton.Enabled = False
    End Sub
    Private Sub BuildingDataGridView_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles BuildingDataGridView.CellClick
        If e.RowIndex > -1 Then
            IDTextBox.Text = BuildingDataGridView.Rows(e.RowIndex).Cells(0).Value()
            makeTextBox.Text = BuildingDataGridView.Rows(e.RowIndex).Cells(1).Value()
            modelComboBox.Text = BuildingDataGridView.Rows(e.RowIndex).Cells(2).Value()
            yearComboBox.Text = BuildingDataGridView.Rows(e.RowIndex).Cells(3).Value()


            updateButton.Enabled = True
        End If
    End Sub

    Private Sub updateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles updateButton.Click
        If makeTextBox.Text = "" Then
            MessageBox.Show("Please enter a make for the car.", "Missing Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            makeTextBox.Focus()
        ElseIf modelComboBox.Text = "" Then
            MessageBox.Show("Please select a model for the car.", "Missing Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            modelComboBox.Focus()
        ElseIf yearComboBox.Text = "" Then
            MessageBox.Show("Please select a year for the car.", "Missing Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            yearComboBox.Focus()

        Else
            DB.AddParam("@Make", makeTextBox.Text)
            DB.AddParam("@Model", modelComboBox.Text)
            DB.AddParam("@Year", yearComboBox.Text)
            DB.AddParam("@VehicleID", IDTextBox.Text)

            DB.ExecuteNonQuery("UPDATE Vehicle SET Make = ?, Model = ?, Year = ? WHERE VehicleID = ?")

            IDTextBox.Text = ""
            makeTextBox.Text = ""
            modelComboBox.Text = ""
            yearComboBox.Text = ""
            makeSearchTextBox.Text = ""
            modelSearchComboBox.Text = ""
            yearSearchComboBox.Text = ""

            DB.ExecuteQuery("SELECT * FROM Vehicle Order BY VehicleID")

            If DB.ErrorFlag Then
                Exit Sub
            End If

            BuildingDataGridView.DataSource = DB.DBDataTable

            updateButton.Enabled = False
        End If
    End Sub

    Private Sub quitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles quitButton.Click
        IDTextBox.Text = ""
        makeTextBox.Text = ""
        modelComboBox.Text = ""
        yearComboBox.Text = ""
        makeSearchTextBox.Text = ""
        modelSearchComboBox.Text = ""
        yearSearchComboBox.Text = ""

        Me.Close()
    End Sub


End Class
