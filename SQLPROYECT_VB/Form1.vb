Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Text
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.IO
Imports Newtonsoft.Json
Imports System.Xml



Namespace SQLPROYECT
    Partial Public Class Form1
        Inherits Form
        Private sql As SqlConnection
        Public Sub New()
            InitializeComponent()
            InitializeDatabaseConnection()
        End Sub
        Private Sub InitializeDatabaseConnection()
            Dim connectionString = "Server=KEILA\SQLEXPRESS01;Database=SQLPROYECT;Integrated Security=True;"
            sql = New SqlConnection(connectionString)
        End Sub
        Private Sub btnOpen_Click(sender As Object, e As EventArgs)
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog With {
    .Filter = "CSV files (.csv)|.csv|XML files (.xml)|.xml|JSON files (.json)|.json"
}
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath = openFileDialog.FileName
                Dim extension As String = Path.GetExtension(filePath).ToLower()
                Select Case extension
                    Case ".csv"
                        LoadCsv(filePath)
                    Case ".xml"
                        LoadXml(filePath)
                    Case ".json"
                        LoadJson(filePath)
                    Case Else
                        MessageBox.Show("Unsupported file format")
                End Select
            End If
        End Sub

        Private Sub SaveDataToDatabase()
            Try
                sql.Open()
                Dim sqlDataAdapter As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM STUDENTS", sql)
                Dim sqlCommandBuilder As SqlCommandBuilder = New SqlCommandBuilder(sqlDataAdapter)
                Dim dataTable = CType(DGV1.DataSource, DataTable)
                sqlDataAdapter.Update(dataTable)
            Catch ex As Exception
                MessageBox.Show("Error saving data: " & ex.Message)
            Finally
                sql.Close()
            End Try
        End Sub
        Private Sub btnSave_Click(sender As Object, e As EventArgs)
            SaveDataToDatabase()
        End Sub
        Private Sub LoadDataFromDatabase()
            Try
                sql.Open()
                Dim sqlDataAdapter As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM STUDENTS", sql)
                Dim dataTable As DataTable = New DataTable()
                sqlDataAdapter.Fill(dataTable)
                DGV1.DataSource = dataTable
            Catch ex As Exception
                MessageBox.Show("Error loading data: " & ex.Message)
            Finally
                sql.Close()
            End Try
        End Sub
        Private Sub btnMost_Click(sender As Object, e As EventArgs)
            LoadDataFromDatabase()
        End Sub
        Private Sub LoadCsv(filePath As String)
            Try
                Dim dataTable As DataTable = New DataTable()
                Using sr As StreamReader = New StreamReader(filePath)
                    Dim headers As String() = sr.ReadLine().Split(","c)
                    For Each header In headers
                        dataTable.Columns.Add(header)
                    Next
                    While Not sr.EndOfStream
                        Dim rows As String() = sr.ReadLine().Split(","c)
                        Dim dataRow As DataRow = dataTable.NewRow()
                        For i = 0 To headers.Length - 1
                            dataRow(i) = rows(i)
                        Next
                        dataTable.Rows.Add(dataRow)
                    End While
                End Using
                DGV1.DataSource = dataTable
            Catch ex As Exception
                MessageBox.Show("Error loading CSV: " & ex.Message)
            End Try
        End Sub
        Private Sub LoadXml(filePath As String)
            Try
                Dim dataSet As DataSet = New DataSet()
                dataSet.ReadXml(filePath)
                DGV1.DataSource = dataSet.Tables(0)
            Catch ex As Exception
                MessageBox.Show("Error loading XML: " & ex.Message)
            End Try
        End Sub
        Private Sub LoadJson(filePath As String)
            Try
                Dim jsonData = File.ReadAllText(filePath)
                Dim dataTable = JsonConvert.DeserializeObject(Of DataTable)(jsonData)
                DGV1.DataSource = dataTable
            Catch ex As Exception
                MessageBox.Show("Error loading JSON: " & ex.Message)
            End Try
        End Sub

        Private Sub SaveToXml(filePath As String)
            Try
                Dim dataTable = CType(DGV1.DataSource, DataTable)
                If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
                    MessageBox.Show("No data to save!")
                    Return
                End If

                Dim dataSet As DataSet = New DataSet()
                dataSet.Tables.Add(dataTable.Copy())
                dataSet.WriteXml(filePath)

                MessageBox.Show("Data saved successfully to XML: " & filePath)
            Catch ex As Exception
                MessageBox.Show("Error saving to XML: " & ex.Message)
            End Try
        End Sub
        Private Sub SaveToCsv(filePath As String)
            Try
                Dim dataTable = CType(DGV1.DataSource, DataTable)
                If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
                    MessageBox.Show("No data to save!")
                    Return
                End If

                Dim csvData As StringBuilder = New StringBuilder()

                ' Write header row
                For Each column As DataColumn In dataTable.Columns
                    csvData.Append(column.ColumnName & ",")
                Next
                csvData.Remove(csvData.Length - 1, 1) ' Remove trailing comma
                csvData.AppendLine()

                ' Write data rows
                For Each row As DataRow In dataTable.Rows
                    For i = 0 To row.ItemArray.Length - 1
                        Dim cellValue As String = row.ItemArray(i).ToString()
                        ' Escape special characters for CSV (optional)
                        ' cellValue = cellValue.Replace(",", "\",");
                        csvData.Append(cellValue & ",")
                    Next
                    csvData.Remove(csvData.Length - 1, 1) ' Remove trailing comma
                    csvData.AppendLine()
                Next

                Call File.WriteAllText(filePath, csvData.ToString())

                MessageBox.Show("Data saved successfully to CSV: " & filePath)
            Catch ex As Exception
                MessageBox.Show("Error saving to CSV: " & ex.Message)
            End Try
        End Sub
        Private Sub SaveToJson(filePath As String)
            Try
                Dim dataTable = CType(DGV1.DataSource, DataTable)
                If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
                    MessageBox.Show("No data to save!")
                    Return
                End If

                Dim jsonData As List(Of Dictionary(Of String, Object)) = New List(Of Dictionary(Of String, Object))()
                For Each row As DataRow In dataTable.Rows
                    Dim rowData As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()
                    For i = 0 To row.ItemArray.Length - 1
                        rowData.Add(dataTable.Columns(i).ColumnName, row.ItemArray(i))
                    Next
                    jsonData.Add(rowData)
                Next

                Dim jsonString = JsonConvert.SerializeObject(jsonData, Newtonsoft.Json.Formatting.Indented)
                File.WriteAllText(filePath, jsonString)

                MessageBox.Show("Data saved successfully to JSON: " & filePath)
            Catch ex As Exception
            End Try
        End Sub
        Private Sub btnCsv_Click(sender As Object, e As EventArgs)
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = "CSV files (.csv)|.csv"
            saveFileDialog.Title = "Save data to CSV"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath = saveFileDialog.FileName
                SaveToCsv(filePath)
            End If
        End Sub
        Private Sub btnXml_Click(sender As Object, e As EventArgs)
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = "XML files (.xml)|.xml"
            saveFileDialog.Title = "Save data to XML"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath = saveFileDialog.FileName
                SaveToXml(filePath)
            End If
        End Sub
        Private Sub btnJson_Click(sender As Object, e As EventArgs)
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = "JSON files (.json)|.json"
            saveFileDialog.Title = "Save data to XML"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath = saveFileDialog.FileName
                SaveToJson(filePath)
            End If
        End Sub
        Private Sub btnAdd_Click(sender As Object, e As EventArgs)
            Dim dataTable = CType(DGV1.DataSource, DataTable)

            ' Create a new row with empty values
            Dim newRow As DataRow = dataTable.NewRow()
            For i = 0 To dataTable.Columns.Count - 1
                newRow(i) = DBNull.Value ' Set default values as appropriate
            Next

            ' Add the new row to the DataTable
            dataTable.Rows.Add(newRow)

            ' Update the DataGridView to reflect the changes
            DGV1.Refresh()
        End Sub

        ' DESIGNER
        Private Sub InitializeComponent()
            Me.btnCsv = New System.Windows.Forms.Button()
            Me.btnXml = New System.Windows.Forms.Button()
            Me.btnJson = New System.Windows.Forms.Button()
            Me.btnOpen = New System.Windows.Forms.Button()
            Me.btnSave = New System.Windows.Forms.Button()
            Me.btnMost = New System.Windows.Forms.Button()
            Me.btnAdd = New System.Windows.Forms.Button()
            Me.DGV1 = New System.Windows.Forms.DataGridView()
            CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'btnCsv
            '
            Me.btnCsv.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnCsv.Location = New System.Drawing.Point(1317, 113)
            Me.btnCsv.Name = "btnCsv"
            Me.btnCsv.Size = New System.Drawing.Size(143, 41)
            Me.btnCsv.TabIndex = 0
            Me.btnCsv.Text = "SAVE CSV"
            Me.btnCsv.UseVisualStyleBackColor = True
            '
            'btnXml
            '
            Me.btnXml.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnXml.Location = New System.Drawing.Point(1317, 173)
            Me.btnXml.Name = "btnXml"
            Me.btnXml.Size = New System.Drawing.Size(143, 41)
            Me.btnXml.TabIndex = 1
            Me.btnXml.Text = "SAVE XML"
            Me.btnXml.UseVisualStyleBackColor = True
            '
            'btnJson
            '
            Me.btnJson.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnJson.Location = New System.Drawing.Point(1317, 235)
            Me.btnJson.Name = "btnJson"
            Me.btnJson.Size = New System.Drawing.Size(143, 41)
            Me.btnJson.TabIndex = 2
            Me.btnJson.Text = "SAVE JSON"
            Me.btnJson.UseVisualStyleBackColor = True
            '
            'btnOpen
            '
            Me.btnOpen.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnOpen.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
            Me.btnOpen.Location = New System.Drawing.Point(1317, 26)
            Me.btnOpen.Name = "btnOpen"
            Me.btnOpen.Size = New System.Drawing.Size(168, 34)
            Me.btnOpen.TabIndex = 3
            Me.btnOpen.Text = "OPEN FILE"
            Me.btnOpen.UseVisualStyleBackColor = True
            '
            'btnSave
            '
            Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnSave.Location = New System.Drawing.Point(62, 512)
            Me.btnSave.Name = "btnSave"
            Me.btnSave.Size = New System.Drawing.Size(500, 51)
            Me.btnSave.TabIndex = 4
            Me.btnSave.Text = "UPDATE IN THE DATABASE"
            Me.btnSave.UseVisualStyleBackColor = True
            '
            'btnMost
            '
            Me.btnMost.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnMost.Location = New System.Drawing.Point(699, 512)
            Me.btnMost.Name = "btnMost"
            Me.btnMost.Size = New System.Drawing.Size(500, 51)
            Me.btnMost.TabIndex = 5
            Me.btnMost.Text = "UPLOAD FROM THE DATABASE"
            Me.btnMost.UseVisualStyleBackColor = True
            '
            'btnAdd
            '
            Me.btnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.btnAdd.Location = New System.Drawing.Point(1317, 336)
            Me.btnAdd.Name = "btnAdd"
            Me.btnAdd.Size = New System.Drawing.Size(168, 34)
            Me.btnAdd.TabIndex = 6
            Me.btnAdd.Text = "ADD ROWS"
            Me.btnAdd.UseVisualStyleBackColor = True
            '
            'DGV1
            '
            Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.DGV1.Location = New System.Drawing.Point(12, 26)
            Me.DGV1.Name = "DGV1"
            Me.DGV1.RowHeadersWidth = 51
            Me.DGV1.RowTemplate.Height = 24
            Me.DGV1.Size = New System.Drawing.Size(1279, 446)
            Me.DGV1.TabIndex = 7
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.MenuHighlight
            Me.ClientSize = New System.Drawing.Size(1655, 624)
            Me.Controls.Add(Me.DGV1)
            Me.Controls.Add(Me.btnAdd)
            Me.Controls.Add(Me.btnMost)
            Me.Controls.Add(Me.btnSave)
            Me.Controls.Add(Me.btnOpen)
            Me.Controls.Add(Me.btnJson)
            Me.Controls.Add(Me.btnXml)
            Me.Controls.Add(Me.btnCsv)
            Me.Name = "Form1"
            Me.Text = "Form1"
            CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub


        Private btnCsv As Windows.Forms.Button
        Private btnXml As Windows.Forms.Button
        Private btnJson As Windows.Forms.Button
        Private btnOpen As Windows.Forms.Button
        Private btnSave As Windows.Forms.Button
        Private btnMost As Windows.Forms.Button
        Private btnAdd As Windows.Forms.Button
        Private DGV1 As Windows.Forms.DataGridView

    End Class
End Namespace
