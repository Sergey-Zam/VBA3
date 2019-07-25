<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblCountOfAssembly = New System.Windows.Forms.Label()
        Me.lblCountOfExcel = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbAssemblyDirectory = New System.Windows.Forms.TextBox()
        Me.tbExcelDirectory = New System.Windows.Forms.TextBox()
        Me.btnCompare = New System.Windows.Forms.Button()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.btnGetDataFromAssembly = New System.Windows.Forms.Button()
        Me.dgvAspects = New System.Windows.Forms.DataGridView()
        Me.btnGetDataFromExcel = New System.Windows.Forms.Button()
        Me.lblLoading = New System.Windows.Forms.Label()
        Me.btnExportToExcel = New System.Windows.Forms.Button()
        CType(Me.dgvAspects, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCountOfAssembly
        '
        Me.lblCountOfAssembly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCountOfAssembly.AutoSize = True
        Me.lblCountOfAssembly.Location = New System.Drawing.Point(859, 73)
        Me.lblCountOfAssembly.Name = "lblCountOfAssembly"
        Me.lblCountOfAssembly.Size = New System.Drawing.Size(13, 13)
        Me.lblCountOfAssembly.TabIndex = 23
        Me.lblCountOfAssembly.Text = "0"
        '
        'lblCountOfExcel
        '
        Me.lblCountOfExcel.AutoSize = True
        Me.lblCountOfExcel.Location = New System.Drawing.Point(179, 73)
        Me.lblCountOfExcel.Name = "lblCountOfExcel"
        Me.lblCountOfExcel.Size = New System.Drawing.Size(13, 13)
        Me.lblCountOfExcel.TabIndex = 22
        Me.lblCountOfExcel.Text = "0"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(682, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(171, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Данные, полученные из сборки:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(161, 13)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Данные, полученные из Excel:"
        '
        'tbAssemblyDirectory
        '
        Me.tbAssemblyDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbAssemblyDirectory.Location = New System.Drawing.Point(234, 45)
        Me.tbAssemblyDirectory.Name = "tbAssemblyDirectory"
        Me.tbAssemblyDirectory.ReadOnly = True
        Me.tbAssemblyDirectory.Size = New System.Drawing.Size(638, 20)
        Me.tbAssemblyDirectory.TabIndex = 19
        '
        'tbExcelDirectory
        '
        Me.tbExcelDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExcelDirectory.Location = New System.Drawing.Point(234, 15)
        Me.tbExcelDirectory.Name = "tbExcelDirectory"
        Me.tbExcelDirectory.ReadOnly = True
        Me.tbExcelDirectory.Size = New System.Drawing.Size(638, 20)
        Me.tbExcelDirectory.TabIndex = 18
        '
        'btnCompare
        '
        Me.btnCompare.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCompare.ForeColor = System.Drawing.Color.Black
        Me.btnCompare.Location = New System.Drawing.Point(12, 525)
        Me.btnCompare.Name = "btnCompare"
        Me.btnCompare.Size = New System.Drawing.Size(552, 23)
        Me.btnCompare.TabIndex = 17
        Me.btnCompare.Text = "Сравнить"
        Me.btnCompare.UseVisualStyleBackColor = True
        '
        'btnClearAll
        '
        Me.btnClearAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearAll.ForeColor = System.Drawing.Color.Black
        Me.btnClearAll.Location = New System.Drawing.Point(724, 525)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(148, 23)
        Me.btnClearAll.TabIndex = 16
        Me.btnClearAll.Text = "Очистить таблицу"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'btnGetDataFromAssembly
        '
        Me.btnGetDataFromAssembly.ForeColor = System.Drawing.Color.Black
        Me.btnGetDataFromAssembly.Location = New System.Drawing.Point(12, 43)
        Me.btnGetDataFromAssembly.Name = "btnGetDataFromAssembly"
        Me.btnGetDataFromAssembly.Size = New System.Drawing.Size(216, 23)
        Me.btnGetDataFromAssembly.TabIndex = 14
        Me.btnGetDataFromAssembly.Text = "Импорт данных из сборки (*.iam)"
        Me.btnGetDataFromAssembly.UseVisualStyleBackColor = True
        '
        'dgvAspects
        '
        Me.dgvAspects.AllowUserToAddRows = False
        Me.dgvAspects.AllowUserToDeleteRows = False
        Me.dgvAspects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAspects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAspects.Location = New System.Drawing.Point(12, 89)
        Me.dgvAspects.Name = "dgvAspects"
        Me.dgvAspects.ReadOnly = True
        Me.dgvAspects.RowHeadersVisible = False
        Me.dgvAspects.Size = New System.Drawing.Size(860, 430)
        Me.dgvAspects.TabIndex = 13
        '
        'btnGetDataFromExcel
        '
        Me.btnGetDataFromExcel.Location = New System.Drawing.Point(12, 13)
        Me.btnGetDataFromExcel.Name = "btnGetDataFromExcel"
        Me.btnGetDataFromExcel.Size = New System.Drawing.Size(216, 23)
        Me.btnGetDataFromExcel.TabIndex = 12
        Me.btnGetDataFromExcel.Text = "Импорт данных из Excel (*.xlsx, *.xls)"
        Me.btnGetDataFromExcel.UseVisualStyleBackColor = True
        '
        'lblLoading
        '
        Me.lblLoading.AutoSize = True
        Me.lblLoading.BackColor = System.Drawing.Color.CornflowerBlue
        Me.lblLoading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLoading.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblLoading.ForeColor = System.Drawing.Color.White
        Me.lblLoading.Location = New System.Drawing.Point(846, 9)
        Me.lblLoading.Name = "lblLoading"
        Me.lblLoading.Padding = New System.Windows.Forms.Padding(12)
        Me.lblLoading.Size = New System.Drawing.Size(26, 42)
        Me.lblLoading.TabIndex = 24
        Me.lblLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLoading.Visible = False
        '
        'btnExportToExcel
        '
        Me.btnExportToExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportToExcel.ForeColor = System.Drawing.Color.Black
        Me.btnExportToExcel.Location = New System.Drawing.Point(570, 525)
        Me.btnExportToExcel.Name = "btnExportToExcel"
        Me.btnExportToExcel.Size = New System.Drawing.Size(148, 23)
        Me.btnExportToExcel.TabIndex = 25
        Me.btnExportToExcel.Text = "Экспорт в Excel"
        Me.btnExportToExcel.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(884, 561)
        Me.Controls.Add(Me.btnExportToExcel)
        Me.Controls.Add(Me.lblLoading)
        Me.Controls.Add(Me.lblCountOfAssembly)
        Me.Controls.Add(Me.lblCountOfExcel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbAssemblyDirectory)
        Me.Controls.Add(Me.tbExcelDirectory)
        Me.Controls.Add(Me.btnCompare)
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.btnGetDataFromAssembly)
        Me.Controls.Add(Me.dgvAspects)
        Me.Controls.Add(Me.btnGetDataFromExcel)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VBA3"
        CType(Me.dgvAspects, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblCountOfAssembly As Label
    Friend WithEvents lblCountOfExcel As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbAssemblyDirectory As TextBox
    Friend WithEvents tbExcelDirectory As TextBox
    Friend WithEvents btnCompare As Button
    Friend WithEvents btnClearAll As Button
    Friend WithEvents btnGetDataFromAssembly As Button
    Friend WithEvents dgvAspects As DataGridView
    Friend WithEvents btnGetDataFromExcel As Button
    Friend WithEvents lblLoading As Label
    Friend WithEvents btnExportToExcel As Button
End Class
