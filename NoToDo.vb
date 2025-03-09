Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Collections.Generic

Public Class MainForm
    Inherits Form

    Private taskGrid As New DataGridView()
    Private inputBox As New TextBox()
    Private addButton As New Button()
    Private iniPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NoToDo.ini")

    ' Variables for drag-and-drop functionality
    Private dragRowIndex As Integer = -1
    Private isDragging As Boolean = False

    Public Sub New()
        ' Form setup
        Me.Text = "NoToDo - Simple Note ToDo by amymor OgomnamO - v1.0"
        Me.Size = New Size(650, 600)
        Me.BackColor = Color.FromArgb(40, 40, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 12)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        AddHandler Me.Resize, AddressOf MainForm_Resize

        ' DataGridView setup
        taskGrid.BackgroundColor = Color.FromArgb(50, 50, 50)
        taskGrid.ForeColor = Color.FromArgb(225, 225, 225)
        taskGrid.BorderStyle = BorderStyle.FixedSingle
        taskGrid.ColumnHeadersVisible = False
        taskGrid.RowHeadersVisible = False
        taskGrid.AllowUserToAddRows = False
        taskGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        taskGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        ' taskGrid.MultiSelect = False
        taskGrid.MultiSelect = True
        taskGrid.RowTemplate.Height = 40
        taskGrid.EnableHeadersVisualStyles = False
        taskGrid.DefaultCellStyle.BackColor = Color.FromArgb(50, 50, 50)
        taskGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(65, 65, 75)
        taskGrid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(255, 225, 200)
        taskGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
		' taskGrid.CellBorderStyle = DataGridViewCellBorderStyle.None
		taskGrid.AllowUserToResizeColumns = False
		taskGrid.AllowUserToResizeRows = False

        ' Define columns
        Dim dragColumn As New DataGridViewTextBoxColumn()
        dragColumn.HeaderText = ""
        dragColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dragColumn.Width = 30
        dragColumn.ReadOnly = True
        dragColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dragColumn.DefaultCellStyle.NullValue = "≡"

        Dim taskColumn As New DataGridViewTextBoxColumn()
        ' taskColumn.HeaderText = "Task"
        taskColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        ' Copy button column
        Dim copyColumn As New DataGridViewButtonColumn()
        copyColumn.HeaderText = ""
        copyColumn.Text = "📋"
        copyColumn.UseColumnTextForButtonValue = True
        ' copyColumn.DefaultCellStyle.ForeColor = Color.FromArgb(225, 225, 225)
        copyColumn.DefaultCellStyle.ForeColor = Color.FromArgb(185,185,185)
        ' copyColumn.DefaultCellStyle.BackColor = Color.FromArgb(33, 69, 84) ' Dark pastel blue
        copyColumn.DefaultCellStyle.BackColor = Color.FromArgb(30, 50, 70)
        ' copyColumn.DefaultCellStyle.SelectionForeColor = Color.FromArgb(225, 225, 225)
        copyColumn.DefaultCellStyle.SelectionForeColor = Color.FromArgb(185,185,185)
        ' copyColumn.DefaultCellStyle.SelectionBackColor = Color.FromArgb(43, 79, 94)
        copyColumn.DefaultCellStyle.SelectionBackColor = Color.FromArgb(35, 55, 80)
        copyColumn.FlatStyle = FlatStyle.Flat
        copyColumn.Width = 80

        ' Delete button column
        Dim deleteColumn As New DataGridViewButtonColumn()
        deleteColumn.HeaderText = ""
        deleteColumn.Text = "❌"
        deleteColumn.UseColumnTextForButtonValue = True
        ' deleteColumn.DefaultCellStyle.ForeColor = Color.FromArgb(255, 200, 200)
        deleteColumn.DefaultCellStyle.ForeColor = Color.FromArgb(185,185,185)
        ' deleteColumn.DefaultCellStyle.BackColor = Color.FromArgb(100, 0, 0) ' Dark pastel red
        deleteColumn.DefaultCellStyle.BackColor = Color.FromArgb(60, 25, 25)
        ' deleteColumn.DefaultCellStyle.SelectionForeColor = Color.FromArgb(255, 200, 200)
        deleteColumn.DefaultCellStyle.SelectionForeColor = Color.FromArgb(185,185,185)
        ' deleteColumn.DefaultCellStyle.SelectionBackColor = Color.FromArgb(139, 26, 26)
        deleteColumn.DefaultCellStyle.SelectionBackColor = Color.FromArgb(75, 30, 30)
        deleteColumn.FlatStyle = FlatStyle.Flat
        deleteColumn.Width = 80

        taskGrid.Columns.AddRange(dragColumn, taskColumn, copyColumn, deleteColumn)

        ' Input box setup
        ' inputBox.Location = New Point(20, 440)
        inputBox.Size = New Size(400, 35)
        inputBox.BackColor = Color.FromArgb(60, 60, 60)
        inputBox.ForeColor = Color.FromArgb(225, 225, 225)
        ' inputBox.Font = New Font("Segoe UI", 14) ' Set font size
        ' inputBox.BorderStyle = BorderStyle.None
        inputBox.BorderStyle = BorderStyle.FixedSingle
		
        inputBox.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        inputBox.Margin = New Padding(15)

        ' Add button setup
        addButton.Text = " Add 📝"
        inputBox.Font = New Font("Segoe UI", 15) ' Set font size
        ' addButton.Location = New Point(430, 440)
        addButton.Size = New Size(130, 35)
        ' addButton.BackColor = Color.FromArgb(0, 122, 204)
        addButton.BackColor = Color.FromArgb(0,90,25)
        addButton.ForeColor = Color.FromArgb(225, 225, 225)
        addButton.FlatStyle = FlatStyle.Flat
        addButton.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        addButton.FlatAppearance.BorderSize = 1
		addButton.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        AddHandler addButton.Click, AddressOf AddTask

        ' Event handlers
        AddHandler taskGrid.CellContentClick, AddressOf TaskGrid_CellContentClick
        AddHandler taskGrid.MouseDown, AddressOf TaskGrid_MouseDown
        AddHandler taskGrid.MouseMove, AddressOf TaskGrid_MouseMove
        AddHandler taskGrid.MouseUp, AddressOf TaskGrid_MouseUp
        AddHandler Me.Load, AddressOf LoadTasks
        AddHandler Me.FormClosing, AddressOf SaveTasks

        ' Add controls to form
        Me.Controls.Add(taskGrid)
        Me.Controls.Add(inputBox)
        Me.Controls.Add(addButton)

        ' Initial layout setup
        UpdateLayout()

		' focus to the TextBox at startup
		inputBox.text = "Type something here and press Add"
		inputBox.Select()
    End Sub

    Private Sub UpdateLayout()
        ' Set positions and sizes dynamically
        Dim padding As Integer = 15

        ' DataGridView
        taskGrid.Location = New Point(padding, padding)
        taskGrid.Size = New Size(Me.ClientSize.Width - (padding * 2), Me.ClientSize.Height - inputBox.Height - addButton.Height - (padding * 2))

        ' InputBox
        inputBox.Location = New Point(padding, Me.ClientSize.Height - inputBox.Height - padding)
        inputBox.Size = New Size(Me.ClientSize.Width - addButton.Width - (padding * 3), inputBox.Height)

        ' AddButton
        addButton.Location = New Point(Me.ClientSize.Width - addButton.Width - padding, Me.ClientSize.Height - addButton.Height - padding)
    End Sub

    Private Sub MainForm_Resize(sender As Object, e As EventArgs)
        UpdateLayout()
    End Sub
' =======================================================

    Private Sub AddTask(sender As Object, e As EventArgs)
        ' Add an empty task if the input box is empty or contains only whitespace
        taskGrid.Rows.Add("≡", inputBox.Text.Trim(), "Copy", "❌")
        inputBox.Clear()
        SaveTasks()
    End Sub

Private Sub TaskGrid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
    If e.RowIndex >= 0 Then
        ' Copy button clicked
        If e.ColumnIndex = 2 Then
            Dim taskValue As Object = taskGrid.Rows(e.RowIndex).Cells(1).Value
            Dim taskText As String = If(taskValue IsNot Nothing, taskValue.ToString(), "")

            ' Disable copy for truly empty rows (not just whitespace)
            If String.IsNullOrEmpty(taskText) Then
                ' MessageBox.Show("Cannot copy an empty task.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Copy the task text to the clipboard
            Clipboard.SetText(taskText)
        End If

        ' Delete button clicked
        If e.ColumnIndex = 3 Then
            taskGrid.Rows.RemoveAt(e.RowIndex)
            SaveTasks()
        End If
    End If
End Sub


    Private Sub LoadTasks(sender As Object, e As EventArgs)
        If File.Exists(iniPath) Then
            Dim lines() As String = File.ReadAllLines(iniPath)
            For Each line As String In lines
                taskGrid.Rows.Add("≡", line, "Copy", "❌")
            Next
        End If
    End Sub

    Private Sub SaveTasks()
        Dim tasks As New List(Of String)
        For Each row As DataGridViewRow In taskGrid.Rows
            Dim taskValue As Object = row.Cells(1).Value
            Dim taskText As String = If(taskValue IsNot Nothing, taskValue.ToString(), "")
            tasks.Add(taskText)
        Next
        File.WriteAllLines(iniPath, tasks)
    End Sub

    ' Overload for FormClosing event
    Private Sub SaveTasks(sender As Object, e As FormClosingEventArgs)
        SaveTasks()
    End Sub

    ' Drag-and-drop functionality
    Private Sub TaskGrid_MouseDown(sender As Object, e As MouseEventArgs)
        Dim hitTestInfo = taskGrid.HitTest(e.X, e.Y)
        If hitTestInfo.Type = DataGridViewHitTestType.Cell AndAlso hitTestInfo.ColumnIndex = 0 Then
            dragRowIndex = hitTestInfo.RowIndex
            isDragging = True
        End If
    End Sub

    Private Sub TaskGrid_MouseMove(sender As Object, e As MouseEventArgs)
        If isDragging AndAlso dragRowIndex >= 0 Then
            Dim hitTestInfo = taskGrid.HitTest(e.X, e.Y)
            If hitTestInfo.Type = DataGridViewHitTestType.Cell Then
                Dim targetRowIndex = hitTestInfo.RowIndex
                If targetRowIndex <> dragRowIndex Then
                    Dim rowToMove = taskGrid.Rows(dragRowIndex)
                    taskGrid.Rows.RemoveAt(dragRowIndex)
                    taskGrid.Rows.Insert(targetRowIndex, rowToMove)
                    dragRowIndex = targetRowIndex
                End If
            End If
        End If
    End Sub

    Private Sub TaskGrid_MouseUp(sender As Object, e As MouseEventArgs)
        If isDragging Then
            isDragging = False
            dragRowIndex = -1
            SaveTasks()
        End If
    End Sub

    <STAThread>
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Class