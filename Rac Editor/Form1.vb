Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports NHunspell
Public Class Form1
    Private fileNameEditor1 As String = String.Empty
    Private fileNameEditor2 As String = String.Empty
    Private FielDirEdito1 As String = Application.StartupPath
    Private FielDirEdito2 As String = Application.StartupPath
    Private dicPath As String
    Private affPath As String
    Private hunspell As Hunspell
    Public Class UndoRedoClass(Of T)
        Private UndoStack As Stack(Of T)
        Private RedoStack As Stack(Of T)

        Public CurrentItem As T
        Public Event UndoHappened As EventHandler(Of UndoRedoEventArgs)
        Public Event RedoHappened As EventHandler(Of UndoRedoEventArgs)

        Public Sub New()
            UndoStack = New Stack(Of T)
            RedoStack = New Stack(Of T)
        End Sub

        Public Sub Clear()
            UndoStack.Clear()
            RedoStack.Clear()
            CurrentItem = Nothing
        End Sub

        Public Sub AddItem(ByVal item As T)
            If CurrentItem IsNot Nothing Then UndoStack.Push(CurrentItem)
            CurrentItem = item
            RedoStack.Clear()
        End Sub

        Public Sub Undo()
            RedoStack.Push(CurrentItem)
            CurrentItem = UndoStack.Pop()
            RaiseEvent UndoHappened(Me, New UndoRedoEventArgs(CurrentItem))
        End Sub

        Public Sub Redo()
            UndoStack.Push(CurrentItem)
            CurrentItem = RedoStack.Pop
            RaiseEvent RedoHappened(Me, New UndoRedoEventArgs(CurrentItem))
        End Sub

        Public Function CanUndo() As Boolean
            Return UndoStack.Count > 0
        End Function

        Public Function CanRedo() As Boolean
            Return RedoStack.Count > 0
        End Function

        Public Function UndoItems() As List(Of T)
            Return UndoStack.ToList
        End Function

        Public Function RedoItems() As List(Of T)
            Return RedoStack.ToList
        End Function
    End Class

    Public Class UndoRedoEventArgs
        Inherits EventArgs

        Private _CurrentItem As Object
        Public ReadOnly Property CurrentItem() As Object
            Get
                Return _CurrentItem
            End Get
        End Property

        Public Sub New(ByVal currentItem As Object)
            _CurrentItem = currentItem
        End Sub
    End Class
    Public Class richtextboxex
        Inherits RichTextBox

    End Class
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        dicPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "en_US.dic")
        affPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "en_US.aff")
        ' Me.Text = Application.StartupPath
        ' Validate dictionary files exist
        If Not File.Exists(dicPath) OrElse Not File.Exists(affPath) Then
            MessageBox.Show("Dictionary files not found in 'dict' folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Editor1ToolStripMenuItem.Text = "Untitled.rtf"
        Editor2ToolStripMenuItem.Text = "Untitled.rtf"
        fileNameEditor1 = Editor1ToolStripMenuItem.Text
        fileNameEditor2 = Editor2ToolStripMenuItem.Text

        Dim pagesDir As String = System.IO.Path.Combine(Application.StartupPath, "Documents")
        'System.IO.Directory.CreateDirectory(pagesDir)
        FielDirEdito1 = System.IO.Path.Combine(pagesDir, fileNameEditor1)
        FielDirEdito2 = System.IO.Path.Combine(pagesDir, fileNameEditor2)

        lblFileDriEditor1.Text = FielDirEdito1
        lblFileDriEditor2.Text = FielDirEdito2
        ' Initialize Hunspell
        hunspell = New Hunspell(affPath, dicPath)


    End Sub

    Private Sub SpellCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpellCheckToolStripMenuItem.Click
        ' Use SpellCheckPreserveCaret to avoid shifting text/caret during replacements
        If hunspell Is Nothing Then
            MessageBox.Show("Dictionary not loaded.")
            Return
        End If

        SpellCheckPreserveCaret(RichTextBox1, hunspell)
        MessageBox.Show("Spell check complete.")
    End Sub

    Private Sub SpellCheckToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SpellCheckToolStripMenuItem1.Click





        If hunspell Is Nothing Then
            MessageBox.Show("Dictionary not loaded.")
            Return
        End If

        SpellCheckPreserveCaret(RichTextBox2, hunspell)
        MessageBox.Show("Spell check complete.")
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim input As String = InputBox("Enter Title (without extension)", "New RTF File", "Untitled")
        If String.IsNullOrWhiteSpace(input) Then
            input = "Untitled"
        End If

        ' Sanitize filename: replace invalid chars
        For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
            input = input.Replace(c, "_"c)
        Next

        ' Ensure .html extension
        If Not input.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase) Then
            input &= ".rtf"
        End If

        fileNameEditor1 = input

        ' Ensure directory exists and build full path
        Dim pagesDir As String = System.IO.Path.Combine(Application.StartupPath, "Documents")
        'System.IO.Directory.CreateDirectory(pagesDir)
        FielDirEdito1 = System.IO.Path.Combine(pagesDir, fileNameEditor1)
        Editor1ToolStripMenuItem.Text = System.IO.Path.GetFileName(FielDirEdito1)
        'Editor1ToolStripMenuItem.Text = System.IO.Path.GetFileName(pagesDir, FileName1)
        Me.Text = FielDirEdito1






        lblFileDriEditor1.Text = FielDirEdito1
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If Editor1ToolStripMenuItem.Text = "Untitled.rtf" Then
            MessageBox.Show("No Text In Editor 1", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim filePath As String = lblFileDriEditor1.Text

            If String.IsNullOrWhiteSpace(filePath) Then
                MessageBox.Show("File Path Is Empty", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' SaveFile correctly accepts a RichTextBoxStreamType for RTF content
            RichTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText)
            MessageBox.Show("File saved successfully In Editor 1 !", "success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error saving file: " & ex.Message)
        End Try

    End Sub

    Private Sub OpenToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem1.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Open Text or RTF File"
            ofd.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            ofd.FilterIndex = 1
            ofd.RestoreDirectory = True

            ' Show dialog and check if user selected a file
            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    ' Determine file type and load accordingly
                    Dim ext As String = Path.GetExtension(ofd.FileName).ToLower()
                    Dim Dir1 As String = Path.GetFullPath(ofd.FileName).ToLower()
                    Dim rtfName1 As String = Path.GetFileName(ofd.FileName).ToLower()
                    lblFileDriEditor1.Text = Dir1
                    Editor1ToolStripMenuItem.Text = rtfName1

                    If ext = ".rtf" Then
                        RichTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.RichText)
                    Else
                        RichTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.PlainText)
                    End If

                Catch ex As IOException
                    MessageBox.Show("Error reading file: " & ex.Message, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As UnauthorizedAccessException
                    MessageBox.Show("Access denied: " & ex.Message, "Permission Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As Exception
                    MessageBox.Show("Unexpected error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim saveFileDialog As New SaveFileDialog()

        ' Set filter for RTF files
        saveFileDialog.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        saveFileDialog.Title = "Save RTF File"

        ' Show the dialog and check if the user clicked OK
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Save the content of RichTextBox1 in RTF format
                Dim ext As String = Path.GetExtension(saveFileDialog.FileName).ToLower()
                Dim Dir1 As String = Path.GetFullPath(saveFileDialog.FileName).ToLower()
                Dim rtfName1 As String = Path.GetFileName(saveFileDialog.FileName).ToLower()
                lblFileDriEditor1.Text = Dir1
                Editor1ToolStripMenuItem.Text = rtfName1
                If ext = ".rtf" Then
                    RichTextBox1.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText)
                    MessageBox.Show(" rtf File saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    RichTextBox1.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText)
                    MessageBox.Show(" txt File saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                ' Handle any errors during the save process
                MessageBox.Show("An error occurred while saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
        RichTextBox1.Select()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
        RichTextBox1.Select()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()

    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub PageFontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageFontToolStripMenuItem.Click
        Dim opFontDLG As New FontDialog


        If opFontDLG.ShowDialog() = DialogResult.OK Then
            RichTextBox1.Font = opFontDLG.Font
        End If
    End Sub

    Private Sub PageTextColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageTextColorToolStripMenuItem.Click
        Dim PageTextColor As New ColorDialog


        If PageTextColor.ShowDialog() = DialogResult.OK Then
            RichTextBox1.ForeColor = PageTextColor.Color
        End If
    End Sub

    Private Sub PageBackColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageBackColorToolStripMenuItem.Click
        Dim PageBackColor As New ColorDialog


        If PageBackColor.ShowDialog() = DialogResult.OK Then
            RichTextBox1.BackColor = PageBackColor.Color
        End If
    End Sub

    Private Sub SearchTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchTextToolStripMenuItem.Click
        Dim searchTerm As String = InputBox("Enter Text To Search for")

        Dim matchesFound As New List(Of (LineNumber As Integer, CharIndex As Integer))

        ' Clear previous highlights
        RichTextBox1.SelectAll()
        'RichTextBox1.SelectionBackColor = Color.White

        ' Loop through each line
        For lineIndex As Integer = 0 To RichTextBox1.Lines.Length - 1
            Dim lineText As String = RichTextBox1.Lines(lineIndex)

            ' Use Regex for case-insensitive search
            Dim matches = Regex.Matches(lineText, Regex.Escape(searchTerm), RegexOptions.IgnoreCase)

            For Each m As Match In matches
                ' Calculate absolute character index in RichTextBox
                Dim charIndex As Integer = RichTextBox1.GetFirstCharIndexFromLine(lineIndex) + m.Index
                matchesFound.Add((lineIndex + 1, charIndex))

                ' Highlight match
                RichTextBox1.Select(charIndex, m.Length)
                RichTextBox1.SelectionBackColor = Color.LightBlue
            Next
        Next

        ' Show results
        If matchesFound.Count > 0 Then
            Dim resultText As String = "Matches found:" & Environment.NewLine
            For Each match In matchesFound
                resultText &= $"Line {match.LineNumber}, CharIndex {match.CharIndex}" & Environment.NewLine
            Next
            MessageBox.Show(resultText, "Search Results")
        Else
            MessageBox.Show("No matches found.", "Search Results")
        End If
    End Sub



    Private Sub ClearAllTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllTextToolStripMenuItem.Click

        Dim result As DialogResult = MessageBox.Show("Do you want to delete the text?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            RichTextBox1.Clear()
            fileNameEditor1 = String.Empty
            FielDirEdito1 = String.Empty

            Editor1ToolStripMenuItem.Text = "Untitled.rtf"
            fileNameEditor1 = Editor1ToolStripMenuItem.Text

            Dim pagesDir As String = System.IO.Path.Combine(Application.StartupPath, "Documents")
            FielDirEdito1 = System.IO.Path.Combine(pagesDir, fileNameEditor1)
            lblFileDriEditor1.Text = FielDirEdito1
        Else
            ' User cancelled — restore focus and do nothing
            RichTextBox1.Focus()
        End If



    End Sub



    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Dim input As String = InputBox("Enter Title (without extension)", "New RTF File", "Untitled")
        If String.IsNullOrWhiteSpace(input) Then
            input = "Untitled"
        End If

        ' Sanitize filename: replace invalid chars
        For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
            input = input.Replace(c, "_"c)
        Next

        ' Ensure .html extension
        If Not input.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase) Then
            input &= ".rtf"
        End If

        fileNameEditor2 = input
        Editor2ToolStripMenuItem.Text = fileNameEditor2
        ' Ensure directory exists and build full path
        Dim pagesDir As String = System.IO.Path.Combine(Application.StartupPath, "Documents")
        'System.IO.Directory.CreateDirectory(pagesDir)
        FielDirEdito1 = System.IO.Path.Combine(pagesDir, fileNameEditor2)
        lblFileDriEditor2.Text = FielDirEdito1
        'Editor1ToolStripMenuItem.Text = System.IO.Path.GetFileName(pagesDir, FileName1)







        lblFileDriEditor1.Text = FielDirEdito1
    End Sub

    Private Sub UndoToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem2.Click
        RichTextBox1.Undo()
        RichTextBox1.Select()
    End Sub

    Private Sub RedoToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem2.Click
        RichTextBox1.Redo()
        RichTextBox1.Select()
    End Sub

    Private Sub SelectedTextFontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedTextFontToolStripMenuItem.Click
        Dim opFontDLG As New FontDialog


        If opFontDLG.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SelectionFont = opFontDLG.Font
        End If
    End Sub

    Private Sub SelectedTextColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedTextColorToolStripMenuItem.Click
        Dim PageTextColor As New ColorDialog


        If PageTextColor.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SelectionColor = PageTextColor.Color
        End If
    End Sub

    Private Sub SelectedTextBackColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedTextBackColorToolStripMenuItem.Click
        Dim PageTextColor As New ColorDialog


        If PageTextColor.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SelectionBackColor = PageTextColor.Color
        End If
    End Sub

    Private Sub ClearSelectedTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearSelectedTextToolStripMenuItem.Click
        If RichTextBox1 Is Nothing Then
            Return
        End If

        If RichTextBox1.SelectionLength > 0 Then
            Dim result As DialogResult = MessageBox.Show("Do you want to delete the selected text?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                RichTextBox1.SelectedText = String.Empty
            End If
        Else
            MessageBox.Show("No text selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        RichTextBox1.Focus()
    End Sub

    Private Sub InsertImageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertImageToolStripMenuItem.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Insert Image"
            ofd.Filter = "Image Files|*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tiff|All Files|*.*"
            ofd.FilterIndex = 1
            ofd.RestoreDirectory = True

            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Try
                ' Load image into memory to avoid locking the file
                Dim imgBytes As Byte() = File.ReadAllBytes(ofd.FileName)
                Using ms As New IO.MemoryStream(imgBytes)
                    Using img As Image = Image.FromStream(ms)
                        Dim originalData As IDataObject = Nothing
                        Try
                            originalData = System.Windows.Forms.Clipboard.GetDataObject()
                        Catch
                            originalData = Nothing
                        End Try

                        Try
                            Dim dataObj As New System.Windows.Forms.DataObject()
                            dataObj.SetData(System.Windows.Forms.DataFormats.Bitmap, True, New Bitmap(img))
                            System.Windows.Forms.Clipboard.SetDataObject(dataObj, True)
                            RichTextBox1.Paste()
                        Finally
                            ' Restore original clipboard (best effort)
                            If originalData IsNot Nothing Then
                                Try
                                    System.Windows.Forms.Clipboard.SetDataObject(originalData, True)
                                Catch
                                    ' ignore restore failures
                                End Try
                            End If
                        End Try
                    End Using
                End Using

                RichTextBox1.Focus()
            Catch ex As Exception
                MessageBox.Show("Failed to insert image: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub SpellCheckToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles SpellCheckToolStripMenuItem2.Click
        If hunspell Is Nothing Then
            MessageBox.Show("Dictionary not loaded.")
            Return
        End If

        SpellCheckPreserveCaret(RichTextBox1, hunspell)
        MessageBox.Show("Spell check complete.")
    End Sub

    Private Sub OpenToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem2.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Open Text or RTF File"
            ofd.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            ofd.FilterIndex = 1
            ofd.RestoreDirectory = True

            ' Show dialog and check if user selected a file
            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    ' Determine file type and load accordingly
                    Dim ext As String = Path.GetExtension(ofd.FileName).ToLower()
                    Dim Dir1 As String = Path.GetFullPath(ofd.FileName).ToLower()
                    Dim rtfName1 As String = Path.GetFileName(ofd.FileName).ToLower()
                    lblFileDriEditor2.Text = Dir1
                    Editor2ToolStripMenuItem.Text = rtfName1

                    If ext = ".rtf" Then
                        RichTextBox2.LoadFile(ofd.FileName, RichTextBoxStreamType.RichText)
                    Else
                        RichTextBox2.LoadFile(ofd.FileName, RichTextBoxStreamType.PlainText)
                    End If

                Catch ex As IOException
                    MessageBox.Show("Error reading file: " & ex.Message, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As UnauthorizedAccessException
                    MessageBox.Show("Access denied: " & ex.Message, "Permission Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As Exception
                    MessageBox.Show("Unexpected error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        If Editor2ToolStripMenuItem.Text = "Untitled.rtf" Then
            MessageBox.Show("No Text In Editor 2", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim filePath As String = lblFileDriEditor2.Text

            If String.IsNullOrWhiteSpace(filePath) Then
                MessageBox.Show("File Path Is Empty", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' SaveFile correctly accepts a RichTextBoxStreamType for RTF content
            RichTextBox2.SaveFile(filePath, RichTextBoxStreamType.RichText)
            MessageBox.Show("File saved successfully In Editor 1 !", "success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error saving file: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveAsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem1.Click
        Dim saveFileDialog As New SaveFileDialog()

        ' Set filter for RTF files
        saveFileDialog.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        saveFileDialog.Title = "Save RTF File"

        ' Show the dialog and check if the user clicked OK
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Save the content of RichTextBox1 in RTF format
                Dim ext As String = Path.GetExtension(saveFileDialog.FileName).ToLower()
                Dim Dir1 As String = Path.GetFullPath(saveFileDialog.FileName).ToLower()
                Dim rtfName1 As String = Path.GetFileName(saveFileDialog.FileName).ToLower()
                lblFileDriEditor2.Text = Dir1
                Editor2ToolStripMenuItem.Text = rtfName1
                If ext = ".rtf" Then
                    RichTextBox2.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText)
                    MessageBox.Show(" rtf File saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    RichTextBox2.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText)
                    MessageBox.Show(" txt File saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                ' Handle any errors during the save process
                MessageBox.Show("An error occurred while saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub UndoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem1.Click
        RichTextBox2.Undo()
        RichTextBox2.Select()
    End Sub

    Private Sub RedoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem1.Click
        RichTextBox2.Redo()
        RichTextBox2.Select()

    End Sub

    Private Sub CutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem1.Click
        RichTextBox2.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem1.Click
        RichTextBox2.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem1.Click
        RichTextBox2.Paste()
    End Sub

    Private Sub PageFontToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PageFontToolStripMenuItem1.Click
        Dim opFontDLG As New FontDialog


        If opFontDLG.ShowDialog() = DialogResult.OK Then
            RichTextBox2.Font = opFontDLG.Font
        End If
    End Sub

    Private Sub PageTextColorToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PageTextColorToolStripMenuItem1.Click
        Dim PageBackColor As New ColorDialog


        If PageBackColor.ShowDialog() = DialogResult.OK Then
            RichTextBox2.ForeColor = PageBackColor.Color
        End If
    End Sub

    Private Sub PageBackColorToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PageBackColorToolStripMenuItem1.Click
        Dim PageBackColor As New ColorDialog


        If PageBackColor.ShowDialog() = DialogResult.OK Then
            RichTextBox2.BackColor = PageBackColor.Color
        End If
    End Sub

    Private Sub SearchTextToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SearchTextToolStripMenuItem1.Click
        Dim searchTerm As String = InputBox("Enter Text To Search for")

        Dim matchesFound As New List(Of (LineNumber As Integer, CharIndex As Integer))

        ' Clear previous highlights
        RichTextBox2.SelectAll()
        'RichTextBox1.SelectionBackColor = Color.White

        ' Loop through each line
        For lineIndex As Integer = 0 To RichTextBox2.Lines.Length - 1
            Dim lineText As String = RichTextBox2.Lines(lineIndex)

            ' Use Regex for case-insensitive search
            Dim matches = Regex.Matches(lineText, Regex.Escape(searchTerm), RegexOptions.IgnoreCase)

            For Each m As Match In matches
                ' Calculate absolute character index in RichTextBox
                Dim charIndex As Integer = RichTextBox2.GetFirstCharIndexFromLine(lineIndex) + m.Index
                matchesFound.Add((lineIndex + 1, charIndex))

                ' Highlight match
                RichTextBox2.Select(charIndex, m.Length)
                RichTextBox2.SelectionBackColor = Color.LightBlue
            Next
        Next

        ' Show results
        If matchesFound.Count > 0 Then
            Dim resultText As String = "Matches found:" & Environment.NewLine
            For Each match In matchesFound
                resultText &= $"Line {match.LineNumber}, CharIndex {match.CharIndex}" & Environment.NewLine
            Next
            MessageBox.Show(resultText, "Search Results")
        Else
            MessageBox.Show("No matches found.", "Search Results")
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to delete the text?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            RichTextBox2.Clear()
            fileNameEditor2 = String.Empty
            FielDirEdito2 = String.Empty

            Editor2ToolStripMenuItem.Text = "Untitled.rtf"
            fileNameEditor2 = Editor2ToolStripMenuItem.Text

            Dim pagesDir As String = System.IO.Path.Combine(Application.StartupPath, "Documents")
            FielDirEdito2 = System.IO.Path.Combine(pagesDir, fileNameEditor2)
            lblFileDriEditor2.Text = FielDirEdito2
        Else
            ' User cancelled — restore focus and do nothing
            RichTextBox2.Focus()
        End If

    End Sub

    Private Sub LeftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub CenterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CenterToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub RightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Right
    End Sub
    Private Sub SpellCheckPreserveCaret(rtb As RichTextBox, hun As Hunspell)
        If hun Is Nothing Then Return

        ' save caret/selection
        Dim origStart = rtb.SelectionStart
        Dim origLen = rtb.SelectionLength
        Dim adjustedStart As Integer = origStart

        ' disable redraw to avoid flicker (optional)
        Dim WM_SETREDRAW As Integer = &HB
        SendMessage(rtb.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)

        Try
            Dim pattern As String = "\b[\w']+\b" ' word boundaries (adjust for your needs)
            Dim matches = Regex.Matches(rtb.Text, pattern)
            For i As Integer = matches.Count - 1 To 0 Step -1
                Dim m As Match = matches(i)
                Dim w As String = m.Value
                If Not hun.Spell(w) Then
                    Dim suggestions = hun.Suggest(w)
                    If suggestions.Count > 0 Then
                        Dim suggestion As String = suggestions(0)
                        If suggestion <> w Then
                            ' select the word and replace in-place
                            rtb.Select(m.Index, m.Length)
                            rtb.SelectedText = suggestion

                            ' adjust saved caret start if replacement happened before it
                            Dim lenDiff As Integer = suggestion.Length - m.Length
                            If m.Index < adjustedStart Then
                                adjustedStart += lenDiff
                            End If
                        End If
                    End If
                End If
            Next
        Finally
            ' restore selection/caret (try to keep the same logical position)
            If adjustedStart <= rtb.TextLength Then
                rtb.Select(adjustedStart, Math.Min(origLen, Math.Max(0, rtb.TextLength - adjustedStart)))
            Else
                rtb.Select(rtb.TextLength, 0)
            End If

            ' re-enable redraw and refresh
            SendMessage(rtb.Handle, WM_SETREDRAW, CType(1, IntPtr), IntPtr.Zero)
            rtb.Invalidate()
        End Try
    End Sub

    ' P/Invoke to turn redraw on/off:
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=False)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Sub UndoToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem3.Click
        RichTextBox2.Undo()
        RichTextBox2.Select()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        RichTextBox2.Redo()
        RichTextBox2.Select()
    End Sub

    Private Sub CutToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem3.Click
        RichTextBox2.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem3.Click
        RichTextBox2.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem3.Click
        RichTextBox2.Paste()
    End Sub

    Private Sub SelectedTextColoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedTextColoToolStripMenuItem.Click
        Dim PageBackColor As New ColorDialog


        If PageBackColor.ShowDialog() = DialogResult.OK Then
            RichTextBox2.SelectionColor = PageBackColor.Color
        End If
    End Sub

    Private Sub SelectedTextBackColorToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SelectedTextBackColorToolStripMenuItem1.Click
        Dim PageTextColor As New ColorDialog


        If PageTextColor.ShowDialog() = DialogResult.OK Then
            RichTextBox2.SelectionBackColor = PageTextColor.Color
        End If
    End Sub

    Private Sub SelectedTextFontToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SelectedTextFontToolStripMenuItem1.Click
        Dim opFontDLG As New FontDialog


        If opFontDLG.ShowDialog() = DialogResult.OK Then
            RichTextBox2.SelectionFont = opFontDLG.Font
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        If RichTextBox2 Is Nothing Then
            Return
        End If

        If RichTextBox2.SelectionLength > 0 Then
            Dim result As DialogResult = MessageBox.Show("Do you want to delete the selected text?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                RichTextBox2.SelectedText = String.Empty
            End If
        Else
            MessageBox.Show("No text selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        RichTextBox2.Focus()
    End Sub

    Private Sub IncertImageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncertImageToolStripMenuItem.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Insert Image"
            ofd.Filter = "Image Files|*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tiff;*.ico|All Files|*.*"
            ofd.FilterIndex = 1
            ofd.RestoreDirectory = True

            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Try
                ' Load image into memory to avoid locking the file
                Dim imgBytes As Byte() = File.ReadAllBytes(ofd.FileName)
                Using ms As New IO.MemoryStream(imgBytes)
                    Using img As Image = Image.FromStream(ms)
                        Dim originalData As IDataObject = Nothing
                        Try
                            originalData = System.Windows.Forms.Clipboard.GetDataObject()
                        Catch
                            originalData = Nothing
                        End Try

                        Try
                            Dim dataObj As New System.Windows.Forms.DataObject()
                            dataObj.SetData(System.Windows.Forms.DataFormats.Bitmap, True, New Bitmap(img))
                            System.Windows.Forms.Clipboard.SetDataObject(dataObj, True)
                            RichTextBox2.Paste()
                        Finally
                            ' Restore original clipboard (best effort)
                            If originalData IsNot Nothing Then
                                Try
                                    System.Windows.Forms.Clipboard.SetDataObject(originalData, True)
                                Catch
                                    ' ignore restore failures
                                End Try
                            End If
                        End Try
                    End Using
                End Using

                RichTextBox2.Focus()
            Catch ex As Exception
                MessageBox.Show("Failed to insert image: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub LeftToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem1.Click
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub CenterToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CenterToolStripMenuItem1.Click
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub RightToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem1.Click
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Right
    End Sub

    Private Sub SpellCheckToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles SpellCheckToolStripMenuItem3.Click
        If hunspell Is Nothing Then
            MessageBox.Show("Dictionary not loaded.")
            Return
        End If

        SpellCheckPreserveCaret(RichTextBox2, hunspell)
        MessageBox.Show("Spell check complete.")
    End Sub



    Private Sub RichTextBox1_Click(sender As Object, e As EventArgs) Handles RichTextBox1.Click
        Me.Text = "Editor 1"
    End Sub



    Private Sub RichTextBox2_Click(sender As Object, e As EventArgs) Handles RichTextBox2.Click
        Me.Text = "Editor 2"
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.Show()
    End Sub
End Class
