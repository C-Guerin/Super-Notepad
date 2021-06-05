Imports System.Windows.Forms.VisualStyles
Imports System.IO
Imports System.Text.RegularExpressions


Public Class FrmMain
    Public boolSavedDocument As Boolean = False 'An indicator of Whether or not the document has a save location specified
    Public strSaveLocation As String = "" 'The location of the document if it has been saved or opened

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ToolStripStatusLabel2.Text = String.Empty
        ToolStripStatusLabel2.Spring = True
        ToolStripStatusLabel3.Text = String.Empty
        ToolStripStatusLabel3.Spring = True
        'Statuslabel2 and 3 are used to shift Statuslabel1 to the right via the '.Spring' method.
        ToolStripStatusLabel1.Text = ("Lines: 0, Char: 0")
        ToolStripStatusLabel1.Spring = True

        SubStatusBarStatus()
        SubWordWrap()

        btnDebug.Visible = False 'Set this to true if you wish to use the debug function

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'Counts the number of 'Lines' and 'Characters' everytime the text is changed, i.e. the user types.
        SubLineCharCount()
    End Sub

    Public Sub SubStatusBarStatus()
        'Toggles the 'Checked' status of 'StatusBar' in the 'ToolStrip' under 'View'.

        Select Case StatusStrip1.Visible
            Case True
                StatusBarToolStripMenuItem.Checked = True
            Case False
                StatusBarToolStripMenuItem.Checked = False

        End Select
    End Sub

    Public Sub SubWordWrap()
        'Toggles the 'Checked' status of 'WordWrap' in the 'ToolStrip' under 'Format'.
        Select Case TextBox1.WordWrap
            Case True
                WordWrapToolStripMenuItem.Checked = True
            Case False
                WordWrapToolStripMenuItem.Checked = False
        End Select

    End Sub

    Public Sub SubLineCharCount()

        ToolStripStatusLabel1.Text = (
            "Lines: " & Regex.Split(TextBox1.Text, vbCrLf).Count & ", " &
            "Char: " & (New Regex("\n").Replace(TextBox1.Text, "").Count - Regex.Split(TextBox1.Text, vbCrLf).Count + 1)
            )
        'The above function accurately calculates the number of lines and characters within 'Textbox1'.
        'To get the number of 'Lines' in 'Textbox1' we use 'Regex.Split' to split the text at every 'Newline' (here: vbCrLf) and then we use the '.Count' method to count the number of 'Substrings'.
        'To get the number of 'Characters' in 'Textbox1' we use 'Regex.Replace' to replace every 'Newline' (here: "\n") with empty space ("") and then count the number of 'Characters' minus the number of 'Newlines' + 1.
    End Sub

    Public Sub SubQuickSave()
        'Instantiates a new SaveFileDialog as 'sfd' with the filter as 'Text Documents' with the extention of '.txt'.
        Dim sfd As New SaveFileDialog With {
            .Filter = "Text Documents (*.txt)|*.txt|All files (*.*)|*.*"
            }

        If strSaveLocation IsNot "" Then 'Checks if there is a established location of the file.
            Dim sw As New StreamWriter(strSaveLocation, False) 'Saves to the established location and overwites the contents.
            sw.WriteLine(TextBox1.Text) 'Writes the contents of TextBox1 to the file location.
            sw.Close() 'Closes the Stream Writer
            TextBox1.Modified = False
        End If
    End Sub

    Public Sub SubSaveAs()
        'Instantiates a new SaveFileDialog as 'sfd' with the filter as 'Text Documents' with the extention of '.txt'.
        Dim sfd As New SaveFileDialog With {
            .Filter = "Text Documents (*.txt)|*.txt|All files (*.*)|*.*"
            }

        If sfd.ShowDialog() = DialogResult.OK Then 'Opens the save file dialog and checks if the Dialog result there in is 'OK'.
            Dim sw As New StreamWriter(sfd.FileName, False) 'Initialises an instance of 'StreamWriter' and writes to the location specified in the Save File Dialog. False represents 'Overwrite' as opposed to 'Append'
            strSaveLocation = sfd.FileName 'Sets the user specified location as the Save Location of the file for additional functions.
            sw.WriteLine(TextBox1.Text) 'Writes the contents of the document to the file location specified.
            sw.Close() 'Closes the Stream Writer
            boolSavedDocument = True 'Sets the status of the Document to 'Saved' as it is.
            TextBox1.Modified = False

        ElseIf MsgBoxResult.Cancel Then

        End If
    End Sub

    Public Sub SubOpenFile()
        'Initalises an instance of OpenFileDialog with the filter as 'Text Documents' with the extention of '.txt'.
        Dim ofd As New OpenFileDialog With {
            .Filter = "Text Documents (*.txt)|*.txt|All files (*.*)|*,*" '
            }
        ofd.ShowDialog() 'Opens the OpenFileDialog.
        TextBox1.Text = My.Computer.FileSystem.ReadAllText(ofd.FileName) 'Fills Textbox1 with the contents of the file that was opened.
        strSaveLocation = ofd.FileName 'Sets the established location of this document as the location from which it was opened.
        boolSavedDocument = True 'Sets the status of this document as having an established location.
    End Sub

    Private Sub AboutSuperNotepadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutSuperNotepadToolStripMenuItem.Click
        'Opens the 'About' window.
        Using frmAbout
            frmAbout.ShowDialog()
        End Using

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        'Checks to see if the document is modified
        'If The Document is not modified, try 'Open File'
        'If the Document is modified, throw a 'Yes / No / Cancel' MsgBox at the user.

        'If the user selects 'No', try 'Open File'

        'If the user selects 'Yes', then check if the document has an established save location
        'If the document has a known save location, then 'Quick Save' (w/o dialog)
        'If the document doesn't have a known save location, then 'Save As' (w/ dialog)
        If TextBox1.Modified Then
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes made to this document?", MsgBoxStyle.YesNoCancel, "Open Document")

            Select Case ask
                Case MsgBoxResult.No
                    Try
                        SubOpenFile()
                    Catch ex As Exception
                    End Try
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Yes
                    Select Case boolSavedDocument
                        Case True
                            SubQuickSave()
                            SubOpenFile()
                            TextBox1.Modified = False
                        Case False
                            SubSaveAs()
                            Try
                                SubOpenFile()
                                TextBox1.Modified = False
                            Catch ex As Exception
                            End Try
                    End Select
            End Select

        Else
            Try
                SubOpenFile()
            Catch ex As Exception
            End Try

        End If
    End Sub


    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        'Checks if the notepad has been modified
        'If 'Unmodified', Clear the TextBox
        'If 'Modifed', throw a 'Yes / No / Cancel' Msgbox.

        'If 'No', clear the document, set boolSavedDocument to False.
        'If 'Cancel', then exit this sub routine.
        'If 'Yes', then it throw a Save dialog before clearing the textbox, and setting boolSavedDocument to False.
        If TextBox1.Modified Then
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes made to this document?", MsgBoxStyle.YesNoCancel, "New Document")

            Select Case ask
                Case MsgBoxResult.No
                    TextBox1.Clear()
                    boolSavedDocument = False
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Yes
                    SubSaveAs()
                    If boolSavedDocument = True Then
                        TextBox1.Clear()
                        boolSavedDocument = False
                    End If
            End Select

        Else
            TextBox1.Clear()

        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        'If the document has an established save location, clicking 'Save' will perform a save without going through dialog.
        'If the document does not have an established save location, clicking 'Save' will perform a save via save dialog.
        Select Case boolSavedDocument
            Case False
                SubSaveAs()
            Case True
                SubQuickSave()
        End Select
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        'Clicking 'SaveAs' will save via dialog.
        SubSaveAs()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        'Checks weather is document is modified, if it isn't the program closes.
        'If the document is modified, it throws a 'Yes / No / Cancel' MsgBox at the user.
        'If the user selects 'Yes', it then checks whether the document has an established save location
        'If the document has an established save location, then it performs a 'Quick Save' (w/o Dialog)
        'If the document does not have an established save location, then it performs a 'Save As' (w/ Dialog)
        If TextBox1.Modified = False Then
            End

        ElseIf TextBox1.Modified = True Then
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes made to this document?", MsgBoxStyle.YesNoCancel, "Exit")
            Select Case ask
                Case MsgBoxResult.Yes
                    Select Case boolSavedDocument
                        Case True
                            SubQuickSave()
                            MsgBox("Document saved to: " & strSaveLocation)
                            End
                        Case False
                            SubSaveAs()
                            End
                    End Select
                Case MsgBoxResult.No
                    End
                Case MsgBoxResult.Cancel
                    Exit Sub
                    End
            End Select

        End If
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        'Checks if undo can be performed, if it can it does.
        If TextBox1.CanUndo Then
            TextBox1.Undo()
        Else
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        'Clears the clipboard.
        'If the text is highlighted, it is copied to the clipboard.
        My.Computer.Clipboard.Clear()
        If TextBox1.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(TextBox1.SelectedText)
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        'Clears the clipboard.
        'If the text is highlighted, it is copied to the clipboard.
        'The highlighted text is then deleted from the document
        My.Computer.Clipboard.Clear()
        If TextBox1.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(TextBox1.SelectedText)
        End If
        TextBox1.SelectedText = ""
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Checks to see if the clipboard contains text, if so, it pastes the context of the clipboard.
        If My.Computer.Clipboard.ContainsText Then
            TextBox1.Paste()
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        'Selects all text in Textbox1
        TextBox1.SelectAll()
    End Sub


    Private Sub BtnDebug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        'DEBUG FUNCTION
        'To use set 'btnDebug.Visible' to true.
        'Opens a MsgBox that reports on that tells the user a variety of information.
        MsgBox(
            "Save Location is: " & strSaveLocation & vbCrLf &
            (If(TextBox1.Modified, "Document is modified", "Document is Unmodified")) & vbCrLf &
            (If(boolSavedDocument, "The Document Is Saved", "The Document Is not Saved")) & vbCrLf &
            "Window size is: " & Size.ToString & vbCrLf &
            "Textbox size is: " & TextBox1.Size.ToString & vbCrLf &
            (If(TextBox1.WordWrap, "WordWrap is on", "WordWrap is off"))
            )

    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        'Toggles the visibility of the statusbar as well as the checked status of the menu item.
        Select Case StatusStrip1.Visible
            Case False
                StatusStrip1.Visible = True
            Case True
                StatusStrip1.Visible = False
        End Select
        SubStatusBarStatus()
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        'Opens the Font Dialog so that the user can change a variety of aspects of the font used to type with.
        Dim fd As New FontDialog With {
            .ShowColor = True,
            .Font = TextBox1.Font,
            .Color = TextBox1.ForeColor
        }

        If fd.ShowDialog() <> DialogResult.Cancel Then
            TextBox1.Font = fd.Font
            TextBox1.ForeColor = fd.Color
        End If
    End Sub

    Private Sub WordWrapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordWrapToolStripMenuItem.Click
        'Toggles the Checked status of the WordWrap menu item.
        Select Case TextBox1.WordWrap
            Case True
                TextBox1.WordWrap = False
            Case False
                TextBox1.WordWrap = True
        End Select
        SubWordWrap()
    End Sub

    Private Sub ViewHelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewHelpToolStripMenuItem.Click
        'Opens up the 'Help' window.
        Using FrmHelp
            FrmHelp.ShowDialog()

        End Using

    End Sub
End Class