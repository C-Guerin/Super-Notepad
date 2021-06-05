Imports System.Text
Imports System.Text.RegularExpressions

Public Class FrmHelp

    Private Sub FrmHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.FixedSingle
        MinimizeBox = False
        MaximizeBox = False
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim fntRegular As New Font(RichTextBox1.Font, FontStyle.Regular)
        Dim fntBold As New Font(RichTextBox1.Font, FontStyle.Bold)
        TreeView1.SelectedNode.Expand()

        Select Case e.Node.Name
            Case "Node1" 'Opening Files
                RichTextBox1.Clear()
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText("To open a file go to ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("'File' -> 'Open' ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(", or you can press ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("'Ctrl + O'.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("Once the ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("'Open File' ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText("Dialog appears, you may navigate to and select any ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("'.Txt' ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText("file that you wish to open.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("If you have any unsaved work, you will be prompted to save your work.")

            Case "Node2" 'Saving Files
                RichTextBox1.Clear()
                RichTextBox1.AppendText("To save a file go to")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText(" 'File' -> 'Save' ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(", or you can press ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText(" Ctrl + S.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("The first time you save, a save window will open." & vbCrLf & "Navigate to the desired location, and name the file.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("The next time you save, the save will write to the previous location without opening any windows.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("If you with to save the file under a different name, or at a different location select ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("'Save As'.")

            Case "Node3" 'Creating New Files
                RichTextBox1.Clear()
                RichTextBox1.AppendText("To create a new file simply select")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText(" 'File' -> 'New' ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(", or you can press")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText(" Ctrl + N.")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("If you have made unsaved changes to the document, you will be prompted to save before creating a new file.")


            Case "Node4" 'Customising Font
                RichTextBox1.Clear()
                RichTextBox1.AppendText("To customise the font used to type with go to")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText(" 'Format' -> 'Font'.")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("Here are a list of a few things that you can customise:" & vbCrLf)
                RichTextBox1.AppendText(">Font (Arial, Times New Roman, etc.)" & vbCrLf)
                RichTextBox1.AppendText(">Font Style (Bold, Italic, etc.)" & vbCrLf)
                RichTextBox1.AppendText(">Font Size (8pt, 12pt, etc.)" & vbCrLf)
                RichTextBox1.AppendText(">Strikeout and Underline" & vbCrLf)
                RichTextBox1.AppendText(">Font Colour")

            Case "Node5" 'WordWrap
                RichTextBox1.Clear()
                RichTextBox1.AppendText("WordWrap confines the text to the width of the application.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("You can turn word wrap off if you would rather scroll horizontally and set the placement of text yourself.")

            Case "Node6" 'Statusbar
                RichTextBox1.Clear()
                RichTextBox1.AppendText("The Statusbar refers to the bar at the bottom of the screen.")
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("The Statusbar shows the number of ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("Characters ")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText("in the document, as well as the number of ")
                RichTextBox1.SelectionFont = fntBold
                RichTextBox1.AppendText("Lines.")
                RichTextBox1.SelectionFont = fntRegular
                RichTextBox1.AppendText(vbCrLf & vbCrLf)
                RichTextBox1.AppendText("The statusbar will only show the number of lines created by the user, not wordwrap.")
        End Select




    End Sub
End Class