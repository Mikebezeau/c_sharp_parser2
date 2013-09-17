#Region " Import Statements "

Imports System.IO
Imports System.Xml
Imports System.Xml.Xsl

#End Region

Public Class MainForm
    Inherits System.Windows.Forms.Form

#Region " Member Variables "

    Private mstrInputCSVFile As String = ""
    Private mstrInputXSLTFile As String = ""
    Private mstrOutputXMLFile As String = ""

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtInputCSVFile As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseForInputFile As System.Windows.Forms.Button
    Friend WithEvents btnBrowseForXSLTFile As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnBrowseForOutputFile As System.Windows.Forms.Button
    Friend WithEvents txtOutputXMLFile As System.Windows.Forms.TextBox
    Friend WithEvents txtInputXSLTFile As System.Windows.Forms.TextBox
    Friend WithEvents btnViewFile As System.Windows.Forms.Button
    Friend WithEvents grpFiles As System.Windows.Forms.GroupBox
    Friend WithEvents radCSV As System.Windows.Forms.RadioButton
    Friend WithEvents radXSL As System.Windows.Forms.RadioButton
    Friend WithEvents radXML As System.Windows.Forms.RadioButton
    Friend WithEvents pnlView As System.Windows.Forms.Panel
    Friend WithEvents radUseNotepad As System.Windows.Forms.RadioButton
    Friend WithEvents radUseShellExecute As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MainForm))
        Me.btnConvert = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtInputCSVFile = New System.Windows.Forms.TextBox
        Me.btnBrowseForInputFile = New System.Windows.Forms.Button
        Me.btnBrowseForXSLTFile = New System.Windows.Forms.Button
        Me.txtInputXSLTFile = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnBrowseForOutputFile = New System.Windows.Forms.Button
        Me.txtOutputXMLFile = New System.Windows.Forms.TextBox
        Me.btnViewFile = New System.Windows.Forms.Button
        Me.grpFiles = New System.Windows.Forms.GroupBox
        Me.pnlView = New System.Windows.Forms.Panel
        Me.radUseNotepad = New System.Windows.Forms.RadioButton
        Me.radUseShellExecute = New System.Windows.Forms.RadioButton
        Me.radXML = New System.Windows.Forms.RadioButton
        Me.radXSL = New System.Windows.Forms.RadioButton
        Me.radCSV = New System.Windows.Forms.RadioButton
        Me.grpFiles.SuspendLayout()
        Me.pnlView.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnConvert
        '
        Me.btnConvert.BackColor = System.Drawing.Color.LightGray
        Me.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnConvert.ForeColor = System.Drawing.Color.Black
        Me.btnConvert.Image = CType(resources.GetObject("btnConvert.Image"), System.Drawing.Image)
        Me.btnConvert.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnConvert.Location = New System.Drawing.Point(392, 200)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(96, 80)
        Me.btnConvert.TabIndex = 7
        Me.btnConvert.Text = "Convert the Data to XML"
        Me.btnConvert.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(336, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Comma Delimited File to Transform:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtInputCSVFile
        '
        Me.txtInputCSVFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInputCSVFile.BackColor = System.Drawing.Color.WhiteSmoke
        Me.txtInputCSVFile.Location = New System.Drawing.Point(16, 34)
        Me.txtInputCSVFile.Name = "txtInputCSVFile"
        Me.txtInputCSVFile.Size = New System.Drawing.Size(400, 20)
        Me.txtInputCSVFile.TabIndex = 1
        Me.txtInputCSVFile.Text = ""
        '
        'btnBrowseForInputFile
        '
        Me.btnBrowseForInputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseForInputFile.BackColor = System.Drawing.Color.LightGray
        Me.btnBrowseForInputFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBrowseForInputFile.ForeColor = System.Drawing.Color.Black
        Me.btnBrowseForInputFile.Location = New System.Drawing.Point(420, 32)
        Me.btnBrowseForInputFile.Name = "btnBrowseForInputFile"
        Me.btnBrowseForInputFile.Size = New System.Drawing.Size(68, 24)
        Me.btnBrowseForInputFile.TabIndex = 2
        Me.btnBrowseForInputFile.Text = "Browse..."
        Me.btnBrowseForInputFile.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnBrowseForXSLTFile
        '
        Me.btnBrowseForXSLTFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseForXSLTFile.BackColor = System.Drawing.Color.LightGray
        Me.btnBrowseForXSLTFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBrowseForXSLTFile.ForeColor = System.Drawing.Color.Black
        Me.btnBrowseForXSLTFile.Location = New System.Drawing.Point(420, 80)
        Me.btnBrowseForXSLTFile.Name = "btnBrowseForXSLTFile"
        Me.btnBrowseForXSLTFile.Size = New System.Drawing.Size(68, 24)
        Me.btnBrowseForXSLTFile.TabIndex = 4
        Me.btnBrowseForXSLTFile.Text = "Browse..."
        Me.btnBrowseForXSLTFile.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtInputXSLTFile
        '
        Me.txtInputXSLTFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInputXSLTFile.BackColor = System.Drawing.Color.WhiteSmoke
        Me.txtInputXSLTFile.Location = New System.Drawing.Point(16, 82)
        Me.txtInputXSLTFile.Name = "txtInputXSLTFile"
        Me.txtInputXSLTFile.Size = New System.Drawing.Size(400, 20)
        Me.txtInputXSLTFile.TabIndex = 3
        Me.txtInputXSLTFile.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(336, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "XSL File Used to Transform the Data:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(16, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(336, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Name of Transformed XML Output File:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'btnBrowseForOutputFile
        '
        Me.btnBrowseForOutputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseForOutputFile.BackColor = System.Drawing.Color.LightGray
        Me.btnBrowseForOutputFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBrowseForOutputFile.ForeColor = System.Drawing.Color.Black
        Me.btnBrowseForOutputFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnBrowseForOutputFile.Location = New System.Drawing.Point(420, 128)
        Me.btnBrowseForOutputFile.Name = "btnBrowseForOutputFile"
        Me.btnBrowseForOutputFile.Size = New System.Drawing.Size(68, 24)
        Me.btnBrowseForOutputFile.TabIndex = 6
        Me.btnBrowseForOutputFile.Text = "Browse..."
        Me.btnBrowseForOutputFile.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtOutputXMLFile
        '
        Me.txtOutputXMLFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutputXMLFile.BackColor = System.Drawing.Color.WhiteSmoke
        Me.txtOutputXMLFile.Location = New System.Drawing.Point(16, 130)
        Me.txtOutputXMLFile.Name = "txtOutputXMLFile"
        Me.txtOutputXMLFile.Size = New System.Drawing.Size(400, 20)
        Me.txtOutputXMLFile.TabIndex = 5
        Me.txtOutputXMLFile.Text = ""
        '
        'btnViewFile
        '
        Me.btnViewFile.BackColor = System.Drawing.Color.LightGray
        Me.btnViewFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewFile.Location = New System.Drawing.Point(16, 16)
        Me.btnViewFile.Name = "btnViewFile"
        Me.btnViewFile.Size = New System.Drawing.Size(88, 23)
        Me.btnViewFile.TabIndex = 1
        Me.btnViewFile.Text = "Show it to me"
        '
        'grpFiles
        '
        Me.grpFiles.Controls.Add(Me.pnlView)
        Me.grpFiles.Controls.Add(Me.radXML)
        Me.grpFiles.Controls.Add(Me.radXSL)
        Me.grpFiles.Controls.Add(Me.radCSV)
        Me.grpFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFiles.Location = New System.Drawing.Point(16, 184)
        Me.grpFiles.Name = "grpFiles"
        Me.grpFiles.Size = New System.Drawing.Size(360, 104)
        Me.grpFiles.TabIndex = 11
        Me.grpFiles.TabStop = False
        Me.grpFiles.Text = "What File Do You Want To See?"
        '
        'pnlView
        '
        Me.pnlView.Controls.Add(Me.radUseNotepad)
        Me.pnlView.Controls.Add(Me.radUseShellExecute)
        Me.pnlView.Controls.Add(Me.btnViewFile)
        Me.pnlView.Location = New System.Drawing.Point(96, 32)
        Me.pnlView.Name = "pnlView"
        Me.pnlView.Size = New System.Drawing.Size(256, 56)
        Me.pnlView.TabIndex = 13
        '
        'radUseNotepad
        '
        Me.radUseNotepad.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radUseNotepad.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radUseNotepad.Location = New System.Drawing.Point(120, 8)
        Me.radUseNotepad.Name = "radUseNotepad"
        Me.radUseNotepad.Size = New System.Drawing.Size(96, 16)
        Me.radUseNotepad.TabIndex = 4
        Me.radUseNotepad.Text = "Using Notepad"
        '
        'radUseShellExecute
        '
        Me.radUseShellExecute.Checked = True
        Me.radUseShellExecute.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radUseShellExecute.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radUseShellExecute.Location = New System.Drawing.Point(120, 32)
        Me.radUseShellExecute.Name = "radUseShellExecute"
        Me.radUseShellExecute.Size = New System.Drawing.Size(128, 16)
        Me.radUseShellExecute.TabIndex = 5
        Me.radUseShellExecute.TabStop = True
        Me.radUseShellExecute.Text = "Let Windows Decide"
        '
        'radXML
        '
        Me.radXML.Checked = True
        Me.radXML.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radXML.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radXML.Location = New System.Drawing.Point(24, 72)
        Me.radXML.Name = "radXML"
        Me.radXML.Size = New System.Drawing.Size(72, 16)
        Me.radXML.TabIndex = 0
        Me.radXML.TabStop = True
        Me.radXML.Text = " XML File"
        '
        'radXSL
        '
        Me.radXSL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radXSL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radXSL.Location = New System.Drawing.Point(24, 48)
        Me.radXSL.Name = "radXSL"
        Me.radXSL.Size = New System.Drawing.Size(72, 16)
        Me.radXSL.TabIndex = 3
        Me.radXSL.TabStop = True
        Me.radXSL.Text = "XSL File"
        '
        'radCSV
        '
        Me.radCSV.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.radCSV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radCSV.Location = New System.Drawing.Point(24, 24)
        Me.radCSV.Name = "radCSV"
        Me.radCSV.Size = New System.Drawing.Size(72, 16)
        Me.radCSV.TabIndex = 2
        Me.radCSV.TabStop = True
        Me.radCSV.Text = "CSV File"
        '
        'MainForm
        '
        Me.AcceptButton = Me.btnConvert
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(504, 309)
        Me.Controls.Add(Me.grpFiles)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnBrowseForOutputFile)
        Me.Controls.Add(Me.txtOutputXMLFile)
        Me.Controls.Add(Me.txtInputXSLTFile)
        Me.Controls.Add(Me.txtInputCSVFile)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBrowseForXSLTFile)
        Me.Controls.Add(Me.btnBrowseForInputFile)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnConvert)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(512, 336)
        Me.Name = "MainForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Conversion Utility"
        Me.grpFiles.ResumeLayout(False)
        Me.pnlView.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click

        mstrInputCSVFile = txtInputCSVFile.Text.Trim()
        mstrInputXSLTFile = txtInputXSLTFile.Text.Trim()
        mstrOutputXMLFile = txtOutputXMLFile.Text.Trim()

        If mstrInputCSVFile = "" Then
            MessageBox.Show("Please specify the comma delimited file to convert.", "Missing File Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtInputCSVFile.Focus()
            Return
        End If

        If mstrInputXSLTFile = "" Then
            MessageBox.Show("Please specify the XSL transform file.", "Missing File Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtInputXSLTFile.Focus()
            Return
        End If

        If mstrOutputXMLFile = "" Then
            MessageBox.Show("Please specify the output XML file.", "Missing File Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtOutputXMLFile.Focus()
            Return
        End If

        If Not File.Exists(mstrInputCSVFile) Then
            MessageBox.Show("Comma delimited file not found.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtInputCSVFile.Focus()
            Return
        End If

        If Not File.Exists(mstrInputXSLTFile) Then
            MessageBox.Show("XSL transform file not found.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtInputXSLTFile.Focus()
            Return
        End If


        Try
            Cursor.Current = Cursors.WaitCursor

            '
            ' Convert the CSV file to an XML document by 
            ' surrounding it with <ROOT></ROOT> tags.
            '
            Dim rdr As New StreamReader(mstrInputCSVFile)
            Dim strLine As String = rdr.ReadLine

            If strLine = "<ROOT>" Then
                rdr.Close()
            Else
                Dim strFile As String = rdr.ReadToEnd
                rdr.Close()

                Dim wtr As New StreamWriter(mstrInputCSVFile, False)
                wtr.WriteLine("<ROOT>")
                wtr.WriteLine(strLine)
                wtr.WriteLine(strFile)
                wtr.WriteLine("</ROOT>")
                wtr.Close()
                wtr = Nothing
                strFile = Nothing
            End If
            rdr = Nothing

            '
            ' Create a resolver with default credentials.
            '
            Dim resolver As XmlUrlResolver = New XmlUrlResolver
            resolver.Credentials = System.Net.CredentialCache.DefaultCredentials

            '
            ' Create the XSL Transform object.
            '
            Dim XSLT As XslTransform = New XslTransform

            '
            ' Load the stylesheet.
            '
            XSLT.Load(mstrInputXSLTFile, resolver)

            '
            ' Transform the file. 
            '
            XSLT.Transform(mstrInputCSVFile, mstrOutputXMLFile, resolver)
            Cursor.Current = Cursors.Default
            MessageBox.Show("Data has been successfully converted.", "Conversion Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

            XSLT = Nothing

        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MessageBox.Show(ex.Message, "Error Converting Data", MessageBoxButtons.OK, MessageBoxIcon.Stop)

        End Try

    End Sub

    Private Sub btnBrowseForInputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseForInputFile.Click
        Dim strFileName As String
        Dim ofd As New OpenFileDialog

        ofd.CheckFileExists = True
        ofd.Filter = "Comma Delimited Files (*.csv) | *.csv"
        ofd.Title = "Select the CSV File"
        ofd.Multiselect = False

        ofd.ShowDialog()
        mstrInputCSVFile = ofd.FileName
        txtInputCSVFile.Text = mstrInputCSVFile

    End Sub

    Private Sub btnBrowseForXSLTFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseForXSLTFile.Click
        Dim strFileName As String
        Dim ofd As New OpenFileDialog

        ofd.CheckFileExists = True
        ofd.Filter = "XSL Transform File (*.xslt) | *.xslt"
        ofd.Title = "Select the Transform File"
        ofd.Multiselect = False

        ofd.ShowDialog()
        mstrInputXSLTFile = ofd.FileName
        txtInputXSLTFile.Text = mstrInputXSLTFile

    End Sub

    Private Sub btnBrowseForOutputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseForOutputFile.Click
        Dim strFileName As String
        Dim sfd As New SaveFileDialog

        sfd.OverwritePrompt = True
        sfd.CheckFileExists = False
        sfd.Filter = "XML File (*.xml) | *.xml"
        sfd.Title = "Save XML Output As"

        sfd.ShowDialog()
        mstrOutputXMLFile = sfd.FileName
        txtOutputXMLFile.Text = mstrOutputXMLFile

    End Sub

    Private Sub btnViewFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewFile.Click

        Try
            Dim strFile As String
            Dim blnViewXML As Boolean = False

            If radCSV.Checked Then
                strFile = mstrInputCSVFile
            ElseIf radXSL.Checked Then
                strFile = mstrInputXSLTFile
            Else
                strFile = mstrOutputXMLFile
                blnViewXML = True
            End If

            If strFile = "" Then
                Return
            End If

            If Not File.Exists(strFile) Then
                Return
            End If

            Dim startInfo As New ProcessStartInfo
            startInfo.WindowStyle = ProcessWindowStyle.Normal
            startInfo.ErrorDialog = False


            If radUseNotepad.Checked Then
                startInfo.FileName = "Notepad.exe "
                startInfo.Arguments = strFile
                startInfo.UseShellExecute = False
            Else
                startInfo.FileName = strFile
                startInfo.UseShellExecute = True

                If blnViewXML Then
                    If MessageBox.Show("Large XML files may take several minutes to load." & _
                            ControlChars.CrLf & "You may want to view it using Notepad instead." & _
                            ControlChars.CrLf & ControlChars.CrLf & "Do you want to continue?", _
                            "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        Return
                    End If
                End If
            End If

            Process.Start(startInfo)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Unable to View File", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try

    End Sub

End Class
