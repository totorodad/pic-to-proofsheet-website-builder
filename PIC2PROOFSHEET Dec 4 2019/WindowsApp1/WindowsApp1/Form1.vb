Imports System.Drawing.Imaging
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Security.Cryptography

Public Class Form1

    Public RenamedCompleted As Boolean = False

    'get the directory of where the images are
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If

        'load the .jpg and .JPG into the SOURCE PIC NAMES listbox
        UpdateSourceFileList()
        UpdateDestinationFileList()

        SaveParameters()

    End Sub

    'get the destination folder
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then

            ' set the 
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
        End If

        UpdateDestinationFileList()

        'enable the ability to copy the pics from source to destination
        COPYPICS_Button8.Enabled = True

        SaveParameters()

    End Sub

    'Collect source directory images
    Private Sub UpdateSourceFileList()
        Dim FileString As String
        Dim PathFileName As String

        'clear out the old listbox pic names if they are there
        ListBox1.Items.Clear()

        Try
            'fine all jpg or JPG files and populate them in the source file list (listbox1)
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(TextBox1.Text)
                FileString = foundFile.ToString
                If (FileString.EndsWith("JPG") = True Or FileString.EndsWith("jpg") = True) Then
                    PathFileName = TextBox1.Text & "\" & Path.GetFileName(foundFile)
                    ListBox1.Items.Add(PathFileName)
                End If
            Next

            'show the count
            Label3.Text = ListBox1.Items.Count.ToString & " Pic's"

        Catch ex As Exception
            MsgBox("Sorry the source directory does not exist")
        End Try
    End Sub

    'collect pictures button
    Private Sub UpdateDestinationFileList()
        Dim PathFileName As String
        Dim ProcessCount As Integer = 1
        'If the destination address is not empty then populate the destination list
        If Not TextBox2.Text = Nothing Then
            'clear out the old listbox pic names if they are there
            ListBox3.Items.Clear()

            'inform the user how much time they will need to get coffee
            PROCESSING_Label3.Visible = True

            'parse thru the images names in listbox1 and show the new names in listbox3
            For Each FileName As String In ListBox1.Items
                PathFileName = TextBox1.Text & "\" & Path.GetFileName(FileName)
                InsertDestinationFileName(ListBox3, PathFileName)

                If ProcessCount Mod 5 = 0 Then
                    PROCESSING_Label3.Text = "Processing " & ProcessCount & " of " & ListBox1.Items.Count.ToString
                    Application.DoEvents()

                End If
                ProcessCount = ProcessCount + 1

            Next
                PROCESSING_Label3.Visible = False
            End If

    End Sub

    'take the source file name and insert it into the destination list (with changes as needed)
    Sub InsertDestinationFileName(ByRef DestinationListBox As ListBox, ByVal PathFileName As String)
        Dim file_name_count_str As String = ""
        Dim file_name_count As Integer = 0
        Dim new_name As String
        Dim YYYYMMStr As String = ""
        Dim NewPathFileName As String

        'Set the new name as the old name as the default case
        new_name = Path.GetFileNameWithoutExtension(PathFileName)
        NewPathFileName = TextBox2.Text & "\pics\" & new_name & ".jpg"


        'convert the name if EXIF is checked
        If (EXIF_CheckBox4.Checked = True) Then
            new_name = Path.GetFileNameWithoutExtension(EXIFGetDate(PathFileName))

            'if we are sorting the images into YYYYMM folders then prepend the YYYYMM folder name
            If YYYYMM_CheckBox5.Checked = True Then
                YYYYMMStr = "\" & Mid(new_name, 1, 4) & "_" & Mid(new_name, 5, 2)
            End If

            'if EXIF could read the date it was set to 1974.  Rename it go to the MISC folder or rename with proper YYYYMM string
            If String.Compare("1974", Mid(new_name, 1, 4)) Then
                NewPathFileName = TextBox2.Text & "\pics" & YYYYMMStr & "\" & new_name & file_name_count_str & ".jpg"
            Else
                NewPathFileName = TextBox2.Text & "\pics\misc\" & Path.GetFileNameWithoutExtension(PathFileName) & ".jpg"
            End If

            'check for duplicates and add an _index number at the end to make unique filename
            While (Not DestinationListBox.FindString(NewPathFileName))
                    file_name_count = file_name_count + 1
                    file_name_count_str = "_" & file_name_count.ToString
                    NewPathFileName = TextBox2.Text & "\pics" & YYYYMMStr & "\" & new_name & file_name_count_str & ".jpg"
                End While

            ElseIf Album_CheckBox3.Checked = True Then
                new_name = Path.GetFileNameWithoutExtension(PathFileName)
            NewPathFileName = TextBox2.Text & "\pics\" & TextBox4.Text & "\" & new_name & ".jpg"
        End If

        'do the add
        DestinationListBox.Items.Add(NewPathFileName)
    End Sub

    'read the EXIF data from the .jpg if possible
    Function EXIFGetDate(ByVal PathFileName As String) As String
        Dim DateTime = &H132 '306

        Dim image As Bitmap = New Bitmap(PathFileName)
        Dim pic_time As PropertyItem
        Dim Str_pic_time As String
        Dim YYYYMMDDHHMMSS As String

        Dim propItems As PropertyItem() = image.PropertyItems

        'Check to see if the jpg has the DateTime EXIF data
        Try
            pic_time = image.GetPropertyItem(DateTime)
            Str_pic_time = Encoding.ASCII.GetString(pic_time.Value, 0, pic_time.Len - 1)
        Catch
            Str_pic_time = "1974:01:01 00:00:00"
        End Try

        YYYYMMDDHHMMSS = Mid(Str_pic_time, 1, 4) + Mid(Str_pic_time, 6, 2) + Mid(Str_pic_time, 9, 2) + "-" + Mid(Str_pic_time, 12, 2) + Mid(Str_pic_time, 15, 2) + Mid(Str_pic_time, 18, 2)
        image.Dispose()

        Return (YYYYMMDDHHMMSS)
    End Function

    'clear the buttons and list boxes

    Sub ClearForm()
        'enable the Collect pics button
        'Button2.Enabled = True

        'clear the source pic names
        ListBox1.Items.Clear()

        'clear the destination pic names
        ListBox3.Items.Clear()

        'clear the preview picture
        PictureBox2.Image = Nothing

        'clear the preview name textbox
        TextBox3.Text = ""

        'turn off the rotate buttons
        Button7.Enabled = False
        Button9.Enabled = False

        'turn off the picture forward and reverse buttons
        Button10.Enabled = False
        Button11.Enabled = False

        'turn off the make proofs button
        MAKEPROOFS_Button5.Enabled = False

        'turn off the Delete Destination Pic button
        DELETE_SELECTED_PIC_Button8.Enabled = False

        'reset the RenamedCompleted
        RenamedCompleted = False

    End Sub

    'copy and rename as needed the images in the destination folder
    Private Sub CopyPics()
        Dim PathName As String
        Dim listbox_count As Integer = 0
        Dim SourcePathFileName As String
        Dim processCount As Integer = 1
        Dim CreatedDirectoryCount As Integer = 0

        'inform the user how much time they will need to get coffee
        PROCESSING_Label3.Text = "COPYING.."
        PROCESSING_Label3.Visible = True
        Application.DoEvents()

        'check to see if the /pics directory exists
        'Add /pics directory
        PathName = Path.GetDirectoryName(TextBox2.Text)
        'If Not My.Computer.FileSystem.DirectoryExists(PathName + "\pics") Then
        'System.IO.Directory.CreateDirectory(PathName + "\pics")
        'End If

        'copy all source directory .jpg's and rename with DateTimeFileNames
        For Each PathFileName As String In ListBox3.Items

            If processCount Mod 5 = 0 Then
                PROCESSING_Label3.Text = "Copying " & processCount & " of " & ListBox3.Items.Count.ToString
                Application.DoEvents()
            End If

            'Check to see if the Destination Directory exists.  If no create it
            PathName = Path.GetDirectoryName(PathFileName)

            If Not My.Computer.FileSystem.DirectoryExists(PathName) Then
                System.IO.Directory.CreateDirectory(PathName)
                CreatedDirectoryCount = CreatedDirectoryCount + 1
            End If

            'get the source file name to copy
            SourcePathFileName = ListBox1.Items(listbox_count)

            'See if the destination FileName has already been copied or already exist
            If (Not My.Computer.FileSystem.FileExists(PathFileName)) Then
                My.Computer.FileSystem.CopyFile(SourcePathFileName, PathFileName)
            End If
            listbox_count = listbox_count + 1
        Next
        PROCESSING_Label3.Visible = False

        MsgBox("Copy Complete " & CreatedDirectoryCount.ToString & " directories created.")
    End Sub

    Public Sub LoadThumbnail()
        Dim DestinationPath As String = TextBox2.Text
        Dim SourcePath As String = TextBox1.Text
        Dim PathFileName As String

        PROCESSING_Label3.Text = "Loading Pic.."
        PROCESSING_Label3.Visible = True
        Application.DoEvents()

        'check that an item is selected in the destination listbox and then display the thumbnail
        If ListBox3.SelectedIndex >= 0 Then
            Dim SourcePathFileName As String = ListBox1.SelectedItem.ToString
            Dim DestinationPathFileName As String = ListBox3.SelectedItem.ToString

            'Check to see if the file exists if not then load the image from the source directory
            If My.Computer.FileSystem.FileExists(DestinationPathFileName) Then
                PathFileName = DestinationPathFileName
            Else
                PathFileName = SourcePathFileName
            End If

            'set preview thumb name
            TextBox3.Text = Path.GetFileName(SourcePathFileName) & " --> " & Path.GetFileName(DestinationPathFileName)

            'load the image
            Dim image1 As Bitmap = CType(Image.FromFile(PathFileName, True), Bitmap)

            'Scale the image to fit into the Picturebox
            Dim ScaleFactor As Double = 1.0

            'If the Width is greater than height (landscape)
            If image1.Width > image1.Height Then
                If image1.Width < PictureBox2.Width Then
                    ScaleFactor = PictureBox2.Width / image1.Width
                ElseIf image1.Width > PictureBox2.Width Then
                    ScaleFactor = image1.Width / PictureBox2.Width
                End If
            End If

            'If the Width is less than height (portrait)
            If image1.Width < image1.Height Then
                If image1.Height < PictureBox2.Height Then
                    ScaleFactor = PictureBox2.Height / image1.Height
                ElseIf image1.Height > PictureBox2.Height Then
                    ScaleFactor = image1.Height / PictureBox2.Height
                End If
            End If

            Dim ResizedImage As Bitmap = New Bitmap(image1, New Size(image1.Width / ScaleFactor, image1.Height / ScaleFactor))

            PictureBox2.Image = ResizedImage

            image1.Dispose()
        End If

        PROCESSING_Label3.Visible = False

    End Sub

    'lock selection with before and after file names
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox3.Items.Count > 0 Then
            ListBox3.SelectedIndex = ListBox1.SelectedIndex
            LoadThumbnail()
            PICBROWSER_GroupBox2.Enabled = True

            'turn on the rotate buttons if an item is selected
            Button7.Enabled = True
            Button9.Enabled = True

            'Turn on pic up/down buttons
            Button10.Enabled = True
            Button11.Enabled = True
        End If
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        If ListBox1.Items.Count > 0 Then
            ListBox1.SelectedIndex = ListBox3.SelectedIndex
            LoadThumbnail()
            PICBROWSER_GroupBox2.Enabled = True

            'turn on the rotate buttons if an item is selected
            Button7.Enabled = True
            Button9.Enabled = True

            'Turn on pic up/down buttons
            Button10.Enabled = True
            Button11.Enabled = True
        End If
    End Sub

    'Rotate image Left
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Left As Integer = 0
        RotateImage(Left)

        'show the selected listbox item in the preview pane
        LoadThumbnail()
    End Sub

    'Rotate Image Right
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim Right As Integer = 1
        RotateImage(Right)

        'show the selected listbox item in the preview pane
        LoadThumbnail()
    End Sub

    Private Sub RotateImage(ByVal RotateDirection As Integer)
        Dim Left As Integer = 0
        Dim Right As Integer = 1

        Dim DestinationFileName As String = ListBox3.SelectedItem.ToString
        Dim DestinationPath As String = Path.GetDirectoryName(DestinationFileName)

        'delete the tmp.jpg file if it exists
        If (My.Computer.FileSystem.FileExists(DestinationPath & "\" & "tmp.jpg")) Then
            My.Computer.FileSystem.DeleteFile(DestinationPath & "\" & "tmp.jpg")
        End If

        'Copy the destination image to a tmp.jpg
        My.Computer.FileSystem.CopyFile(DestinationFileName, DestinationPath & "\" & "tmp.jpg")

        'delete the orginal
        My.Computer.FileSystem.DeleteFile(DestinationFileName)

        'rotate the tmp.jpg and store it under the origianl file name
        Dim bitmap As Bitmap = CType(Bitmap.FromFile(DestinationPath & "\" & "tmp.jpg"), Bitmap)
        If bitmap IsNot Nothing Then
            If (RotateDirection = Left) Then
                bitmap.RotateFlip(RotateFlipType.Rotate270FlipNone)
            Else
                bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone)
            End If

            'write out he newly rotated pic
            bitmap.Save(DestinationFileName, System.Drawing.Imaging.ImageFormat.Jpeg)
        End If

        'delete the tmp.jpg file if it exists
        If My.Computer.FileSystem.FileExists(DestinationPath & "\" & "tmp.jpg") Then
            My.Computer.FileSystem.DeleteFile(DestinationPath & "\" & "tmp.jpg")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'start with the YYYYMM Checkbox disabled
        YYYYMM_CheckBox5.Enabled = False
        LoadParameters()

    End Sub

    Private Sub LoadParameters()
        Dim strPath As String = My.Application.Info.DirectoryPath
        Dim FilePath As String = strPath & "\paths.ini"

        If File.Exists(FilePath) Then
            Using reader As StreamReader = New StreamReader(FilePath)
                TextBox1.Text = reader.ReadLine 'read source path
                TextBox2.Text = reader.ReadLine 'read destination path
                TextBox5.Text = reader.ReadLine 'read archive path
            End Using

            'UpdateSourceFileList()
            'UpdateDestinationFileList()
        End If

    End Sub

    Private Sub SaveParameters()
        Dim strPath As String = My.Application.Info.DirectoryPath

        'if the file exist then delete it
        Dim DestinationFile As String = strPath & "\paths.ini"
        If (My.Computer.FileSystem.FileExists(DestinationFile)) Then
            My.Computer.FileSystem.DeleteFile(DestinationFile)
        End If

        Dim file = My.Computer.FileSystem.OpenTextFileWriter(strPath & "\paths.ini", True)
        file.WriteLine(TextBox1.Text) 'store source path
        file.WriteLine(TextBox2.Text) 'store destination path
        file.WriteLine(TextBox5.Text) 'store archvie path
        file.Close()
    End Sub

    'draw proof sheet button
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles MAKEPROOFS_Button5.Click
        Dim PathName As String

        PROCESSING_Label3.Visible = True
        PROCESSING_Label3.Text = "PROCESSING.."
        Application.DoEvents()

        'clear out our path's we've processed listbox
        ListBox4.Items.Clear()
        ListBox4.Items.Add(Path.GetDirectoryName(ListBox3.Items(0))) 'add the first item in our "have we processed listbox

        'process the first path directory
        PathName = ListBox4.Items(0)
        PROCESSING_Label3.Text = "Processing: " & PathName
        Application.DoEvents()

        MakeProofSheetsInPath(PathName)

        'now process all other directory names if they are unique
        For Each FilePathName As String In ListBox3.Items
            PathName = Path.GetDirectoryName(FilePathName)
            If Not ListBox4.Items.Contains(PathName) Then
                ListBox4.Items.Add(PathName) 'add the new path found

                PROCESSING_Label3.Text = "Processing: " & PathName
                Application.DoEvents()

                'Make the proofs for theat path
                MakeProofSheetsInPath(PathName)
            End If
        Next

        PROCESSING_Label3.Visible = False
        MsgBox("Proofs Created!")
    End Sub

    Private Sub MakeProofSheetsInPath(ByVal DestinationPath As String)

        Dim BlackPen As New Pen(Color.FromArgb(255, 0, 0, 0))
        Dim WhiteBrush As New SolidBrush(Color.White)
        Dim proofx As Integer = 1024
        Dim proofy As Integer = 768

        Dim proofsheet = New Bitmap(proofx, proofy)

        Dim MyGraphics As Graphics = Graphics.FromImage(proofsheet)

        MyGraphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        MyGraphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        MyGraphics.Clear(Color.White)

        'copy all source directory .jpg's and rename with DateTimeFileNames
        Dim ScaleFactor As Single
        Dim OffsetX As Integer
        Dim OffsetY As Integer
        Dim FileName As String
        Dim count As Integer = 0
        Dim EdgeOffsetX = 50
        Dim EdgeOffsetY = 10
        Dim ThumbX As Integer = 135
        Dim ThumbY As Integer = 90
        Dim GapX As Integer = 25
        Dim GapY As Integer = 35
        Dim PortraitOffsetX As Integer = 32 'how much portrait images are shifted over to center them
        Dim done As Boolean
        Dim rect As New Rectangle(0, 0, proofx - 1, proofy - 1)
        Dim ProofSheetCount As Integer
        Dim x As Integer
        Dim y As Integer
        Dim AtLeastOneImageDrawn As Boolean
        Dim FileString As String

        'clear out listbox2 which will store all the pic names to publish into proofs in the destinationpath
        ListBox2.Items.Clear()

        'fine all jpg or JPG files and populate them in the source file list (listbox1)
        For Each FileFound As String In my.computer.filesystem.getfiles(DestinationPath)
            FileString = FileFound.ToString
            If (FileString.EndsWith("JPG") = True Or FileString.EndsWith("jpg") = True) Then
                ListBox2.Items.Add(FileString)
            End If
        Next

        done = False ' not done makeing proof sheets
        count = 0
        ProofSheetCount = 1

        'Start HTML output
        HTML_New(DestinationPath)

        HTML_Start(DestinationPath, ListBox2)


        'draw small images on the picturebox1
        While (Not done)

            AtLeastOneImageDrawn = False

            'blank out proofsheet drawing area and add black border
            MyGraphics.FillRectangle(WhiteBrush, rect)
            MyGraphics.DrawRectangle(BlackPen, rect)


            If count <= ListBox2.Items.Count - 1 Then
                HTML_SRC_1(DestinationPath, ProofSheetCount)
            End If


            ' for current sized Thumbs create 6x6 grid
            For y = 0 To 5
                For x = 0 To 5
                    If count <= ListBox2.Items.Count - 1 Then

                        FileName = ListBox2.Items(count)
                        Dim SourceImage As Bitmap = CType(Bitmap.FromFile(FileName), Bitmap)

                        ScaleFactor = CalculateScaleFactor(SourceImage, ThumbX, ThumbY)

                        OffsetX = EdgeOffsetX + (x * (ThumbX + GapX))
                        OffsetY = EdgeOffsetY + (y * (ThumbY + GapY))

                        HTML_SRC_2A(DestinationPath, OffsetX, OffsetY, OffsetX + ThumbX, OffsetY + ThumbY, Path.GetFileName(FileName), count, SourceImage.Width, SourceImage.Height)

                        Dim DestinationBitmap As Bitmap = New Bitmap(CInt(SourceImage.Width * ScaleFactor), CInt(SourceImage.Height * ScaleFactor))

                        'if Image is portrait
                        If (SourceImage.Height > SourceImage.Width) Then
                            MyGraphics.DrawImage(SourceImage, OffsetX + PortraitOffsetX, OffsetY, DestinationBitmap.Width + 1, DestinationBitmap.Height + 1)

                        Else
                            'image is standard 1.5x1 ratio landscape
                            MyGraphics.DrawImage(SourceImage, OffsetX, OffsetY, DestinationBitmap.Width + 1, DestinationBitmap.Height + 1)
                        End If

                        ' Add title to each image
                        Dim textFont As New Font("Arial", 10, FontStyle.Bold)
                        Dim textBrush As New SolidBrush(Color.Black)
                        Dim textPoint As New Point(OffsetX + 1, OffsetY + ThumbY + 12)
                        MyGraphics.DrawString(Path.GetFileName(FileName), textFont, textBrush, textPoint)
                        textFont.Dispose()
                        textBrush.Dispose()

                        SourceImage.Dispose()
                        DestinationBitmap.Dispose()

                        PictureBox1.Image = proofsheet
                        count = count + 1
                        AtLeastOneImageDrawn = True


                    Else
                        done = True
                    End If

                Next x
            Next y


            'render the proof to a File make sure there are at least on image was drawin on the latest proof sheet before rendering it to a saved image
            If (AtLeastOneImageDrawn = True) Then
                PictureBox1.Image.Save(DestinationPath & "\" & "proof" & ProofSheetCount.ToString & ".PNG", System.Drawing.Imaging.ImageFormat.Png)
                ProofSheetCount = ProofSheetCount + 1
                HTML_SRC_3(DestinationPath)
            End If


        End While

        'clean up and message the user that we are done
        HTML_End(DestinationPath)

        PictureBox1.Dispose()

    End Sub

    'calculate the best scalling factor
    Function CalculateScaleFactor(ByVal bm As Bitmap, ByVal ThumbX As Integer, ByVal ThumbY As Integer) As Single
        'Only scale down
        If bm.Width < ThumbX Then
            Return (1.0)
        End If

        'Scale for landscape image
        If bm.Width > bm.Height Then
            Return (Convert.ToSingle(ThumbX) / Convert.ToSingle(bm.Width))
        End If

        'Assume portrait
        Return (Convert.ToSingle(ThumbY) / Convert.ToSingle(bm.Height))
    End Function

    'delete image selected from the list
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles DELETE_SELECTED_PIC_Button8.Click
        'check if the user really wants to delete the selected index
        Dim result As Integer = MessageBox.Show(ListBox3.SelectedItem.ToString, "Remove Pic from List (does not delete pic)?", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            'unselect the itmes in listbox1 and 3 in preprate to remove the previously selected items
            Dim SelectedIndex As Integer = ListBox3.SelectedIndex
            ListBox3.Items.RemoveAt(SelectedIndex)
            ListBox1.Items.RemoveAt(SelectedIndex)
            ListBox3.SelectedIndex = SelectedIndex - 1
        End If
    End Sub

    'select previous image (up)
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If ListBox3.SelectedIndex > 0 Then
            ListBox3.SelectedIndex = ListBox3.SelectedIndex - 1
            'ListBox1.SelectedIndex = ListBox1.SelectedIndex - 1
            LoadThumbnail()
        End If
    End Sub

    'select next image from pic listbox
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ListBox3.SelectedIndex < ListBox3.Items.Count - 1 Then
            ListBox3.SelectedIndex = ListBox3.SelectedIndex + 1
            'ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
            LoadThumbnail()
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------
    'HTML writing function
    Public Sub HTML_New(ByVal DestinationPath As String)
        'delete the tmp.JPG file if it exists
        Dim DestinationFile As String = DestinationPath & "\index.html"
        If (My.Computer.FileSystem.FileExists(DestinationFile)) Then
            My.Computer.FileSystem.DeleteFile(DestinationFile)
        End If
    End Sub

    'find the oldest YYYYMMDD-MMHHSS named pic in the listbox3
    Public Function oldest_date(ByVal FileListBox As ListBox) As String
        Dim OldestDateName = Path.GetFileNameWithoutExtension(FileListBox.Items(0))

        For Each DateName As String In FileListBox.Items
            If (DateNameToDate(Path.GetFileNameWithoutExtension(DateName)) > DateNameToDate(OldestDateName)) Then
                OldestDateName = Path.GetFileNameWithoutExtension(DateName)
            End If
        Next
        Return (GetDayMonthYear(OldestDateName))
    End Function

    Public Function GetDayMonthYear(ByVal DateName As String) As String
        'YYYYMMDD-HHMMSS -> DD-MM-YYYY
        Dim DateStr As String = Mid(DateName, 5, 2) & "-" & Mid(DateName, 7, 2) & "-" & Mid(DateName, 1, 4)
        Return (DateStr)
    End Function

    'find the youngest YYYYMMDD-MMHHSS named pic in the listbox3
    Public Function newest_date(ByVal FileListBox As ListBox) As String
        Dim NewestDateName = Path.GetFileNameWithoutExtension(FileListBox.Items(0))

        For Each DateName As String In FileListBox.Items
            If (DateNameToDate(Path.GetFileNameWithoutExtension(DateName)) < DateNameToDate(NewestDateName)) Then
                NewestDateName = Path.GetFileNameWithoutExtension(DateName)
            End If
        Next

        Return (GetDayMonthYear(NewestDateName))
    End Function

    'Convert YYYYMMDD-HHMMSS to a vb.net Date format
    Public Function DateNameToDate(ByVal DateName As String) As System.DateTime
        Dim YearStr As String = Mid(DateName, 1, 4)
        If YearStr.CompareTo("YYYY") = 0 Then
            DateName = "19740707-080810"
        End If
        Dim Years As Integer = CInt(Mid(DateName, 1, 4))
        Dim Months As Integer = CInt(Mid(DateName, 5, 2))
        Dim Days As Integer = CInt(Mid(DateName, 7, 2))
        Dim Hours As Integer = CInt(Mid(DateName, 10, 2))
        Dim Minutes As Integer = CInt(Mid(DateName, 12, 2))
        Dim Seconds As Integer = CInt(Mid(DateName, 14, 2))
        Dim d1 As New System.DateTime(Years, Months, Days, Hours, Minutes, Seconds)
        Return (d1)
    End Function

    Public Sub HTML_Start(ByVal DestinationPath As String, ByVal FileListBox As ListBox)
        Dim str As String

        HTML_Write("index.html", DestinationPath, "<!DOCTYPE html>")
        HTML_Write("index.html", DestinationPath, "<!This document autocreated by PIC2PROOF 1.2 (HTML) J. Nolan Dec 4 2019>")
        HTML_Write("index.html", DestinationPath, "<html>")
        HTML_Write("index.html", DestinationPath, "<body>")

        'if in album mode change the title text to the album name
        If Album_CheckBox3.Checked = True Then
            HTML_Write("index.html", DestinationPath, "<center><h1>Photo Album " & TextBox4.Text & "</h1>")
        End If

        str = Microsoft.VisualBasic.Right(DestinationPath, 4)

        'If it is a misc (uncategorized) photo then title the html appropriately
        If String.Compare("misc", str) = 0 Then
            HTML_Write("index.html", DestinationPath, "<center><h1>Photos - MISC" & "</h1>")
        Else
            If EXIF_CheckBox4.Checked = True Then
                HTML_Write("index.html", DestinationPath, "<center><h1>Photos " & newest_date(FileListBox) & " to " & oldest_date(FileListBox) & "</h1>")
            End If
        End If

        'nothing selected for Renaming
        If Album_CheckBox3.Checked = False And EXIF_CheckBox4.Checked = False Then
            HTML_Write("index.html", DestinationPath, "<center><h1>Random Photos</h1>")
        End If
    End Sub

    'HTML write img src with map for each new proof sheet
    Public Sub HTML_SRC_1(ByVal DestinationPath As String, ByVal ProofNumber As Integer)
        HTML_Write("index.html", DestinationPath, "<img src=""proof" & ProofNumber.ToString & ".PNG"" alt=""Proof" & ProofNumber.ToString & """ usemap=""#proof" & ProofNumber.ToString & """>")
        HTML_Write("index.html", DestinationPath, "<map name=""proof" & ProofNumber.ToString & """>")
    End Sub

    Public Sub HTML_SRC_2(ByVal DestinationPath As String, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, FileName As String)
        HTML_Write("index.html", DestinationPath, "<area shape=""rect"" coords=""" & x1.ToString & "," & y1.ToString & "," & x2.ToString & "," & y2.ToString & """ alt=""Pic"" href=""" & FileName & """>")
    End Sub

    Public Sub HTML_SRC_2A(ByVal DestinationPath As String, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, FileName As String, ZoomCount As Integer, width As Integer, height As Integer)

        Dim hs(20) As String

        hs(0) = "<script>"
        hs(1) = "function Zoom" & ZoomCount & "() " & Chr(123) ' start brace
        hs(2) = "    var myWindow = window.open("""", ""Photos"", ""width=640, height=" & Screen.PrimaryScreen.Bounds.Height.ToString & """);"
        hs(3) = "    myWindow.document.write(""<html><head></head><body>"");"
        hs(4) = "    myWindow.document.write(""<script>"");"
        hs(5) = "    myWindow.document.write(""function CloseWindow() " & Chr(123) & """);" 'start brace
        hs(6) = "    myWindow.document.write(""  window.close();"");"
        hs(7) = "    myWindow.document.write(""" & Chr(125) & """);" ' end brace
        hs(8) = "    myWindow.document.write(""<"" + String.fromCharCode(47) + ""script>"");" 'chr(47)
        hs(9) = "    myWindow.document.write(""<center>Picture: " & FileName & " -- CLICK PICTURE TO CLOSE WINDOW<br>"");"

        'write out the img tag with appropriate height or width percent
        If (height > width) Then
            hs(10) = "    myWindow.document.write(""<img src=\""" & FileName & "\"" height=\""100%\"" onclick=\""CloseWindow();\"">"");"
        Else
            hs(10) = "    myWindow.document.write(""<img src=\""" & FileName & "\"" width=\""100%\"" onclick=\""CloseWindow();\"">"");"
        End If

        hs(11) = "    myWindow.document.write(""<br><a href=\""" & FileName & "\"" download>DOWNLOAD PICTURE </a></center>"");"
        hs(12) = "    myWindow.document.write(""</body></html>"");"
        hs(13) = Chr(125) ' end brace
        hs(14) = "</script>"
        hs(15) = "<area shape= ""rect"" coords=""" & x1.ToString & "," & y1.ToString & "," & x2.ToString & "," & y2.ToString & """ alt=""Pic"" onclick=""Zoom" & ZoomCount & "()"" >"

        'write out script strings
        For i As Integer = 0 To 15
            HTML_Write("index.html", DestinationPath, hs(i))
        Next

    End Sub

    Public Sub HTML_SRC_3(ByVal DestinationPath As String)
        HTML_Write("index.html", DestinationPath, "</map")
        HTML_Write("index.html", DestinationPath, "<br>")
    End Sub

    Public Sub HTML_End(ByVal DestinationPath As String)
        HTML_Write("index.html", DestinationPath, "</center>")
        HTML_Write("index.html", DestinationPath, "<h3><Last Updated 2017-24-11</h3>")
        HTML_Write("index.html", DestinationPath, "</map>")
        HTML_Write("index.html", DestinationPath, "</body>")
        HTML_Write("index.html", DestinationPath, "</html>")
    End Sub

    Public Sub HTML_Write(ByVal FileName As String, ByVal DestinationPath As String, str As String)
        Dim FilePath As String = DestinationPath & "\" & FileName

        If Not File.Exists(FilePath) Then
            'create the file
            Using sw As StreamWriter = File.CreateText(FilePath)
                sw.WriteLine(str)
            End Using
        Else
            'append the text to the current index.html
            Using sw As StreamWriter = File.AppendText(FilePath)
                sw.WriteLine(str)
            End Using
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.Text = TextBox1.Text 'source and destination are the same
            TextBox2.Enabled = False
        Else
            TextBox2.Enabled = True
        End If
    End Sub

    'handle checkbox for album mode
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles Album_CheckBox3.CheckedChanged
        If Album_CheckBox3.Checked = True Then
            TextBox4.Enabled = True
            'turn off the exif and YYYYMM options
            EXIF_CheckBox4.Enabled = False
            EXIF_CheckBox4.Checked = False
            YYYYMM_CheckBox5.Enabled = False
            YYYYMM_CheckBox5.Checked = False
        Else
            TextBox4.Enabled = False
            TextBox4.Clear()
            'renable EXIf checkbox
            EXIF_CheckBox4.Enabled = True
        End If
    End Sub

    'refresh button
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        UpdateSourceFileList()
        UpdateDestinationFileList()

        'turn off pic editing until the files are copied
        PICEDITOR_GroupBox2.Enabled = False

        'turn off the MAKEPROOFS button until the items are copied again
        MAKEPROOFS_Button5.Enabled = False

        'turn on the COPY button
        COPYPICS_Button8.Enabled = True
    End Sub

    Private Sub EXIF_CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles EXIF_CheckBox4.CheckedChanged
        If EXIF_CheckBox4.Checked = True Then
            YYYYMM_CheckBox5.Enabled = True

            'turn off the album making if using YYYYMM filtering
            Album_CheckBox3.Enabled = False
        Else
            Album_CheckBox3.Enabled = True
            YYYYMM_CheckBox5.Enabled = False
            YYYYMM_CheckBox5.Checked = False
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then

            ' set the 
            TextBox5.Text = FolderBrowserDialog1.SelectedPath
        End If

        SaveParameters()

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles COPYPICS_Button8.Click
        CopyPics()
        PICEDITOR_GroupBox2.Enabled = True

        'pic the first item listbox1
        ListBox1.SelectedIndex = 0

        'enable the MakeProofs button
        MAKEPROOFS_Button5.Enabled = True
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        If (Not String.IsNullOrEmpty(TextBox1.Text)) Then
            Process.Start(TextBox1.Text)
        End If
    End Sub

    Private Sub Button8_Click_2(sender As Object, e As EventArgs) Handles Button8.Click
        If (Not String.IsNullOrEmpty(TextBox2.Text)) Then
            Process.Start(TextBox2.Text)
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If (Not String.IsNullOrEmpty(TextBox5.Text)) Then
            Process.Start(TextBox5.Text)
        End If
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Dim webAddress As String = "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
        Process.Start(webAddress)
    End Sub

    'archive the destinatino folder into the archive folder
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim startPath As String = TextBox2.Text
        Dim zipPath As String = TextBox5.Text & "\PIC_ARCHIVE_" & DateTime.Now.ToString("yyyy_MM_dd.HH.mm.dd") & ".zip"

        PROCESSING_Label3.Visible = True
        PROCESSING_Label3.Text = "Zipping Archive.."
        Application.DoEvents()

        ZipFile.CreateFromDirectory(startPath, zipPath)
        PROCESSING_Label3.Visible = False
        MsgBox("Done Creating Archvie..")
    End Sub

    Public Function GenerateSHA256String(ByVal inputString As String) As String

        Dim SHA256 As SHA256 = SHA256Managed.Create()
        Dim bytes() As Byte = Encoding.UTF8.GetBytes(inputString)
        Dim hash() As Byte = SHA256.ComputeHash(bytes)
        Return GetStringFromHash(hash)
    End Function

    Private Function GetStringFromHash(ByVal hash() As Byte) As String

        Dim result As StringBuilder = New StringBuilder()
        For i As Integer = 0 To hash.Length - 1
            result.Append(hash(i).ToString("x2"))
        Next
        Return result.ToString()
    End Function

    'build the collectector index.html file

    Private Function IsYYYYMM(ByVal PathFileName As String) As Boolean
        Dim FilePath As String = Path.GetDirectoryName(PathFileName)
        Dim Directory As String = Mid(FilePath, FilePath.Length - 6, 7)
        Dim i As Integer

        For i = 0 To 3
            If Not Char.IsNumber(Directory(i)) Then
                Return False
            End If
        Next

        For i = 5 To 6
            If Not Char.IsNumber(Directory(i)) Then
                Return False
            End If
        Next

        'no non numbers were found so must all be numbers
        Return True
    End Function

    Private Sub BuildCollector()
        Dim MonthStr() As String = {"nope", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "nope"}
        Dim DestinationPath As String = TextBox2.Text + "\pics"
        Dim year_found As Integer = 0
        Dim out_str As String = ""
        Dim RelativePath As String
        Dim FilePath As String
        Dim CollectorFileName As String = GenerateSHA256String(WEBSITEPASSWROD_TextBox6.Text) + ".html"

        'get all the directories where there is an index.html to a proof set
        Dim txtFilesArray As String() = Directory.GetFiles(DestinationPath, "index.html", SearchOption.AllDirectories)
        Dim i As Integer

        'write webstie index.html if needed
        WriteIndexHtML(TextBox2.Text)

        'write preamble
        Collector_HTML_Start(CollectorFileName, DestinationPath)

        For year As Integer = 1974 To 3001 'lol y3k kiss my ass (hey robot overlord what is 3 divided 0)
            For month As Integer = 1 To 12
                For i = 0 To txtFilesArray.Length - 1
                    If IsYYYYMM(txtFilesArray(i)) Then
                        If YYYYMM_Found(txtFilesArray(i), month, year) Then
                            If year_found = 0 Then
                                HTML_Write(CollectorFileName, DestinationPath, "<br><h2>" & year.ToString & "</h2><br><h3>")
                                year_found = year
                            End If

                            'create a relative path
                            FilePath = Path.GetDirectoryName(txtFilesArray(i))
                            RelativePath = Mid(FilePath, FilePath.Length - 6, 7) + "\" + Path.GetFileName(txtFilesArray(i))

                            out_str = out_str + "  " + "<a href=" + Chr(34) + RelativePath + Chr(34) + ">" + MonthStr(month) + "</a>"
                        End If
                    End If
                Next i
            Next month
            year_found = 0
            If out_str <> "" Then
                out_str = out_str + "</h3>"

                HTML_Write(CollectorFileName, DestinationPath, out_str)
                out_str = ""
            End If
        Next year


        'Collect all the album's and misc
        out_str = ""
        HTML_Write(CollectorFileName, DestinationPath, "<br><h2>MISC</h2><br>")

        For i = 0 To txtFilesArray.Length - 1
            If Not IsYYYYMM(txtFilesArray(i)) Then
                'create a relative path
                RelativePath = CreateRelativePath(txtFilesArray(i))
                out_str = out_str + "    " + "<a href=" + Chr(34) + RelativePath + Chr(34) + ">" + RelativePath + "</a>"
            End If

            If out_str <> "" Then
                out_str = out_str + "<br>"

                HTML_Write(CollectorFileName, DestinationPath, out_str)
                out_str = ""
            End If
        Next i

        'close out the collector html file

        Collector_HTML_End(CollectorFileName, DestinationPath)

        MsgBox("Collector written")
    End Sub

    Private Function CreateRelativePath(ByVal PathFileName As String) As String
        Dim PathName As String = Path.GetDirectoryName(PathFileName)
        Dim FileName As String = Path.GetFileName(PathFileName)
        For i = 1 To PathName.Length - 1
            If Mid(PathName, i, 5) = "pics\" Then
                Dim str As String = Mid(PathName, i + 5, PathName.Length - i - 4) & "\" & FileName
                Return (str)
            End If
        Next

        Return (PathFileName)
    End Function

    Private Function YYYYMM_Found(ByVal file_str As String, ByVal month As Integer, ByVal year As String) As Boolean
        Dim FilePath As String = Path.GetDirectoryName(file_str)
        Dim year_test = CInt(Mid(FilePath, FilePath.Length - 6, 4))
        Dim month_test = CInt(Mid(FilePath, FilePath.Length - 1, 2))

        If year_test = year And month_test = month Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub Collector_HTML_Start(ByVal FileName As String, ByVal DestinationPath As String)
        HTML_Write(FileName, DestinationPath, "<!DOCTYPE html>")
        HTML_Write(FileName, DestinationPath, "<!This document autocreated by PIC2PROOF 1.2 (HTML) J. Nolan Dec 4 2019>")
        HTML_Write(FileName, DestinationPath, "<html>")
        HTML_Write(FileName, DestinationPath, "<body>")

        HTML_Write(FileName, DestinationPath, "<center><h1>Photos</h1>")
    End Sub

    Public Sub Collector_HTML_End(ByVal Filename As String, ByVal DestinationPath As String)
        HTML_Write(Filename, DestinationPath, "<br></center>")
        HTML_Write(Filename, DestinationPath, "<h3><Last Updated 2016-12-06</h3>")
        HTML_Write(Filename, DestinationPath, "</map>")
        HTML_Write(Filename, DestinationPath, "</body>")
        HTML_Write(Filename, DestinationPath, "</html>")
    End Sub

    'collect all html files into one
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If WEBSITEPASSWROD_TextBox6.Text = Nothing Then
            MsgBox("Set Website PassWord first")
        Else
            BuildCollector()
        End If
    End Sub

    'index.html file for the website has javascript sha256 code.  I know right!
    Public Sub WriteIndexHtML(ByVal DestinationPath As String)

        Dim file(200) As String
        file(0) = "<html><body><center>"
        file(1) = "<input id='hacker-eat-me' type='text'  />"
        file(2) = "<a href=""collector.html"" onclick=""javascript:return validatePass()"">Submit</a></center>"
        file(3) = "<script>"
        file(4) = "function validatePass(){"
        file(5) = "   var pw = SHA256(document.getElementById('hacker-eat-me').value);"
        file(6) = "   window.open(""pics/"" + pw + "".html"");"
        file(7) = "}"

        file(8) = "/** Secure Hash Algorithm (SHA256) http//www.webtoolkit.info/ Original code by Angel Marin, Paul Johnston. */"

        file(9) = "function SHA256(s){"
        file(10) = "  var chrsz = 8;"
        file(11) = "  var hexcase = 0;"
        file(12) = "  function safe_add(x, y) {"
        file(13) = "     var lsw = (x & 0xFFFF) + (y & 0xFFFF);"
        file(14) = "     var msw = (x >> 16) + (y >> 16) + (lsw >> 16);"
        file(15) = "     return (msw << 16) | (lsw & 0xFFFF);"
        file(16) = "  }"

        file(17) = "   function S(X, n) { return ( X >>> n ) | (X << (32 - n)); }"
        file(18) = "   function R(X, n) { return ( X >>> n ); }"
        file(19) = "   function Ch(x, y, z) { return ((x & y) ^ ((~x) & z)); }"
        file(20) = "   function Maj(x, y, z) { return ((x & y) ^ (x & z) ^ (y & z)); }"
        file(21) = "   function Sigma0256(x) { return (S(x, 2) ^ S(x, 13) ^ S(x, 22)); }"
        file(22) = "   function Sigma1256(x) { return (S(x, 6) ^ S(x, 11) ^ S(x, 25)); }"
        file(23) = "   function Gamma0256(x) { return (S(x, 7) ^ S(x, 18) ^ R(x, 3)); }"
        file(24) = "   function Gamma1256(x) { return (S(x, 17) ^ S(x, 19) ^ R(x, 10)); }"

        file(25) = "   function core_sha256(m, l) {"
        file(26) = "      var K = new Array(0x428A2F98, 0x71374491, 0xB5C0FBCF, 0xE9B5DBA5, 0x3956C25B, 0x59F111F1, 0x923F82A4, 0xAB1C5ED5, 0xD807AA98, 0x12835B01, 0x243185BE, 0x550C7DC3, 0x72BE5D74, 0x80DEB1FE, 0x9BDC06A7, 0xC19BF174, 0xE49B69C1, 0xEFBE4786, 0xFC19DC6, 0x240CA1CC, 0x2DE92C6F, 0x4A7484AA, 0x5CB0A9DC, 0x76F988DA, 0x983E5152, 0xA831C66D, 0xB00327C8, 0xBF597FC7, 0xC6E00BF3, 0xD5A79147, 0x6CA6351, 0x14292967, 0x27B70A85, 0x2E1B2138, 0x4D2C6DFC, 0x53380D13, 0x650A7354, 0x766A0ABB, 0x81C2C92E, 0x92722C85, 0xA2BFE8A1, 0xA81A664B, 0xC24B8B70, 0xC76C51A3, 0xD192E819, 0xD6990624, 0xF40E3585, 0x106AA070, 0x19A4C116, 0x1E376C08, 0x2748774C, 0x34B0BCB5, 0x391C0CB3, 0x4ED8AA4A, 0x5B9CCA4F, 0x682E6FF3, 0x748F82EE, 0x78A5636F, 0x84C87814, 0x8CC70208, 0x90BEFFFA, 0xA4506CEB, 0xBEF9A3F7, 0xC67178F2);"
        file(27) = "      var HASH = new Array(0x6A09E667, 0xBB67AE85, 0x3C6EF372, 0xA54FF53A, 0x510E527F, 0x9B05688C, 0x1F83D9AB, 0x5BE0CD19);"
        file(28) = "      var W = new Array(64);"
        file(29) = "      var a, b, c, d, e, f, g, h, i, j;"
        file(30) = "      var T1, T2;"
        file(31) = "      m[l >> 5] |= 0x80 << (24 - l % 32);"
        file(32) = "      m[((l + 64 >> 9) << 4) + 15] = l;"
        file(33) = "      for (var i = 0; i<m.length; i+=16 ) {"
        file(34) = "         a = HASH[0];"
        file(35) = "         b = HASH[1];"
        file(36) = "         c = HASH[2];"
        file(37) = "         d = HASH[3];"
        file(38) = "         e = HASH[4];"
        file(39) = "         f = HASH[5];"
        file(40) = "         g = HASH[6];"
        file(41) = "         h = HASH[7];"
        file(42) = "         for (var j = 0; j<64; j++) {"
        file(43) = "             if (j < 16) W[j] = m[j + i];"
        file(44) = "                else W[j] = safe_add(safe_add(safe_add(Gamma1256(W[j - 2]), W[j - 7]), Gamma0256(W[j - 15])), W[j - 16]);"
        file(45) = "             T1 = safe_add(safe_add(safe_add(safe_add(h, Sigma1256(e)), Ch(e, f, g)), K[j]), W[j]);"
        file(46) = "             T2 = safe_add(Sigma0256(a), Maj(a, b, c));"
        file(47) = "             h = g;"
        file(48) = "             g = f;"
        file(49) = "             f = e;"
        file(50) = "             e = safe_add(d, T1);"
        file(51) = "             d = c;"
        file(52) = "             c = b;"
        file(53) = "             b = a;"
        file(54) = "             a = safe_add(T1, T2);"
        file(55) = "         }"
        file(56) = "         HASH[0] = safe_add(a, HASH[0]);"
        file(57) = "         HASH[1] = safe_add(b, HASH[1]);"
        file(58) = "         HASH[2] = safe_add(c, HASH[2]);"
        file(59) = "         HASH[3] = safe_add(d, HASH[3]);"
        file(60) = "         HASH[4] = safe_add(e, HASH[4]);"
        file(61) = "         HASH[5] = safe_add(f, HASH[5]);"
        file(62) = "         HASH[6] = safe_add(g, HASH[6]);"
        file(63) = "         HASH[7] = safe_add(h, HASH[7]);"
        file(64) = "      }"
        file(65) = "      return HASH;"
        file(66) = "    }"

        file(67) = "   function str2binb(str) {"
        file(68) = "      var bin = Array();"
        file(69) = "      var mask = (1 << chrsz) - 1;"
        file(70) = "      for (var i = 0; i < str.length * chrsz; i += chrsz) {"
        file(71) = "         bin[i>>5] |= (str.charCodeAt(i / chrsz) & mask) << (24 - i%32);"
        file(72) = "      }"
        file(73) = "      return bin;"
        file(74) = "   }"

        file(75) = "   function Utf8Encode(string) {"
        file(76) = "      string = string.replace(/\r\n/g,""\n"");"
        file(77) = "      var utftext = """";"
        file(78) = "        for (var n = 0; n < string.length; n++) {"
        file(79) = "          var c = string.charCodeAt(n);"
        file(80) = "          if (c < 128) {"
        file(81) = "            utftext += String.fromCharCode(c);"
        file(82) = "        }"
        file(83) = "        else if ((c > 127) && (c < 2048)) {"
        file(84) = "        utftext += String.fromCharCode((c >> 6) | 192);"
        file(85) = "        utftext += String.fromCharCode((c & 63) | 128);"
        file(86) = "   }"
        file(87) = "   else {"
        file(88) = "     utftext += String.fromCharCode((c >> 12) | 224);"
        file(89) = "     utftext += String.fromCharCode(((c >> 6) & 63) | 128);"
        file(90) = "     utftext += String.fromCharCode((c & 63) | 128);"
        file(91) = "   }"
        file(92) = " }"
        file(93) = " return utftext;"
        file(94) = "}"

        file(95) = "function binb2hex(binarray) {"
        file(96) = "   var hex_tab = hexcase ? ""0123456789ABCDEF"" : ""0123456789abcdef"";"
        file(97) = "   var str = """";"
        file(98) = "   for (var i = 0; i < binarray.length * 4; i++) {"
        file(99) = "     str += hex_tab.charAt((binarray[i>>2] >> ((3 - i%4) * 8 + 4)) & 0xF) +"
        file(100) = "           hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8  )) & 0xF);"
        file(101) = "  }"
        file(102) = "  return str;"
        file(103) = "}"

        file(104) = " s = Utf8Encode(s);"
        file(105) = " return binb2hex(core_sha256(str2binb(s), s.length * chrsz));"
        file(106) = "}"
        file(107) = "</script></body></html>"

        'if the main website index.html does not exist then write it to the destination folder
        If Not My.Computer.FileSystem.FileExists(DestinationPath & "\index.hmtl") Then
            For i As Integer = 0 To 107
                HTML_Write("index.html", DestinationPath, file(i))
            Next
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            'turn on the make proofs buttton
            MAKEPROOFS_Button5.Enabled = True

            Dim result As Integer = MessageBox.Show("Are you sure you want to process proofs only", "caption", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                CheckBox2.Checked = False
                MAKEPROOFS_Button5.Enabled = False
            ElseIf result = DialogResult.Yes Then
                LoadDestinationFiles()
            End If
        End If

    End Sub

    Private Sub loaddestinationfiles()
        'get all the directories where there is an index.html to a proof set
        Dim txtFilesArray As String() = Directory.GetFiles(TextBox2.Text, "*.jpg", SearchOption.AllDirectories)
        For Each Str As String In txtFilesArray
            ListBox3.Items.Add(Str)
        Next
    End Sub
End Class
