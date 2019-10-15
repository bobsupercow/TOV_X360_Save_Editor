Imports System.Windows.Forms
Imports X360.IO
Imports X360.STFS

Public Class MainForm
    Inherits Form

    Private pack As STFSPackage
    Private entry As FileEntry
    Private savedata As DJsIO

    Public Sub New()
        InitializeComponent()
        X360.XAbout.WriteLegalLocally()
    End Sub

    ' Use the xSave to write data like xSave.Write(400)
    ' NOTE: The offsets are NOT in terms of the package, say a file in the package
    ' was located at 0xD000, you wanted to write a byte at 0xD032.  xSave is the extracted
    ' data so you would write xSave.Position = &H32, NOT xSave.Position = &HD032

    ''' <summary>
    ''' Edits an attribute for a given character, writing the changes to the file.
    ''' If the user doesn't max out the values, they will have to equip a skill 
    ''' or accessory for their modified value to be updated.
    ''' </summary>
    ''' <param name="characterBaseAddress">The base memory address where a character's data begins.</param>
    ''' <param name="attribute">The attribute to be edited.</param>
    ''' <param name="value">The value.</param>
    Private Sub editCharacterAttribute(ByVal characterBaseAddress As CharacterBaseOffsets, _
                                       ByVal attribute As CharacterValues, _
                                       ByVal value As Integer)

        Dim valueToSet As Integer
        'Make sure we don't set the value too high.
        Select Case attribute
            Case CharacterValues.HP
                If value > MAX_HP Then
                    valueToSet = MAX_HP
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress
                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.HPCurrent
                savedata.Write(valueToSet)
                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.HPModified
                savedata.Write(valueToSet)
                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.HPBase
                savedata.Write(valueToSet)
            Case CharacterValues.TP
                If value > MAX_TP Then
                    valueToSet = MAX_TP
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.TPCurrent
                savedata.Write(valueToSet)
                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.TPModified
                savedata.Write(valueToSet)
                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.TPBase
                savedata.Write(valueToSet)
            Case CharacterValues.EXP
                If value > MAX_EXP Then
                    valueToSet = MAX_EXP
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.EXP
                savedata.Write(valueToSet)
            Case CharacterValues.Agility
                If value > MAX_AGILITY Then
                    valueToSet = MAX_AGILITY
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.Agility
                savedata.Write(valueToSet)
            Case CharacterValues.PhysicalAttack
                If value > MAX_PHYSICAL_ATTACK Then
                    valueToSet = MAX_PHYSICAL_ATTACK
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.PhysicalAttack
                savedata.Write(valueToSet)
            Case CharacterValues.PhysicalDefense
                If value > MAX_PHYSICAL_DEFENSE Then
                    valueToSet = MAX_PHYSICAL_DEFENSE
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.PhysicalDefense
                savedata.Write(valueToSet)
            Case CharacterValues.MagicAttack
                If value > MAX_MAGIC_ATTACK Then
                    valueToSet = MAX_MAGIC_ATTACK
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.MagicAttack
                savedata.Write(valueToSet)
            Case CharacterValues.MagicDefense
                If value > MAX_MAGIC_DEFENSE Then
                    valueToSet = MAX_MAGIC_DEFENSE
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.MagicDefense
                savedata.Write(valueToSet)
            Case CharacterValues.Luck
                If value > MAX_LUCK Then
                    valueToSet = MAX_LUCK
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.Luck
                savedata.Write(valueToSet)
            Case CharacterValues.Level
                If value > MAX_LEVEL Then
                    valueToSet = MAX_LEVEL
                Else
                    valueToSet = value
                End If

                savedata.Position = characterBaseAddress + Statics.CharacterValueOffsets.Level
                savedata.Write(valueToSet)

        End Select
    End Sub


#Region "Edit Other"

    ''' <summary>
    ''' Edits the Gald. Writes the changes to the save file.
    ''' </summary>
    ''' <param name="value">The value.</param>
    Private Sub editGald(ByVal value As Integer)
        Dim valueToSet As Integer
        'Make sure we don't set the value too high.
        If value > MAX_GALD Then
            valueToSet = MAX_GALD
        Else
            valueToSet = value
        End If

        savedata.Position = OtherMemoryLocations.Gald
        savedata.Write(value)

    End Sub

#End Region

#Region "Edit Item Counts"

#End Region

#Region "Open Methods"

    ''' <summary>
    ''' Opens a CON package for editing. Extracts the saveData and sets appropriate class-level variables.
    ''' </summary>
    ''' <param name="fileName">The full path of the CON file.</param>
    Private Sub openPackage(ByVal fileName As String)

        'Close remaining open IO streams in case there is already an open file.
        If pack IsNot Nothing Then
            pack.CloseIO()
        End If
        If savedata IsNot Nothing Then
            savedata.Close()
        End If

        'Open the CON package.
        Dim xPackage As New STFSPackage(fileName, Nothing)
        Try
            If Not xPackage.ParseSuccess Then
                MessageBox.Show(My.Resources.msgBoxErrorOpenParsingFailure, My.Resources.msgBoxTitleErrorOpen)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.msgBoxErrorOpenParsingFailure, My.Resources.msgBoxTitleErrorOpen)
        End Try

        'Get the SaveData from the container
        Dim xent As FileEntry = xPackage.GetFile("save.dat")
        Try
            If xent Is Nothing Then
                'If this occurs then the Save Data in the CON file couldn't be located.
                xPackage.CloseIO()
                MessageBox.Show(My.Resources.msgBoxErrorOpenSaveMissing, My.Resources.msgBoxTitleErrorOpen)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.msgBoxErrorOpenSaveMissing, My.Resources.msgBoxTitleErrorOpen)
        End Try

        'Extract the SaveData from the container so we can work on it without keeping the CON open.
        Dim outlocale As String = X360.Other.VariousFunctions.GetTempFileLocale()
        Try
            If Not xent.Extract(outlocale) Then
                'If this occurs then the Save Data failed to extract properly. 
                xPackage.CloseIO()
                MessageBox.Show(My.Resources.msgBoxErrorOpenSaveExtractFailure, My.Resources.msgBoxTitleErrorOpen)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.msgBoxErrorOpenSaveExtractFailure, My.Resources.msgBoxTitleErrorOpen)
        End Try



        'Set class level variables.
        savedata = New DJsIO(outlocale, DJFileMode.Open, True)
        entry = xent
        pack = xPackage


        'Populate the values on the main form any time a file is opened.
        populateValues()

        MessageBox.Show(My.Resources.msgBoxOpenSuccess)
    End Sub

    ''' <summary>
    ''' Populates the values on the form based on savedata
    ''' </summary>
    Private Sub populateValues()
        Try
            If savedata IsNot Nothing Then
                'Populate Yuri's Information
                savedata.Position = Statics.CharacterBaseOffsets.Yuri + Statics.CharacterValueOffsets.HPBase
                Me.uxYuriHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Yuri + Statics.CharacterValueOffsets.TPBase
                Me.uxYuriTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Yuri + Statics.CharacterValueOffsets.EXP
                Me.uxYuriEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Repede's Information
                savedata.Position = Statics.CharacterBaseOffsets.Repede + Statics.CharacterValueOffsets.HPBase
                Me.uxRepedeHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Repede + Statics.CharacterValueOffsets.TPBase
                Me.uxRepedeTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Repede + Statics.CharacterValueOffsets.EXP
                Me.uxRepedeEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Estelle's Information
                savedata.Position = Statics.CharacterBaseOffsets.Estelle + Statics.CharacterValueOffsets.HPBase
                Me.uxEstelleHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Estelle + Statics.CharacterValueOffsets.TPBase
                Me.uxEstelleTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Estelle + Statics.CharacterValueOffsets.EXP
                Me.uxEstelleEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Karol's Information
                savedata.Position = Statics.CharacterBaseOffsets.Karol + Statics.CharacterValueOffsets.HPBase
                Me.uxKarolHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Karol + Statics.CharacterValueOffsets.TPBase
                Me.uxKarolTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Karol + Statics.CharacterValueOffsets.EXP
                Me.uxKarolEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Rita's Information
                savedata.Position = Statics.CharacterBaseOffsets.Rita + Statics.CharacterValueOffsets.HPBase
                Me.uxRitaHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Rita + Statics.CharacterValueOffsets.TPBase
                Me.uxRitaTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Rita + Statics.CharacterValueOffsets.EXP
                Me.uxRitaEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Raven's Information
                savedata.Position = Statics.CharacterBaseOffsets.Raven + Statics.CharacterValueOffsets.HPBase
                Me.uxRavenHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Raven + Statics.CharacterValueOffsets.TPBase
                Me.uxRavenTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Raven + Statics.CharacterValueOffsets.EXP
                Me.uxRavenEXPTextBox.Text = savedata.ReadInt32.ToString
                'Populate Judith's Information
                savedata.Position = Statics.CharacterBaseOffsets.Judith + Statics.CharacterValueOffsets.HPBase
                Me.uxJudithHPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Judith + Statics.CharacterValueOffsets.TPBase
                Me.uxJudithTPTextBox.Text = savedata.ReadInt32.ToString
                savedata.Position = Statics.CharacterBaseOffsets.Judith + Statics.CharacterValueOffsets.EXP
                Me.uxJudithEXPTextBox.Text = savedata.ReadInt32.ToString

                'Populate Other Information
                savedata.Position = OtherMemoryLocations.Gald
                Me.uxOtherGaldTextBox.Text = savedata.ReadInt32.ToString

            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.msgBoxErrorPopulateValuesFailed)
        End Try
    End Sub

    Private Sub uxOpenButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxOpenButton.Click
        Dim ofd As New OpenFileDialog
        If ofd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            openPackage(ofd.FileName)
        End If
    End Sub

    Private Sub MainForm_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Open the first file
            openPackage(MyFiles(0))
        End If
    End Sub

    Private Sub MainForm_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

#End Region

#Region "Save Methods"

    Private Sub savePackage()
        Try
            preparePackage()
            If entry IsNot Nothing Then
                'Change savedata to save name or path in the CON FILE!
                If Not entry.Replace(savedata) Then
                    'Error due to not being able to replace save in CON FILE!
                    MessageBox.Show(My.Resources.msgBoxErrorSaveCONInjectionFailure, My.Resources.msgBoxTitleErrorSave)
                    Exit Sub
                End If

                If Not pack.FlushPackage(New RSAParams(Application.StartupPath & "/KV.bin")) Then
                    'Error due most likely b/c of lack of KV.bin in editor folder!
                    MessageBox.Show(My.Resources.msgBoxErrorSaveReHashFailure)
                    Exit Sub
                End If

                MessageBox.Show(My.Resources.msgBoxSaveSuccess)
                populateValues()
                Exit Sub
            End If

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub uxSaveButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxSaveButton.Click
        savePackage()
    End Sub

    ''' <summary>
    ''' Prepares the package for saving by reading in data from the forms.
    ''' </summary>
    Private Sub preparePackage()
        Try
            If savedata IsNot Nothing Then
                'Get HP Values
                If uxMaxAllHPCheckBox.Checked Then
                    Dim characters As Integer() = System.Enum.GetValues(GetType(CharacterBaseOffsets))
                    For i = 0 To characters.Length - 1
                        editCharacterAttribute(characters(i), CharacterValues.HP, MAX_HP)
                    Next
                Else
                    editCharacterAttribute(CharacterBaseOffsets.Yuri, CharacterValues.HP, CInt(Me.uxYuriHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Repede, CharacterValues.HP, CInt(Me.uxRepedeHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Estelle, CharacterValues.HP, CInt(Me.uxEstelleHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Karol, CharacterValues.HP, CInt(Me.uxKarolHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Rita, CharacterValues.HP, CInt(Me.uxRitaHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Raven, CharacterValues.HP, CInt(Me.uxRavenHPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Judith, CharacterValues.HP, CInt(Me.uxJudithHPTextBox.Text))
                End If

                'Get TP Values
                If uxMaxAllTPCheckBox.Checked Then
                    Dim characters As Integer() = System.Enum.GetValues(GetType(CharacterBaseOffsets))
                    For i = 0 To characters.Length - 1
                        editCharacterAttribute(characters(i), CharacterValues.TP, MAX_TP)
                    Next
                Else
                    editCharacterAttribute(CharacterBaseOffsets.Yuri, CharacterValues.TP, CInt(Me.uxYuriTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Repede, CharacterValues.TP, CInt(Me.uxRepedeTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Estelle, CharacterValues.TP, CInt(Me.uxEstelleTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Karol, CharacterValues.TP, CInt(Me.uxKarolTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Rita, CharacterValues.TP, CInt(Me.uxRitaTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Raven, CharacterValues.TP, CInt(Me.uxRavenTPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Judith, CharacterValues.TP, CInt(Me.uxJudithTPTextBox.Text))
                End If

                'Get EXP Values
                If uxMaxAllEXPCheckBox.Checked Then
                    Dim characters As Integer() = System.Enum.GetValues(GetType(CharacterBaseOffsets))
                    For i = 0 To characters.Length - 1
                        editCharacterAttribute(characters(i), CharacterValues.TP, MAX_EXP)
                    Next
                Else
                    editCharacterAttribute(CharacterBaseOffsets.Yuri, CharacterValues.EXP, CInt(Me.uxYuriEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Repede, CharacterValues.EXP, CInt(Me.uxRepedeEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Estelle, CharacterValues.EXP, CInt(Me.uxEstelleEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Karol, CharacterValues.EXP, CInt(Me.uxKarolEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Rita, CharacterValues.EXP, CInt(Me.uxRitaEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Raven, CharacterValues.EXP, CInt(Me.uxRavenEXPTextBox.Text))
                    editCharacterAttribute(CharacterBaseOffsets.Judith, CharacterValues.EXP, CInt(Me.uxJudithEXPTextBox.Text))
                End If

                'Get Gald Values
                If uxMaxGaldCheckBox.Checked Then
                    savedata.Position = OtherMemoryLocations.Gald
                    savedata.Write(MAX_GALD)
                Else
                    savedata.Position = OtherMemoryLocations.Gald
                    savedata.Write(CInt(Me.uxOtherGaldTextBox.Text))
                End If
                If uxMaxAllStatsCheckBox.Checked Then
                    Dim characters As Integer() = System.Enum.GetValues(GetType(CharacterBaseOffsets))
                    For i = 0 To characters.Length - 1
                        editCharacterAttribute(characters(i), CharacterValues.PhysicalAttack, MAX_PHYSICAL_ATTACK)
                        editCharacterAttribute(characters(i), CharacterValues.PhysicalDefense, MAX_PHYSICAL_DEFENSE)
                        editCharacterAttribute(characters(i), CharacterValues.MagicAttack, MAX_MAGIC_ATTACK)
                        editCharacterAttribute(characters(i), CharacterValues.MagicDefense, MAX_MAGIC_DEFENSE)
                        editCharacterAttribute(characters(i), CharacterValues.Agility, MAX_AGILITY)
                        editCharacterAttribute(characters(i), CharacterValues.Luck, MAX_LUCK)
                    Next
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.msgBoxErrorPopulateValuesFailed)
        End Try
    End Sub

#End Region

#Region "TextBox_TextChanged Handlers"

#Region "Yuri"

    Private Sub uxYuriHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxYuriHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxYuriTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxYuriTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxYuriEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxYuriEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Repede"

    Private Sub uxRepedeHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRepedeHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRepedeTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRepedeTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRepedeEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRepedeEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Estelle"

    Private Sub uxEstelleHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxEstelleHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxEstelleTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxEstelleTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxEstelleEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxEstelleEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Karol"

    Private Sub uxKarolHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxKarolHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxKarolTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxKarolTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxKarolEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxKarolEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Rita"

    Private Sub uxRitaHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRitaHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRitaTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRitaTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRitaEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRitaEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Raven"

    Private Sub uxRavenHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRavenHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRavenTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRavenTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxRavenEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxRavenEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#Region "Judith"

    Private Sub uxJudithHPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxJudithHPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxJudithTPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxJudithTPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

    Private Sub uxJudithEXPTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles uxJudithEXPTextBox.TextChanged
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If isInputValid(textBox.Text) = False Then
            textBox.Text = ""
        End If
    End Sub

#End Region

#End Region

#Region "TextBox_KeyPress Handlers"

#Region "Yuri"

    Private Sub uxYuriHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxYuriHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxYuriTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxYuriTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxYuriEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxYuriEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub


#End Region

#Region "Repede"

    Private Sub uxRepedeHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRepedeHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRepedeTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRepedeTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRepedeEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRepedeEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

#Region "Estelle"

    Private Sub uxEstelleHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxEstelleHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxEstelleTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxEstelleTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxEstelleEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxEstelleEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

#Region "Karol"

    Private Sub uxKarolHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxKarolHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxKarolTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxKarolTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxKarolEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxKarolEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

#Region "Rita"

    Private Sub uxRitaHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRitaHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRitaTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRitaTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRitaEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRitaEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

#Region "Raven"

    Private Sub uxRavenHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRavenHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRavenTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRavenTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxRavenEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxRavenEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

#Region "Judith"

    Private Sub uxJudithHPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxJudithHPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxJudithTPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxJudithTPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

    Private Sub uxJudithEXPTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles uxJudithEXPTextBox.KeyPress
        e.Handled = blockInput(e.KeyChar)
    End Sub

#End Region

    ''' <summary>
    ''' Blocks any input that isn't valid.
    ''' </summary>
    ''' <param name="value">The value.</param><returns></returns>
    Private Function blockInput(ByVal value As Char) As Boolean
        Dim returnValue As Boolean = True
        'By default, don't allow any input.
        returnValue = True
        'If the input was a control character. e.g. "Backspace", allow it.
        If Char.IsControl(value) Then returnValue = False
        'If the input was numeric, allow it.
        If IsNumeric(value) Then returnValue = False
        Return returnValue
    End Function

    ''' <summary>
    ''' Determines whether [is input valid] [the specified value].
    ''' </summary>
    ''' <param name="value">The value.</param><returns>
    '''   <c>true</c> if [is input valid] [the specified value]; otherwise, <c>false</c>.
    ''' </returns>
    Private Function isInputValid(ByVal value As String) As Boolean
        For Each c In value.ToCharArray
            If blockInput(c) Then
                Return False
            End If
        Next
        Return True
    End Function

#End Region

End Class
