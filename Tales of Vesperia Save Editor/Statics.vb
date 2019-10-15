Module Statics

#Region "Maximum Values"
    ''' <summary>
    ''' The maximum value of a character's HP.
    ''' </summary>
    Public Const MAX_HP As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's TP.
    ''' </summary>
    Public Const MAX_TP As Integer = 999
    ''' <summary>
    ''' The maximum value of a character's EXP
    ''' </summary>
    Public Const MAX_EXP As Integer = 99999999
    ''' <summary>
    ''' The maximum value of the party's GALD count
    ''' </summary>
    Public Const MAX_GALD As Integer = 99999999
    ''' <summary>
    ''' The maximum value of a character's physical attack attribute
    ''' </summary>
    Public Const MAX_PHYSICAL_ATTACK As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's physical defense attribute
    ''' </summary>
    Public Const MAX_PHYSICAL_DEFENSE As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's magic attack attribute
    ''' </summary>
    Public Const MAX_MAGIC_ATTACK As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's magic attack attribute
    ''' </summary>
    Public Const MAX_MAGIC_DEFENSE As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's magic attack attribute
    ''' </summary>
    Public Const MAX_AGILITY As Integer = 9999
    ''' <summary>
    ''' The maximum value of a character's magic attack attribute
    ''' </summary>
    Public Const MAX_LUCK As Integer = 100
    ''' <summary>
    ''' The maximum value of a character's level attribute
    ''' </summary>
    Public Const MAX_LEVEL As Integer = 9999
#End Region


    ''' <summary>
    ''' The base memory address offset where each character's data begins in save.dat
    ''' </summary>
    ''' <remarks>The base memory address is also the address of the character's level.</remarks>
    Public Enum CharacterBaseOffsets
        Estelle = &HCC10
        Karol = &H10B50
        Raven = &H189D0
        Repede = &H20850
        Rita = &H14A90
        Judith = &H1C910
        Yuri = &H8CD0
    End Enum

    ''' <summary>
    ''' The memory offset of various character traits from the CharacterBaseOffset address.
    ''' </summary>
    Public Enum CharacterValueOffsets As Integer
        Level = &H0
        HPCurrent = &H4
        TPCurrent = &H8
        HPModified = &HC
        TPModified = &H10
        Unknown1 = &H14
        EXP = &H18
        HPBase = &H1C
        TPBase = &H20
        PhysicalAttack = &H24
        MagicAttack = &H28
        PhysicalDefense = &H2C
        MagicDefense = &H30
        Unknown2 = &H34
        Agility = &H38
        Luck = &H3C
    End Enum

    ''' <summary>
    ''' Each type of attribute which can be changed for a given character.
    ''' </summary>
    Public Enum CharacterValues
        Level = 0
        HP = 1
        TP = 2
        EXP = 3
        PhysicalAttack = 4
        PhysicalDefense = 5
        MagicAttack = 6
        MagicDefense = 7
        Agility = 8
        Luck = 9
        Unknown1 = 10
        Unknown2 = 11
    End Enum

    ''' <summary>
    ''' Memory Locations not listed elsewhere
    ''' </summary>
    Public Enum OtherMemoryLocations
        Gald = &H42D8
    End Enum







    ''' <summary>
    ''' This method of accessing the addresses is obselete as of version 1.0
    ''' Offset address of various character traits.
    ''' </summary>
    Public Enum LegacyMemoryLocations
        'HP
        EstelleHPLocCurrent = &HCC14
        EstelleHPLocModified = &HCC1C
        EstelleHPLocBase = &HCC2C
        JudithHPLocCurrent = &H1C914
        JudithHPLocModified = &H1C91C
        JudithHPLocBase = &H1C92C
        KarolHPLocCurrent = &H10B54
        KarolHPLocModified = &H10B5C
        KarolHPLocBase = &H10B6C
        RavenHPLocCurrent = &H189D4
        RavenHPLocModified = &H189DC
        RavenHPLocBase = &H189EC
        RepedeHPLocCurrent = &H20854
        RepedeHPLocModified = &H2085C
        RepedeHPLocBase = &H2086C
        RitaHPLocCurrent = &H14A94
        RitaHPLocModified = &H14A9C
        RitaHPLocBase = &H14AAC
        YuriHPLocCurrent = &H8CD4
        YuriHPLocModified = &H8CDC
        YuriHPLocBase = &H8CEC

        'TP
        EstelleTPLocCurrent = &HCC18
        EstelleTPLocModified = &HCC20
        EstelleTPLocBase = &HCC30
        JudithTPLocCurrent = &H1C918
        JudithTPLocModified = &H1C920
        JudithTPLocBase = &H1C930
        KarolTPLocCurrent = &H10B58
        KarolTPLocModified = &H10B60
        KarolTPLocBase = &H10B70
        RavenTPLocCurrent = &H189D8
        RavenTPLocModified = &H189D0
        RavenTPLocBase = &H189E0
        RepedeTPLocCurrent = &H20858
        RepedeTPLocModified = &H20860
        RepedeTPLocBase = &H20870
        RitaTPLocCurrent = &H14A98
        RitaTPLocModified = &H14AA0
        RitaTPLocBase = &H14AB0
        YuriTPLocCurrent = &H8CD8
        YuriTPLocModified = &H8CE0
        YuriTPLocBase = &H8CF0

        'EXP
        EstelleEXPLoc = &HCC28
        JudithEXPLoc = &H1C928
        KarolEXPLoc = &H10B68
        RavenEXPLoc = &H189E8
        RepedeEXPLoc = &H20868
        RitaEXPLoc = &H14AA8
        YuriEXPLoc = &H8CE8

        'YuriLevel = &H8CD0
        YuriPhysicalAttack = &H8CF4
        YuriMagicAttack = &H8CF8
        YuriPhysicalDefense = &H8CFC
        YuriMagicDefense = &H8D00
        'YuriUnknown = &H8D04
        YuriAgility = &H8D08
        YuriLuck = &H8D0C
    End Enum
End Module
