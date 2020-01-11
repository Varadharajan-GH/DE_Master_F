Imports System.IO
Imports System.Text
Imports DEM_FCR_Library

Public Class frmCharMap
    Private Const CharsFontSize As Integer = 8
    Private Const NoOfCols As Integer = 16
    Private Const ColWidth As Integer = 33 '24
    Friend WithEvents chkKeepOnTop As CheckBox
    Public Shared lstCharCodes As List(Of String)
    Public Shared lstCharText As List(Of String)
    Dim DictCharMap As Dictionary(Of Integer, String)
    'Dim CharMapFile As String = frmLogin.sourcePath & "Admin\charmap.txt"
    Dim CharMapFile As String = SettingsReader.ReadSetting("SOURCE_PATH") _
                                & Path.DirectorySeparatorChar _
                                & SettingsReader.ReadSetting("CHARMAP_FILE")

    Private Sub frmCharMap_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        e.Cancel = True
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'lstCharCodes = {"162", "163", "167", "169", "174", "176", "177", "188", "189", "190", "197", "913", _
        '                "914", "915", "916", "917", "918", "919", "920", "921", "922", "923", "924", "925", _
        '                "926", "927", "928", "929", "931", "932", "933", "934", "935", "936", "937",
        '                "938", "939", "940", "941", "942", "943", "944", _
        '                "945", "946", "947", "948", "949", "950", "951", "952", "953", "954", "955", "956", _
        '                "957", "958", "959", "960", "961", "962", "963", "964", "965", "966", "967", "968", "969", _
        '                 "970", "971", "972", "973", "974", "975", "976", _
        '                 "8482", "8592", "8593", "8594", "8595", "8596", "8597", "8704", "8706", _
        '                "8712", "8713", "8723", "8730", "8733", "8734", "8747", "60", "62", "8804", "8805", "8810", "8811", _
        '                "8814", "8815", "8816", "8817", "8818", "8819", "8822", "8823", "8834", "8835", "8836", _
        '                "8837", "8838", "8839", "8841", "8842", "8924", "8925"}.ToList
        'lstCharText = {"&CENT;", "&POUND;", "&SECTION;", "&COPY;", "&REG;", "&DEG;", "&PLUSMN;", "1/4", "1/2", "3/4", "&ANGS;", _
        '               "&UAlpha;", "&UBeta;", "&UGamma;", "&UDelta;", "&UEpsilon;", "&UZeta;", "&UEta;", "&UTheta;", _
        '               "&UIota;", "&UKappa;", "&ULambda;", "&UMu;", "&UNu;", "&UXi;", "&UOmicron;", "&UPi;", "&URho;", _
        '               "&USigma;", "&UTau;", "&UUpsilon;", "&UPhi;", "&UChi;", "&UPsi;", "&UOmega;", _
        '               "&#938;", "&#939;", "&#940;", "&#941;", "&#942;", "&#943;", "&#944;", _
        '                "&alpha;", "&beta;", "&gamma;", "&delta;", "&epsilon;", "&zeta;", "&eta;", "&theta;", "&iota;", _
        '               "&kappa;", "&lambda;", "&mu;", "&nu;", "&xi;", "&omicron;", "&pi;", "&rho;", "", "&sigma;", "&tau;", _
        '               "&upsilon;", "&phi;", "&chi;", "&psi;", "&omega;", _
        '               "&#970;", "&#971;", "&#972;", "&#973;", "&#974;", "&#975;", "&#976;", _
        '               "&TRADE;", "&LARR;", "&UARR;", "&RARR;", "&DARR;", _
        '               "&LRARR;", "&UDARR;", "&FORALL;", "&PARTIAL;", _
        '                "&ISIN;", "&NI;", "&MNPLUS;", "&RADIC;", "&PROP;", "&INFIN;", "&INT;", "&LT;", "&GT;", "&LE;", "&GE;", "&MLT;", "&MGT;", _
        '                "&NLT;", "&NGT;", "&NLE;", "&NGE;", "&LSIM;", "&GSIM;", "&LGT;", "&GLT;", "&SUB;", "&SUP;", "&NSUB;", _
        '                "&NSUP;", "&SUBE;", "&SUPE;", "&NSUBE;", "&NSUPE;", "&ELT;", "&EGT;"}.ToList

        LoadCharMapDict()
        SetupDataGrid()
        CharGrid(0, 0).Selected = True
        Me.BringToFront()
        'Me.TopMost = True
    End Sub

    Private Sub AddCharsToGrid(ByVal Row As Integer, ByVal Index As Integer)
        For Counter As Integer = 0 To 15
            CharGrid(Counter, Row).Value = ChrW(Index + Counter)
        Next Counter
    End Sub

    Private Sub AddListToGrid(row As Integer, lstChars As List(Of String))
        Dim col As Integer = 0
        For Each ch In lstChars
            CharGrid(col Mod 16, row + CInt(Math.Floor(col / 16))).Value = ChrW(ch)
            'Debug.Print(col & " - " & row + CInt(Math.Floor(col / 16)) & " - " & ch & " - " & ChrW(ch))
            col = col + 1
        Next
    End Sub

    Private Sub SetupDataGrid()
        CharGrid.ColumnCount = NoOfCols

        For Each Column As DataGridViewColumn In CharGrid.Columns
            Column.Width = ColWidth
        Next Column
        'CharGrid.Rows.Add(14)
        'AddCharsToGrid(0, 913)
        'AddCharsToGrid(1, 929)
        'AddCharsToGrid(2, 945)
        'AddCharsToGrid(3, 961)
        'AddListToGrid(0, lstCharCodes)

        AddDictToGrid(DictCharMap)
    End Sub

    Private Sub AddDictToGrid(dict As Dictionary(Of Integer, String))
        Dim cell As Integer = 0

        CharGrid.Rows.Add(CInt(Math.Ceiling(dict.Count / NoOfCols)))
        For Each code As Integer In dict.Keys
            CharGrid(cell Mod NoOfCols, CInt(Math.Floor(cell / NoOfCols))).Value = ChrW(code)
            'Debug.Print(CInt(Math.Floor(cell / NoOfCols)) & " - " & cell Mod NoOfCols & " - " & ChrW(code))
            cell += 1
        Next
    End Sub

    Private Sub CharGrid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CharGrid.CellClick
        Dim ch As Char = CharGrid(e.ColumnIndex, e.RowIndex).Value
        'Dim CS As String = ""
        'frmFMain.insertText(ch & "-" & AscW(ch) & "-" & lstCharText(lstCharCodes.FindIndex(Function(value As String)
        'Return value = AscW(ch)
        '                                                                                  End Function)))
        'If DictCharMap.TryGetValue(AscW(ch), CS) Then
        frmFMain.insertText(ch)
        'End If
    End Sub

    Public Function ReplaceSymbols(strText As String) As String
        Dim icode As Integer, scode As String, sbtemp As New StringBuilder
        For Each ch As Char In strText
            icode = AscW(ch)
            scode = ""
            DictCharMap.TryGetValue(icode, scode)
            'ich = lstCharCodes.FindIndex(Function(value As String)
            'Return value = AscW(ch)
            '                             End Function)
            If scode = "" Then
                sbtemp.Append(ch)
            Else
                sbtemp.Append(scode)
            End If
        Next
        Return sbtemp.ToString
    End Function

    Private Sub chkKeepOnTop_CheckStateChanged(sender As Object, e As EventArgs) Handles chkKeepOnTop.CheckStateChanged
        If chkKeepOnTop.CheckState = CheckState.Checked Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Private Sub LoadCharMapDict()
        Dim TextLine As String = ""
        Dim CodeInt As Integer, CodeString As String

        DictCharMap = New Dictionary(Of Integer, String)
        If System.IO.File.Exists(CharMapFile) = True Then
            Dim objReader As New System.IO.StreamReader(CharMapFile)

            Do While (objReader.Peek() <> -1)
                TextLine = objReader.ReadLine()
                If (Not TextLine.Trim.StartsWith("#")) And (TextLine.Trim <> "") Then
                    Try
                        CodeInt = CInt(Split(TextLine.Trim, " - ").First)
                        CodeString = Split(TextLine.Trim, " - ").Last
                        DictCharMap.Add(CodeInt, CodeString)
                    Catch ex As Exception
                        ' Debug.Print("ERROR_LINE: " & TextLine)
                    End Try
                End If
            Loop
        Else
            MessageBox.Show("Charmap File Does Not Exist")
        End If
    End Sub

End Class
