Imports System.IO
Imports System.Xml
Imports DEM_FCR_Library

Structure Log
    Dim System_Type As String
    Dim System As String
    Dim Session_Start_Date As String
    Dim Session_Start_Time As String
    Dim Session_End_Date As String
    Dim Session_End_Time As String
    Dim Session_Production As String
    Dim Session_Suspend As String
    Dim Item_Accession As String
    Dim Item_Itemno As String
    Dim Status As String
    Dim references_Open As String
End Structure

Structure CI_INFO
    Dim Author As String
    Dim Volume As String
    Dim Page As String
    Dim Year As String
    Dim Title As String
    Dim ARTN1 As String
    Dim ARTNType1 As String
    Dim ARTN2 As String
    Dim ARTNType2 As String
    Dim ARTN3 As String
    Dim ARTNType3 As String
    Dim ARTN4 As String
    Dim ARTNType4 As String
End Structure

Public Class frmFMain
    Dim oResize As New Resizer
    Private dtTick As DateTime
    Dim fLog As New frmLogin
    Dim strRootPath As String '= frmLogin.sourcePath
    'Dim strRootPath As String = "\\172.21.15.100\data\F_Process\"
    'Dim strRootPath As String = ".\"

    Dim strInDir As String
    Dim strPriorDir As String
    Dim strCurDir As String
    Dim strOutDir As String
    'Dim strCompDir As String
    Dim strTempDir As String
    Dim strQueryDir As String
    Dim strInvalidDir As String
    Dim strLogPath As String
    Dim strPolicyFile As String

    Dim strInFile As String
    Dim bIsPrior As Boolean
    Dim iNo As Integer
    Dim xdDoc As New XmlDocument
    Dim swProdTime As Stopwatch = New Stopwatch
    Dim swSuspTime As Stopwatch = New Stopwatch
    Dim swCurProd As Stopwatch = New Stopwatch
    Dim swCurSusp As Stopwatch = New Stopwatch
    Dim xnlCitations As XmlNodeList
    Dim xpathOCR As String = "CI_CAPTURE/CI_CAPTURE_BLURB"
    Dim xpathImage As String = "CI_CAPTURE/CI_IMAGE_CLIP_NAME"
    Dim iNum As Integer
    Dim lstTags As List(Of String)
    Dim sbLogs As New System.Text.StringBuilder
    Dim cdText As XmlCDataSection
    Dim lstCharText As List(Of String)
    Dim xdElements As New XmlDocument
    Private blvrefs_events As Boolean
    Public lstRefNum As New List(Of Integer)
    Dim iCompItems As String
    Dim iCompRefs As String
    Dim LogEntry As New Log
    Private QC_Control As Integer
    Private RQC_Enabled As String
    Private RQCFile As String
    Private RQCExportDir As String

    Private Sub frmFMain_EnabledChanged(sender As Object, e As EventArgs) Handles Me.EnabledChanged
        If Me.Enabled = True Then
            swSuspTime.Stop()
            swProdTime.Start()
            swCurSusp.Stop()
            swCurProd.Start()
        Else
            swProdTime.Stop()
            swSuspTime.Start()
            swCurProd.Stop()
            swCurSusp.Start()
        End If
    End Sub

    Private Sub frmFMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.ApplicationExitCall Then
            frmChoose.Show()
        End If
        If MsgBox("Do you want to exit?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            e.Cancel = True
        Else
            frmChoose.Show()
        End If
    End Sub

    Private Sub frmFMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F3
                btnNext.PerformClick()
                e.Handled = True
            Case Keys.F4
                btnPrev.PerformClick()
                e.Handled = True
        End Select
    End Sub

    Private Sub frmFMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        strRootPath = SettingsReader.ReadSetting("SOURCE_PATH") & Path.DirectorySeparatorChar
        RQCExportDir = Path.Combine(strRootPath, "QC\Export\" & Date.Now.ToString("yyyyMMdd") & "\")
        RQCFile = Path.Combine(RQCExportDir, frmLogin.UserName & ".CSV")

        If frmChoose.chosenTool = ToolMode.QC Then
            strInDir = "Output\"
            strPriorDir = "QC\Priority\"
            strCurDir = "QC\Current\"
            strOutDir = "QC\Output\"
            'strCompDir = "QC\Completed\"
            strTempDir = "QC\Temp\"
            strQueryDir = "QC\Query\"
            strInvalidDir = "QC\Invalid\"
            strLogPath = strRootPath & "QC\F_Logs\" & Format(Today, "yyyyMMdd") & "\"
            'pnlVIL.Visible = True
        Else
            strInDir = "Input\"
            strPriorDir = "Priority\"
            strCurDir = "Current\"
            strOutDir = "Output\"
            'strCompDir = "Completed\"
            strTempDir = "Temp\"
            strQueryDir = "Query\"
            strInvalidDir = "Invalid\"
            strLogPath = strRootPath & "F_Logs\" & Format(Today, "yyyyMMdd") & "\"
            'pnlVIL.Visible = False
        End If

        RQC_Enabled = SettingsReader.ReadSetting("RQC_ENABLED")

        strPolicyFile = strRootPath & SettingsReader.ReadSetting("POLICY_FILE") '& "Admin\policy.txt"
        iCompItems = 0
        iCompRefs = 0
        oResize.FindAllControls(Me)
        Me.WindowState = FormWindowState.Maximized
        oResize.ResizeAllControls(Me)
        blvrefs_events = True
        frmCharMap.WindowState = FormWindowState.Minimized
        frmCharMap.Show()
        frmCharMap.Hide()
        xdElements.LoadXml("<?xml version='1.0' ?><Elements>" &
                           "<Book-Series><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><DOI></DOI><ARTN></ARTN><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Book-Series>" &
                           "<Communication><SOURCE_PUB_TITLE><General></General><Translated></Translated></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><MISC></MISC></Communication>" &
                           "<General><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></General>" &
                           "<InPress><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></InPress>" &
                           "<Journal><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Journal>" &
                           "<Magazine-Newspaper><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Magazine-Newspaper>" &
                           "<Meeting><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><VOLUME_ID></VOLUME_ID><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Meeting>" &
                           "<Patent><SOURCE_PUB_TITLE><General></General><Translated></Translated></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><NAME><Translator></Translator><Inventor></Inventor><Assignee></Assignee><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Translator></Translator><Inventor></Inventor><Assignee></Assignee></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><MISC></MISC></Patent>" &
                           "<Report><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Report>" &
                           "<Thesis><SOURCE_PUB_TITLE><General></General><Translated></Translated></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><VOLUME_ID></VOLUME_ID><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Thesis>" &
                           "<Unpublished><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Unpublished>" &
                           "<Non-Traditional-Reference><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Chapter></Chapter><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><ARTN></ARTN><DOI></DOI><Abst-Num></Abst-Num><Elocation-ID></Elocation-ID><Patent-Number></Patent-Number><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator><Inventor></Inventor><Assignee></Assignee></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></Non-Traditional-Reference>" &
                           "<DCI><SOURCE_PUB_TITLE><General></General><Translated></Translated><Series></Series></SOURCE_PUB_TITLE><SOURCE_ID><ISSN></ISSN><ISBN></ISBN><Other></Other></SOURCE_ID><PUBLISHER_NAME></PUBLISHER_NAME><PUBLISHER_LOC></PUBLISHER_LOC><ITEM_TITLE><Article></Article><Translated></Translated><Other></Other></ITEM_TITLE><EDITION></EDITION><VOLUME_ID></VOLUME_ID><ISSUE_ID></ISSUE_ID><SUPPLEMENT></SUPPLEMENT><ISSUE_TITLE></ISSUE_TITLE><PUB_DATE></PUB_DATE><PUB_ID><DOI></DOI><ARTN></ARTN><Elocation-ID></Elocation-ID><Other></Other></PUB_ID><PAGE_RANGE></PAGE_RANGE><SIZE><B></B><KB></KB><MB></MB><GB></GB><TB></TB><PB></PB><PP></PP></SIZE><NAME><Author></Author><Editor></Editor><Translator></Translator><Other></Other></NAME><EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></EXTERNAL_LINK><ANONYMOUS_IND><Other></Other></ANONYMOUS_IND><ET_AL_IND><Other></Other></ET_AL_IND><COLLAB_GRP_NAME></COLLAB_GRP_NAME><INSTITUTION></INSTITUTION><G_NAME><Author></Author><Editor></Editor><Translator></Translator></G_NAME><G_EXTERNAL_LINK><URI></URI><FTP></FTP><EMAIL></EMAIL><Other></Other></G_EXTERNAL_LINK><CONF_NAME></CONF_NAME><CONF_DATE></CONF_DATE><CONF_LOC></CONF_LOC><CONF_SPONSOR></CONF_SPONSOR><MISC></MISC></DCI>" &
                           "</Elements>")
        If Not My.Settings.MainBackColor.IsEmpty Then
            Me.BackColor = My.Settings.MainBackColor
            tpFCR.BackColor = My.Settings.MainBackColor
            tpQuery.BackColor = My.Settings.MainBackColor
        End If
        InitializeTagList()
        subInitContextMenu()
        lstCharText = {"&CENT;", "&POUND;", "&SECTION;", "&COPY;", "&REG;", "&DEG;", "&PLUSMN;", "1/4", "1/2", "3/4", "&ANGS;",
                       "&UAlpha;", "&UBeta;", "&UGamma;", "&UDelta;", "&UEpsilon;", "&UZeta;", "&UEta;", "&UTheta;",
                       "&UIota;", "&UKappa;", "&ULambda;", "&UMu;", "&UNu;", "&UXi;", "&UOmicron;", "&UPi;", "&URho;",
                       "&USigma;", "&UTau;", "&UUpsilon;", "&UPhi;", "&UChi;", "&UPsi;", "&UOmega;",
                        "&alpha;", "&beta;", "&gamma;", "&delta;", "&epsilon;", "&zeta;", "&eta;", "&theta;", "&iota;",
                       "&kappa;", "&lambda;", "&mu;", "&nu;", "&xi;", "&omicron;", "&pi;", "&rho;", "&sigma;", "&tau;",
                       "&upsilon;", "&phi;", "&chi;", "&psi;", "&omega;", "&TRADE;", "&LARR;", "&UARR;", "&RARR;", "&DARR;",
                       "&UDARR;", "&FORALL;", "&PARTIAL;",
                        "&ISIN;", "&NI;", "&MNPLUS;", "&RADIC;", "&PROP;", "&INFIN;", "&INT;", "&LE;", "&GE;", "&MLT;", "&MGT;",
                        "&NLT;", "&NGT;", "&NLE;", "&NGE;", "&LSIM;", "&GSIM;", "&LGT;", "&GLT;", "&SUB;", "&SUP;", "&NSUB;",
                        "&NSUP;", "&SUBE;", "&SUPE;", "&NSUBE;", "&NSUPE;", "&ELT;", "&EGT;"}.ToList
        sbLogs.AppendLine("Error logs for current session on " & Today.ToLongDateString)
        sbLogs.AppendLine("=========================================================")


        Directory.CreateDirectory(strRootPath & strOutDir)
        Directory.CreateDirectory(strRootPath & strCurDir)
        Directory.CreateDirectory(strRootPath & strTempDir)
        Directory.CreateDirectory(strRootPath & strPriorDir)
        Directory.CreateDirectory(strRootPath & strQueryDir)
        Directory.CreateDirectory(strRootPath & strInvalidDir)
        Directory.CreateDirectory(strLogPath)
        Directory.CreateDirectory(RQCExportDir)

        MakeLogFile(strLogPath & frmLogin.UserName & ".XML")
        LogEntry.System_Type = ""
        LogEntry.System = ""
        If Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList.Count > 3 Then
            LogEntry.System_Type = "IP_Address"
            LogEntry.System = Net.Dns.GetHostEntry(Net.Dns.GetHostName()).AddressList(3).ToString()
        Else
            LogEntry.System_Type = "Name"
            LogEntry.System = Net.Dns.GetHostName()
        End If
        'tryagain_Load:
        strInFile = GetXMLFile(strRootPath & strPriorDir)
        If String.IsNullOrEmpty(strInFile) Then
            strInFile = GetXMLFile(strRootPath & strInDir)
            bIsPrior = False
            If String.IsNullOrEmpty(strInFile) Then
                MsgBox("No input files found")
                'If MsgBox("No files for input. Copy files to input, then retry", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then
                'GoTo tryagain_Load
                'End If
            Else
                LoadFile(strInFile)
            End If
        Else
            bIsPrior = True
            LoadFile(strInFile)
        End If
        tsslStatus.Text = "Ready"
    End Sub

    Private Sub frmFMain_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin
        dtTick = Now
    End Sub

    Private Sub frmFMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim tickNow As DateTime = Now
        Dim delta As TimeSpan = tickNow - dtTick
        If (delta.TotalMilliseconds > 125) Then
            dtTick = tickNow
            oResize.ResizeAllControls(Me)
        End If
    End Sub

    Private Sub frmFMain_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        oResize.ResizeAllControls(Me)
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If lvTags.SelectedItems.Count = 1 Then
            lvTags.SelectedItems(0).Remove()
        End If
    End Sub

    Private Sub cmbElements_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbElements.SelectedIndexChanged
        rtbOCR.ContextMenuStrip = Nothing
        Select Case cmbElements.SelectedItem
            Case "Book-Series"
                rtbOCR.ContextMenuStrip = cmsBookSeries
            Case "Communication"
                rtbOCR.ContextMenuStrip = cmsCommunication
            Case "General"
                rtbOCR.ContextMenuStrip = cmsGeneral
            Case "InPress"
                rtbOCR.ContextMenuStrip = cmsInPress
            Case "Journal"
                rtbOCR.ContextMenuStrip = cmsJournal
            Case "Magazine-Newspaper"
                rtbOCR.ContextMenuStrip = cmsMagazineNewspaper
            Case "Meeting"
                rtbOCR.ContextMenuStrip = cmsMeeting
            Case "Patent"
                rtbOCR.ContextMenuStrip = cmsPatent
            Case "Report"
                rtbOCR.ContextMenuStrip = cmsReport
            Case "Thesis"
                rtbOCR.ContextMenuStrip = cmsThesis
            Case "Unpublished"
                rtbOCR.ContextMenuStrip = cmsUnpublished
            Case "Non-Traditional-Reference"
                rtbOCR.ContextMenuStrip = cmsNTR
                'rtbOCR.ContextMenuStrip = cmsLetters
            Case "DCI"
                rtbOCR.ContextMenuStrip = cmsDCI
        End Select
        'rtbOCR.Focus()
    End Sub

    Private Sub rtbOCR_KeyDown(sender As Object, e As KeyEventArgs) Handles rtbOCR.KeyDown
        Dim strSelText As String
        Try
            strSelText = frmCharMap.ReplaceSymbols(rtbOCR.SelectedText.Trim)
        Catch ex As Exception
            strSelText = rtbOCR.SelectedText.Trim
        End Try

        If (e.KeyCode = Keys.S) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("SOURCE_PUB_TITLE", "General", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.A) And (e.Modifiers = Keys.Control) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                Dim strText As String = strSelText
                If strText.ToLower.Contains(" of ") Or strText.ToLower.Contains(" the ") Then
                    strText = "<SURNAME>" & strText.Trim & "</SURNAME>"
                    AddTagToList("NAME", "Editor", strText)
                Else
                    Dim names() As String
                    If strText.Contains(",") Then
                        names = Split(strText, ",")
                    ElseIf strText.Contains(" ") Then
                        names = Split(strText, " ")
                    Else
                        ReDim names(1)
                        names(0) = strText
                        names(1) = ""
                    End If
                    Dim sname As String = names.First.Trim
                    Dim gname As String = names.Last.Trim

                    strText = "<SURNAME>" & sname.Trim & "</SURNAME><GIVEN_NAMES>" & gname.Trim & "</GIVEN_NAMES>"
                    AddTagToList("NAME", "Editor", strText)
                End If
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.K) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("EDITION", "", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.W) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("EXTERNAL_LINK", "URI", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.X) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("G_EXTERNAL_LINK", "URI", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.Z) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("PUB_ID", "Other", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.B) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("COLLAB_GRP_NAME", "", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.A) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                process_name("Author", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.M) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("MISC", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.E) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                process_name("Editor", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.H) And (e.Modifiers = Keys.Alt) Then

            'Dim strtext As String = strSelText
            'strtext = Process_Date(strtext)
            'AddTagToList("PUB_DATE", String.Empty, strtext)
            AddTagToList("PUB_DATE", String.Empty, Process_Date(strSelText))
            e.Handled = True
        ElseIf (e.KeyCode = Keys.V) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("VOLUME_ID", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.P) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("PAGE_RANGE", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.T) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("ITEM_TITLE", "Article", strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.O) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("PUBLISHER_LOC", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.C) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("CONF_NAME", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.U) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("PUBLISHER_NAME", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.I) And (e.Modifiers = Keys.Alt) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("ISSUE_ID", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf (e.KeyCode = Keys.P) And (e.Modifiers = Keys.Control) Then
            If strSelText = String.Empty Then
                MsgBox("Select text first")
            ElseIf cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
            Else
                AddTagToList("CONF_SPONSOR", String.Empty, strSelText)
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.F6 Then
            rtbOCR.SelectedText = "<SUP>" & rtbOCR.SelectedText & "</SUP>"
            e.Handled = True
        ElseIf e.KeyCode = Keys.F7 Then
            rtbOCR.SelectedText = "<INF>" & rtbOCR.SelectedText & "</INF>"
            e.Handled = True
        ElseIf e.KeyCode = Keys.F2 Then
            rtbOCR.SelectedText = rtbOCR.SelectedText.Replace(" ", String.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F9 Then
            Dim tstr As String = rtbOCR.SelectedText
            If tstr.ToUpper = tstr Then
                rtbOCR.SelectedText = rtbOCR.SelectedText.ToLower
            Else
                rtbOCR.SelectedText = rtbOCR.SelectedText.ToUpper
            End If
            e.Handled = True
        End If
        If e.Modifiers = Keys.Alt Then
            e.SuppressKeyPress = True
        End If
        Try
            tpFCR.Focus()
            rtbOCR.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub rtbOCR_MouseDown(sender As Object, e As MouseEventArgs) Handles rtbOCR.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If cmbElements.SelectedItem Is Nothing Then
                MsgBox("Choose an Element first", MsgBoxStyle.Information)
                Exit Sub
            ElseIf rtbOCR.SelectionLength = 0 Then
                MsgBox("Select text to tag first", MsgBoxStyle.Information)
            End If
        End If
    End Sub

    'Public Shared Function fnGetXMLFileold(strDir As String) As String
    '    Try
    '        'Return System.Directory.GetFiles(strDir, "*.XML", SearchOption.TopDirectoryOnly)(0)
    '        Dim fxml As String = System.Directory.GetFiles(strDir, "*.XML", SearchOption.TopDirectoryOnly)(0)
    '        Dim ftmp As String = Replace(Path.GetFileNameWithoutExtension(fxml), ".out", "", Compare:=CompareMethod.Text) & ".tmp"

    '        My.Computer.FileSystem.RenameFile(fxml, ftmp)
    '        Return Path.GetDirectoryName(fxml) & "\" & ftmp
    '    Catch e As System.IndexOutOfRangeException
    '        Return String.Empty
    '    End Try
    'End Function

    Public Shared Function GetXMLFile(strDir As String) As String
        Try
            'Dim fxml As String = System.Directory.GetFiles(strDir, "*.XML", SearchOption.TopDirectoryOnly)(0)
            Dim files() = New DirectoryInfo(strDir).GetFiles("*.XML", SearchOption.TopDirectoryOnly).OrderBy(Function(fi) fi.LastWriteTime).ToArray
            Dim fxml As String = files.First.ToString
            Dim ftmp As String = Path.ChangeExtension(fxml, "tmp")
            My.Computer.FileSystem.RenameFile(strDir & fxml, ftmp)
            Return strDir & ftmp
        Catch e As Exception
            Return String.Empty
        End Try
    End Function

    Private Sub LoadFile(strFile As String)
        iNo = 0
        Try
            File.Delete(strRootPath & strCurDir & Path.GetFileNameWithoutExtension(strFile) & ".XML")
        Catch ex As Exception
            Debug.Print(Err.Erl & " : " & ex.Message)
        End Try

        'My.Computer.FileSystem.RenameFile(strFile, Replace(Path.GetFileName(strFile), ".OUT", "", Compare:=CompareMethod.Text))
        'strFile = Replace(strFile, ".OUT", "", Compare:=CompareMethod.Text)
        My.Computer.FileSystem.MoveFile(strFile, strRootPath & strCurDir & Path.GetFileNameWithoutExtension(strFile) & ".XML")

        strFile = strRootPath & strCurDir & Path.GetFileNameWithoutExtension(strFile) & ".XML"

        File.Copy(strFile, strFile.Replace(strCurDir, strTempDir), True)
        strInFile = strInFile.Replace(".tmp", ".XML")
        Dim fname As String = Path.GetFileNameWithoutExtension(strFile)

        If fname.Contains(".") Then
            Dim fpath As String = Path.GetDirectoryName(strFile)
            Dim ext As String = Path.GetExtension(strFile)
            My.Computer.FileSystem.RenameFile(strFile, fname.Split(".").First & ext)
            strFile = fpath & "\" & fname.Split(".").First & ext
            strInFile = strInFile.Replace(fname, fname.Split(".").First)
        End If
        xdDoc = New XmlDocument
        Try
            xdDoc.Load(strFile)
        Catch ex As Exception
            MsgBox("Invalid XML File : " & Path.GetFileName(strFile) & vbCrLf &
                   "Source files will be moved to Invalid\ folder" & vbCrLf & vbCrLf &
                   "More Details: " & ex.Message, MsgBoxStyle.Exclamation)
            Try
                File.Move(strFile, Replace(strFile, strCurDir, strInvalidDir))
            Catch ex1 As Exception
            End Try
            Dim FileLocation As DirectoryInfo = New DirectoryInfo(strRootPath & strInDir)
            Dim fi As FileInfo() = FileLocation.GetFiles(Path.GetFileNameWithoutExtension(strInFile) & "*.TIF")

            For Each file In fi
                If Not file.Name.Contains("backup") Then
                    file.CopyTo(Replace(file.FullName, strInDir, strInvalidDir), True)
                    Try
                        file.Delete()
                    Catch ex1 As Exception
                        sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                    End Try
                End If
            Next file
            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                strInFile = GetXMLFile(strRootPath & strPriorDir)

                If String.IsNullOrEmpty(strInFile) Then
                    bIsPrior = False
                    strInFile = GetXMLFile(strRootPath & strInDir)
                End If
                If String.IsNullOrEmpty(strInFile) Then
                    MsgBox("No files for input", MsgBoxStyle.Information)
                Else
                    bIsPrior = False
                    LoadFile(strInFile)
                End If
            End If
            Exit Sub
        End Try

        swCurProd.Restart()
        swCurSusp.Reset()

        tsslStatus.Text = "File loaded : " & strFile

        xnlCitations = xdDoc.DocumentElement.SelectNodes("//CI_CITATION")
        If xnlCitations.Count = 0 Then
            tsslStatus.Text = "No citations to load."
            'Exit Sub
        End If
        CheckDeleted()
        lvRefs.Items.Clear()
        lblAccn.Text = xdDoc.SelectSingleNode("//ID_ACCESSION").InnerText
        lblItem.Text = xdDoc.SelectSingleNode("//ITEM").Attributes.GetNamedItem("ITEMNO").InnerText

        lstRefNum = New List(Of Integer)
        For i = 0 To xnlCitations.Count - 1
            If (xnlCitations(i).SelectSingleNode("UT_CODE") Is Nothing) And (xnlCitations(i).Attributes("D") Is Nothing) Then
                Dim lvi As New ListViewItem(lvRefs.Items.Count + 1)
                Dim strOCR As String = String.Empty
                Try
                    strOCR = xnlCitations(i).SelectSingleNode(xpathOCR).InnerText
                Catch ex As Exception
                    sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                End Try
                lvi.SubItems.Add(strOCR)
                lvRefs.Items.Add(lvi)
                lstRefNum.Add(i)
            Else
                Try
                    xnlCitations(i).RemoveChild(xnlCitations(i).SelectSingleNode("FULL_CI_INFO"))
                Catch ex As Exception
                    sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                End Try
            End If
        Next

        Dim xnode As XmlNode
        InitLog(LogEntry)
        LogEntry.Session_Start_Date = Format(Today, "dd-MM-yyyy")
        LogEntry.Session_Start_Time = Now.ToLongTimeString

        xnode = xdDoc.SelectSingleNode("//ID_ACCESSION")
        If xnode IsNot Nothing Then LogEntry.Item_Accession = xnode.InnerText
        xnode = Nothing

        xnode = xdDoc.SelectSingleNode("//ITEM")
        If xnode IsNot Nothing Then xnode = xnode.Attributes.GetNamedItem("ITEMNO")
        If xnode IsNot Nothing Then LogEntry.Item_Itemno = xnode.InnerText
        xnode = Nothing

        LogEntry.references_Open = GetValidRefsCount()
        LogEntry.Session_Production = swCurProd.Elapsed.ToString("hh\:mm\:ss")
        LogEntry.Session_Suspend = swCurSusp.Elapsed.ToString("hh\:mm\:ss")
        LogEntry.Session_End_Date = Format(Today, "dd-MM-yyyy")
        LogEntry.Session_End_Time = Now.ToLongTimeString

        lblTotRef.Text = lstRefNum.Count
        If lstRefNum.Count = 0 Then

            Dim TiffOutDir As String

            If MsgBox("No valid references in File  " & Path.GetFileName(strFile) & vbCrLf &
                   "Do you want to move this to Output folder anyway", vbYesNo) = vbYes Then
                Try
                    LogEntry.Status = "CO"
                    File.Move(strFile, Replace(strFile, strCurDir, strOutDir))
                Catch ex1 As Exception
                End Try
                TiffOutDir = strOutDir
            Else
                Try
                    LogEntry.Status = "INC"
                    File.Move(strFile, Replace(strFile, strCurDir, strInvalidDir))
                Catch ex1 As Exception
                End Try
                TiffOutDir = strInvalidDir
            End If

            Dim FileLocation As DirectoryInfo = New DirectoryInfo(strRootPath & strInDir)
            Dim fi As FileInfo() = FileLocation.GetFiles(Path.GetFileNameWithoutExtension(strInFile) & "*.TIF")
            For Each file In fi
                If Not file.Name.Contains("backup") Then
                    file.CopyTo(Replace(file.FullName, strInDir, TiffOutDir), True)
                    Try
                        file.Delete()
                    Catch ex As Exception
                        sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                    End Try
                End If
            Next file
            MakeLogEntry(strLogPath & frmLogin.UserName & ".XML")
            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                strInFile = GetXMLFile(strRootPath & strPriorDir)

                If String.IsNullOrEmpty(strInFile) Then
                    bIsPrior = False
                    strInFile = GetXMLFile(strRootPath & strInDir)
                End If
                If String.IsNullOrEmpty(strInFile) Then
                    MsgBox("No files for input", MsgBoxStyle.Information)
                Else
                    bIsPrior = False
                    LoadFile(strInFile)
                End If
            End If
            Exit Sub
        End If
        iNum = 0

        LogEntry.Status = "IP"
        MakeLogEntry(strLogPath & frmLogin.UserName & ".XML")

        lblCurRef.Text = iNum
        subLoadRefs(xnlCitations, lstRefNum(iNo), "Next")
    End Sub

    Private Sub CheckDeleted()
        For Each nod As XmlNode In xnlCitations
            If nod.Attributes("D") IsNot Nothing Then
                If nod.Attributes("D").Value = "Y" Then
                    Dim nodFCI As XmlNode
                    nodFCI = nod.SelectSingleNode("FULL_CI_INFO")
                    If nodFCI IsNot Nothing Then
                        nodFCI.ParentNode.RemoveChild(nodFCI)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub makeTree(xnNode As XmlNode, strTree As String, Optional value As String = "!@!@!",
                               Optional InsertFirst As Boolean = False, Optional InsertLast As Boolean = False,
                               Optional InsertAfter As XmlNode = Nothing, Optional InsertBefore As XmlNode = Nothing)
        Dim nod As XmlNode, nodnew As XmlNode, xeElement As XmlElement
        nod = Nothing
        nodnew = Nothing
        xeElement = Nothing
        nod = xnNode
        'For Each str As String In Split(strTree, "/")
        '    If Trim(str) <> String.Empty Then
        '        nodnew = nod.SelectSingleNode(str)
        '        If nodnew Is Nothing Then
        '            xeElement = xnNode.OwnerDocument.CreateElement(str)
        '            nod = nod.InsertAfter(xeElement, nod.LastChild)
        '        Else
        '            nod = nodnew
        '        End If
        '    End If
        'Next
        For Each str As String In Split(strTree, "/")
            If Trim(str) <> String.Empty Then
                nodnew = nod.SelectSingleNode(str)
                If nodnew Is Nothing Then
                    xeElement = xnNode.OwnerDocument.CreateElement(str)
                    If str = Split(strTree, "/").Last Then
                        If InsertFirst Then
                            If nod.FirstChild IsNot Nothing Then
                                nod = nod.InsertBefore(xeElement, nod.FirstChild)
                            Else
                                nod = nod.AppendChild(xeElement)
                            End If
                        ElseIf InsertLast Then
                            nod = nod.AppendChild(xeElement)
                        ElseIf InsertAfter IsNot Nothing Then
                            nod = nod.InsertAfter(xeElement, InsertAfter)
                        ElseIf InsertBefore IsNot Nothing Then
                            nod = nod.InsertBefore(xeElement, InsertBefore)
                        Else
                            nod = nod.AppendChild(xeElement)
                        End If
                    Else
                        nod = nod.AppendChild(xeElement)
                    End If
                Else
                    nod = nodnew
                End If
            End If
        Next
        If value <> "!@!@!" Then
            nod.InnerText = value
        End If
    End Sub

    Public Sub subLoadRefs(cits As XmlNodeList, iRNum As Integer, sender As String)

        If cits Is Nothing Then
            tsslStatus.Text = "No Citation to load"
            Exit Sub
        End If
        If (iRNum < 0) Or (iRNum >= cits.Count) Then Exit Sub
        If cits(iRNum).SelectSingleNode("UT_CODE") IsNot Nothing Then
            Try
                cits(iRNum).RemoveChild(cits(iRNum).SelectSingleNode("FULL_CI_INFO"))
            Catch ex As Exception
            End Try
            If sender = "Next" Then
                subLoadRefs(cits, lstRefNum(iRNum + 1), "Next")
            ElseIf sender = "Prev" Then
                subLoadRefs(cits, lstRefNum(iRNum - 1), "Prev")
            End If
            Exit Sub
        End If
        Dim inClipName As String = String.Empty
        subClearFields()
        lblCurRef.Text = lstRefNum.IndexOf(iRNum)
        lblTotRef.Text = lstRefNum.Count
        getAllTags(cits(iRNum).SelectSingleNode("FULL_CI_INFO"))
        'If frmChoose.chosenTool = ToolMode.QC Then
        '    getCIInfo(cits(iRNum).SelectSingleNode("CI_INFO"))
        'End If

        Try
            Dim nTemp As XmlNode = cits(iRNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE")
            If nTemp IsNot Nothing Then
                Dim sTemp As String = nTemp.InnerText
                If sTemp = "NON TRADITIONAL REF" Then
                    cmbElements.Text = "Non-Traditional-Reference"
                Else
                    sTemp = ""
                    Try
                        sTemp = cits(iRNum).SelectSingleNode("CI_INFO/CI_JOURNAL/CI_TITLE").InnerText
                    Catch ex1 As Exception
                    End Try

                    If sTemp = "**DATA OBJECT**" Then
                        cmbElements.Text = "DCI"
                    Else
                        cmbElements.Text = cits(iRNum).SelectSingleNode("FULL_CI_INFO").Attributes("type").Value
                    End If
                End If
            Else
                cmbElements.Text = cits(iRNum).SelectSingleNode("FULL_CI_INFO").Attributes("type").Value
            End If
        Catch ex As Exception
        End Try

        Try
            inClipName = cits(iRNum).SelectSingleNode(xpathImage).InnerText
        Catch ex As Exception
            sbLogs.AppendLine(Err.Erl & " : " & ex.Message)
        End Try

        If inClipName = vbNullString Then
            makeTree(cits(iRNum), xpathImage, ".\" & StrConv(cits(iRNum).OwnerDocument.SelectSingleNode("//ID_ACCESSION").InnerText _
                             & cits(iRNum).OwnerDocument.SelectSingleNode("//ITEM").Attributes("ITEMNO").Value & "C_" _
                             & cits(iRNum).Attributes("seq").InnerText & ".TIF", VbStrConv.Uppercase))
            inClipName = cits(iRNum).SelectSingleNode(xpathImage).InnerText
        End If
        inClipName = Replace(inClipName, Split(inClipName, "_").Last, cits(iRNum).Attributes("seq").InnerText & ".TIF")
        cits(iRNum).SelectSingleNode(xpathImage).InnerText = inClipName

        If My.Computer.FileSystem.FileExists(strRootPath & strInDir & Split(inClipName, "\").Last) Then
            pbImage.ImageLocation = strRootPath & strInDir & Split(inClipName, "\").Last
        End If


        Try
            rtbOCR.Text = cits(iRNum).SelectSingleNode(xpathOCR).InnerText
        Catch ex As Exception
            makeTree(cits(iRNum), xpathOCR)
            rtbOCR.Text = String.Empty
        End Try
        iNum = iRNum
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click

        If lblCurRef.Text.Trim <> lblTotRef.Text.Trim Then
            If SaveReference() = -1 Then Exit Sub
        End If
        If iNum <= 0 Then
            If lblCurRef.Text.Trim <> lblTotRef.Text.Trim Then
                'subLoadRefs(xnlCitations, lstRefNum(lstRefNum.IndexOf(iNum) - 1), "Prev")
                MsgBox("No more references to view", MsgBoxStyle.Information, "Cannot load Previous reference")
                Exit Sub
            Else
                subLoadRefs(xnlCitations, lstRefNum.Last, "Prev")
            End If
        Else
            Try
                If lblCurRef.Text.Trim <> lblTotRef.Text.Trim Then
                    subLoadRefs(xnlCitations, lstRefNum(lstRefNum.IndexOf(iNum) - 1), "Prev")
                Else
                    subLoadRefs(xnlCitations, lstRefNum.Last, "Prev")
                End If
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If lblCurRef.Text.Trim = lblTotRef.Text.Trim Then
            Exit Sub
        End If
        If cmbElements.SelectedItem Is Nothing Then
            MsgBox("Select an Element first")
            cmbElements.Focus()
            Exit Sub
        End If

        If lblCurRef.Text.Trim <> lblTotRef.Text.Trim Then
            If SaveReference() = -1 Then Exit Sub
        End If
        If xnlCitations Is Nothing Then
            tsslStatus.Text = "No citations to load"
            Exit Sub
        End If
        If iNum >= xnlCitations.Count - 1 Then
            MsgBox("No more references to view", MsgBoxStyle.Information, "Cannot load Next reference")
            subClearFields()
            lblCurRef.Text = lstRefNum.Count
            lblTotRef.Text = lstRefNum.Count
            iNum = lstRefNum.Last
            Exit Sub
        Else
            Try
                subLoadRefs(xnlCitations, lstRefNum(lstRefNum.IndexOf(iNum) + 1), "Next")
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                MsgBox("No more references to view", MsgBoxStyle.Information, "Cannot load Next reference")
                subClearFields()
                lblCurRef.Text = lstRefNum.Count
                lblTotRef.Text = lstRefNum.Count
                iNum = lstRefNum.Last
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub EventDeleteButtonClicked(sender As Object, e As EventArgs) Handles btnDelete.Click
        If MsgBox("All tags in list will be deleted. Continue?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.Yes Then
            lvTags.Items.Clear()
        End If
    End Sub

    Private Sub InitializeTagList()
        lstTags = {"SOURCE_PUB_TITLE", "SOURCE_ID", "PUBLISHER_NAME", "PUBLISHER_LOC", "ITEM_TITLE", "EDITION",
                   "VOLUME_ID", "ISSUE_ID", "SUPPLEMENT", "ISSUE_TITLE", "PUB_DATE", "PUB_ID", "PAGE_RANGE",
                   "SIZE", "NAME", "EXTERNAL_LINK", "ANONYMOUS_IND", "ET_AL_IND", "COLLAB_GRP_NAME", "INSTITUTION",
                   "G_NAME", "G_EXTERNAL_LINK", "CONF_NAME", "CONF_DATE", "CONF_LOC", "CONF_SPONSOR", "MISC"}.ToList
    End Sub

    Public Function SaveReference() As Integer

        If xnlCitations Is Nothing Then
            tsslStatus.Text = "No citations to load"
            Return -1
        End If

        If fnValidateTags() = -1 Then Return -1
        tsslStatus.Text = "Saving Reference..."
        Dim bIsCollab As Boolean = False, bIsGExt As Boolean = False
        Dim bIsName As Boolean = False, bIsExt As Boolean = False
        Dim totChars As Integer = 0
        Dim sbTagText As New Text.StringBuilder

        For Each item As ListViewItem In lvTags.Items
            Select Case item.SubItems(0).Text
                Case "NAME"
                    bIsName = True
                Case "COLLAB_GRP_NAME"
                    bIsCollab = True
                Case "EXTERNAL_LINK"
                    bIsExt = True
                Case "G_EXTERNAL_LINK"
                    bIsGExt = True
            End Select
            totChars = totChars + CharsInTag(item.SubItems(2).Text.Replace(" ", String.Empty))
            sbTagText.Append(TextInTag(item.SubItems(2).Text))
        Next


        Dim xnNode As XmlNode
        Try
            xnNode = xnlCitations(iNum).SelectSingleNode("CI_INFO/CI_JOURNAL")
        Catch
        End Try

        Try
            Dim nod As XmlNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO")
            nod.ParentNode.RemoveChild(nod)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
        Dim xanodeType As XmlAttribute = xdDoc.CreateAttribute("type")
        If cmbElements.SelectedItem = "Non-Traditional-Reference" Then
            xanodeType.Value = "Non-Traditional Reference"
            If lvTags.Items.Count > 0 Then
                MsgBox("NTR Reference should not contain tags")
                Return -1
            End If
        Else
            xanodeType.Value = cmbElements.SelectedItem
        End If
        If cmbElements.SelectedItem <> "Non-Traditional-Reference" Then
            tsslStatus.Text = totChars & " chars tagged of " & rtbOCR.Text.Replace(" ", String.Empty).Length &
                                                ". Min limit: " & CInt(rtbOCR.Text.Replace(" ", String.Empty).Length * (4 / 5)) + 1 & " chars"

            If totChars < rtbOCR.Text.Replace(" ", String.Empty).Length * (4 / 5) Then
                MsgBox("Atleast 80% of Blurrb must be tagged")
            End If
        End If

        If cmbElements.SelectedItem = "DCI" Then
            Dim tnode As XmlNode = Nothing
            Try
                tnode = xnlCitations(iNum).SelectSingleNode("DATASETID")
            Catch ex As Exception
            End Try
            If tnode Is Nothing Then
                tnode = xdDoc.CreateElement("DATASETID")
                tnode.InnerText = "UNKNOWN"
                xnlCitations(iNum).InsertAfter(tnode, xnlCitations(iNum).SelectSingleNode("CI_INFO"))
            End If
        End If

        'Try
        '    xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO").Attributes.Append(xanodeType)
        'Catch ex As Exception
        '    Dim telem As XmlElement
        '    telem = xdDoc.CreateElement("FULL_CI_INFO")
        '    xnlCitations(iNum).InsertAfter(telem, xnlCitations(iNum).SelectSingleNode("CI_INFO"))
        '    xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO").Attributes.Append(xanodeType)
        'End Try

        xnNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO")
        If xnNode Is Nothing Then
            Dim prenode As XmlNode
            prenode = xnlCitations(iNum).SelectSingleNode("DATASETID")
            If prenode Is Nothing Then prenode = xnlCitations(iNum).SelectSingleNode("UT_CODE")
            If prenode Is Nothing Then prenode = xnlCitations(iNum).SelectSingleNode("RI_CITATIONIDENTIFIER")
            If prenode Is Nothing Then prenode = xnlCitations(iNum).SelectSingleNode("CI_INFO")

            makeTree(xnlCitations(iNum), "FULL_CI_INFO", InsertAfter:=prenode)
            xnNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO")
        End If
        Try
            xnNode.Attributes.Append(xanodeType)
        Catch ex As Exception
            tsslStatus.Text = ex.Message
        End Try


        lvTags.BeginUpdate()
        ReorderTags()
        lvTags.EndUpdate()

        For Each item As ListViewItem In lvTags.Items
            Dim xeElem As XmlElement

            xeElem = xdDoc.CreateElement(item.SubItems(0).Text.Replace("G_", "").Trim)
            If item.SubItems(1).Text <> String.Empty Then
                Dim xaAttr As XmlAttribute
                Select Case item.SubItems(0).Text
                    Case "NAME", "G_NAME"
                        xaAttr = xdDoc.CreateAttribute("role")
                        Dim strText As String = item.SubItems(2).Text
                        Dim strTemp() As String
                        Dim sname As String = "", gname As String = ""
                        If strText.Contains("</SURNAME><GIVEN_NAMES>") Then
                            strTemp = Split(strText, "</SURNAME><GIVEN_NAMES>")
                            sname = Replace(strTemp(0), "<SURNAME>", String.Empty)
                            gname = Replace(strTemp(1), "</GIVEN_NAMES>", String.Empty)
                        Else
                            sname = Replace(strText, "<SURNAME>", String.Empty)
                            sname = Replace(sname, "</SURNAME>", String.Empty)
                        End If

                        Dim xeTemp As XmlElement
                        xeTemp = xdDoc.CreateElement("SURNAME")
                        cdText = xdDoc.CreateCDataSection(sname)
                        xeTemp.AppendChild(cdText)
                        xeElem.AppendChild(xeTemp)
                        xeTemp = Nothing
                        If gname <> "" Then
                            xeTemp = xdDoc.CreateElement("GIVEN_NAMES")
                            cdText = xdDoc.CreateCDataSection(gname)
                            xeTemp.AppendChild(cdText)
                            xeElem.AppendChild(xeTemp)
                            xeTemp = Nothing
                        End If
                        xaAttr.Value = item.SubItems(1).Text
                        xeElem.Attributes.Append(xaAttr)
                    Case "SOURCE_ID"
                        xaAttr = xdDoc.CreateAttribute("type")
                        xaAttr.Value = item.SubItems(1).Text
                        xeElem.Attributes.Append(xaAttr)
                        xeElem.InnerText = item.SubItems(2).Text.Trim
                    Case "SIZE"
                        xaAttr = xdDoc.CreateAttribute("units")
                        xaAttr.Value = item.SubItems(1).Text
                        xeElem.Attributes.Append(xaAttr)
                        cdText = xdDoc.CreateCDataSection(item.SubItems(2).Text.Trim)
                        xeElem.AppendChild(cdText)
                    Case Else
                        xaAttr = xdDoc.CreateAttribute("type")
                        xaAttr.Value = item.SubItems(1).Text
                        xeElem.Attributes.Append(xaAttr)
                        cdText = xdDoc.CreateCDataSection(item.SubItems(2).Text.Trim)
                        xeElem.AppendChild(cdText)
                End Select
            Else
                Select Case item.SubItems(0).Text
                    Case "PUB_DATE", "CONF_DATE"
                        Dim strText As String = item.SubItems(2).Text
                        If strText.Contains("<DATE><YEAR>") Then
                            strText = strText.Replace("<DATE><YEAR>", String.Empty)
                            strText = strText.Replace("</YEAR></DATE>", String.Empty).Trim
                            cdText = xdDoc.CreateCDataSection(strText)
                            Dim xeTemp As XmlElement = xdDoc.CreateElement("DATE")
                            xeTemp.AppendChild(xdDoc.CreateElement("YEAR"))
                            xeElem.AppendChild(xeTemp)
                            xeElem.FirstChild.FirstChild.AppendChild(cdText)
                        ElseIf strText.Contains("</DATE_STRING><YEAR>") Then
                            Dim dtstr As String = strText.Replace("</DATE_STRING><YEAR>", "#").Split("#")(0).Replace("<DATE><DATE_STRING>", "").Trim
                            Dim yr As String = strText.Replace("</DATE_STRING><YEAR>", "#").Split("#>")(1).Replace("</YEAR></DATE>", "").Trim
                            cdText = xdDoc.CreateCDataSection(dtstr)
                            Dim xeTemp As XmlElement = xdDoc.CreateElement("DATE")
                            xeTemp.AppendChild(xdDoc.CreateElement("DATE_STRING"))
                            xeTemp.FirstChild.AppendChild(cdText)
                            cdText = xdDoc.CreateCDataSection(yr)
                            xeTemp.AppendChild(xdDoc.CreateElement("YEAR"))
                            xeTemp.LastChild.AppendChild(cdText)
                            xeElem.AppendChild(xeTemp)
                            'xeElem.FirstChild.FirstChild.AppendChild(cdText)
                        ElseIf strText.Contains("</DATE_STRING></YEAR>") Then
                            strText = strText.Replace("</DATE_STRING></YEAR>", String.Empty)
                            strText = strText.Replace("<DATE><DATE_STRING>", String.Empty).Trim
                            cdText = xdDoc.CreateCDataSection(strText)
                            Dim xeTemp As XmlElement = xdDoc.CreateElement("DATE")
                            xeTemp.AppendChild(xdDoc.CreateElement("DATE_STRING"))
                            xeElem.AppendChild(xeTemp)
                            xeElem.FirstChild.FirstChild.AppendChild(cdText)
                        Else
                            strText = strText.Replace("</DATE_STRING></DATE>", String.Empty)
                            strText = strText.Replace("<DATE><DATE_STRING>", String.Empty).Trim
                            cdText = xdDoc.CreateCDataSection(strText)
                            Dim xeTemp As XmlElement = xdDoc.CreateElement("DATE")
                            xeTemp.AppendChild(xdDoc.CreateElement("DATE_STRING"))
                            xeElem.AppendChild(xeTemp)
                            xeElem.FirstChild.FirstChild.AppendChild(cdText)
                        End If
                    Case "ET_AL_IND"
                        cdText = xdDoc.CreateCDataSection("Y")
                        xeElem.AppendChild(cdText)
                    Case "ANONYMOUS_IND"
                        cdText = xdDoc.CreateCDataSection("Y")
                        xeElem.AppendChild(cdText)
                    Case "SOURCE_ID"
                        xeElem.InnerText = item.SubItems(2).Text.Trim
                    Case Else
                        cdText = xdDoc.CreateCDataSection(item.SubItems(2).Text.Trim)
                        xeElem.AppendChild(cdText)
                End Select
            End If

            xnNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO")
            Select Case xeElem.Name
                Case "SOURCE_PUB_TITLE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/SOURCE_PUB_TITLES")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "SOURCE_PUB_TITLES")
                    Else
                        xnNode = xnCheckNode
                    End If
                    'xeElem = makeParent(xeElem, "SOURCE_PUB_TITLES")
                Case "SOURCE_ID"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/SOURCE_IDENTIFIERS")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "SOURCE_IDENTIFIERS")
                    Else
                        xnNode = xnCheckNode
                    End If
                    'xeElem = makeParent(xeElem, "SOURCE_IDENTIFIERS")
                Case "PUBLISHER_NAME"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PUBLISHER")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "PUBLISHER")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "PUBLISHER_LOC"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PUBLISHER")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "PUBLISHER")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "ITEM_TITLE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/ITEM_TITLES")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "ITEM_TITLES")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "EDITION"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CI_PUB_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "VOLUME_ID"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO/VOLUME_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "VOLUME_INFO")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "ISSUE_ID"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO/ISSUE_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "ISSUE_INFO")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "ISSUE_TITLE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO/ISSUE_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "ISSUE_INFO")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "SUPPLEMENT"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO/ISSUE_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "ISSUE_INFO")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "PUB_DATE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CI_PUB_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "PAGE_RANGE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PAGINATION_SIZE")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "PAGINATION_SIZE")
                    Else
                        xnNode = xnCheckNode
                    End If
                    Try
                        Dim parr() As String = {"FPAGE", "LPAGE"}, i As Integer = 0, pr() As String
                        pr = Split(xeElem.InnerText, "-")
                        If pr(0).ToString.Trim = String.Empty Then pr(0) = pr(1)
                        If pr(1).ToString.Trim = String.Empty Then pr(1) = pr(0)
                        For i = 0 To 1
                            Dim xtemp As XmlNode = xdDoc.CreateElement(parr(i))
                            xtemp.InnerText = pr(i).ToString.Trim
                            xeElem.AppendChild(xtemp)
                            xtemp = Nothing
                        Next

                    Catch ex As Exception
                        sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                    End Try
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CI_PUB_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "SIZE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO/PAGINATION_SIZE")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "PAGINATION_SIZE")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "EXTERNAL_LINK"
                    If item.SubItems(0).Text = "G_EXTERNAL_LINK" Then
                        Dim xnCheckNode As XmlNode
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "COLLAB")
                            xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                            If xnCheckNode Is Nothing Then
                                xeElem = makeParent(xeElem, "PERSON_GROUP")
                            Else
                                xnNode = xnCheckNode
                            End If
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        Dim xnCheckNode As XmlNode
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CI_PUB_INFO")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "CI_PUB_INFO")
                        Else
                            xnNode = xnCheckNode
                        End If
                    End If
                Case "G_EXTERNAL_LINK"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "COLLAB")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "PUB_ID"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PUB_IDENTIFIERS")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "PUB_IDENTIFIERS")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "COLLAB_GRP_NAME"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "COLLAB")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "INSTITUTION"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "COLLAB")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "NAME"
                    If item.SubItems(0).Text = "G_NAME" Then
                        Dim xnCheckNode As XmlNode
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "COLLAB")
                            xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                            If xnCheckNode Is Nothing Then
                                xeElem = makeParent(xeElem, "PERSON_GROUP")
                            Else
                                xnNode = xnCheckNode
                            End If
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        Dim xnCheckNode As XmlNode
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    End If
                Case "G_NAME"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "COLLAB")
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "CONF_NAME"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CONF_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CONF_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "CONF_DATE"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CONF_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CONF_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "CONF_LOC"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CONF_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CONF_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "CONF_SPONSOR"
                    Dim xnCheckNode As XmlNode
                    xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/CONF_INFO")
                    If xnCheckNode Is Nothing Then
                        xeElem = makeParent(xeElem, "CONF_INFO")
                    Else
                        xnNode = xnCheckNode
                    End If
                Case "ET_AL_IND"
                    Dim xnCheckNode As XmlNode
                    xeElem = makeParent(xeElem, "NAME")
                    Dim xaAttr As XmlAttribute
                    xaAttr = xdDoc.CreateAttribute("role")
                    xaAttr.Value = "Other"
                    xeElem.Attributes.Append(xaAttr)
                    If item.SubItems(0).Text = "G_NAME" Then
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "COLLAB")
                            xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                            If xnCheckNode Is Nothing Then
                                xeElem = makeParent(xeElem, "PERSON_GROUP")
                            Else
                                xnNode = xnCheckNode
                            End If
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    End If

                    'Dim xaAttr As XmlAttribute
                    'xaAttr = xdDoc.CreateAttribute("role")
                    'xaAttr.Value = "Other"
                    'xeElem.Attributes.Append(xaAttr)
                    'xnNode = xnCheckNode
                Case "ANONYMOUS_IND"
                    Dim xnCheckNode As XmlNode
                    xeElem = makeParent(xeElem, "NAME")
                    Dim xaAttr As XmlAttribute
                    xaAttr = xdDoc.CreateAttribute("role")
                    xaAttr.Value = "Other"
                    xeElem.Attributes.Append(xaAttr)
                    If item.SubItems(0).Text = "G_NAME" Then
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP/COLLAB")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "COLLAB")
                            xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                            If xnCheckNode Is Nothing Then
                                xeElem = makeParent(xeElem, "PERSON_GROUP")
                            Else
                                xnNode = xnCheckNode
                            End If
                        Else
                            xnNode = xnCheckNode
                        End If
                    Else
                        xnCheckNode = xnlCitations(iNum).SelectSingleNode("FULL_CI_INFO/PERSON_GROUP")
                        If xnCheckNode Is Nothing Then
                            xeElem = makeParent(xeElem, "PERSON_GROUP")
                        Else
                            xnNode = xnCheckNode
                        End If
                    End If
                    'Dim xaAttr As XmlAttribute
                    'xaAttr = xdDoc.CreateAttribute("role")
                    'xaAttr.Value = "Other"
                    'xeElem.Attributes.Append(xaAttr)
                    'xnNode = xnCheckNode
            End Select
            xnNode.AppendChild(xeElem)
        Next
        cdText = Nothing


        cdText = xdDoc.CreateCDataSection(frmCharMap.ReplaceSymbols(rtbOCR.Text.Trim))
        xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").InnerXml = String.Empty
        xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_BLURB").AppendChild(cdText)
        lvRefs.Items(lstRefNum.IndexOf(iNum)).SubItems(1).Text = rtbOCR.Text
        If cmbElements.SelectedItem = "Non-Traditional-Reference" Then
            xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText = "NON TRADITIONAL REF"
        Else
            Try
                If xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText = "NON TRADITIONAL REF" Then
                    xnlCitations(iNum).SelectSingleNode("CI_CAPTURE/CI_CAPTURE_TITLE").InnerText = ""
                End If
            Catch ex As Exception
            End Try
        End If
lbl_deleted:

        xdDoc.Save(strRootPath & strTempDir & Split(xdDoc.BaseURI, "/").Last)
        xnlCitations = xdDoc.SelectNodes("//CI_CITATION")
        tsslStatus.Text = "Ready"
        Return 1
    End Function

    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strType As String = CType(sender, ToolStripMenuItem).Text
        Dim lvi As New ListViewItem(strType)

        Dim strText As String = frmCharMap.ReplaceSymbols(rtbOCR.SelectedText.Trim)

        Select Case strType
            Case "PUB_DATE", "CONF_DATE"
                'Dim tstrtext() As String
                'Dim dtDate As Date
                'tstrtext = {"2011-12", "31 may 2018", "2017-07-17", " May 2016", "2015 Jul 30", "2014 Aug", "2013 Aug to Sep",
                '            "Aug to Sep 2013 ", "2012 Jul 30- Aug 30", " Jul 30- Aug 30 2012 ", "2011", "20 may", "1911 sept 23"
                '            }

                'for Each str As String In tstrtext
                ' strText = str
                strText = Process_Date(strText)
                If strText = String.Empty Then
                    MsgBox("Invalid year")
                    Exit Sub
                End If
            Case "Subscript"
                rtbOCR.SelectedText = "<INF>" & strText & "</INF>"
                Exit Sub
            Case "Superscript"
                rtbOCR.SelectedText = "<SUP>" & strText & "</SUP>"
                Exit Sub
            Case "UPPER CASE"
                rtbOCR.SelectedText = rtbOCR.SelectedText.ToUpper
                Exit Sub
            Case "lower case"
                rtbOCR.SelectedText = rtbOCR.SelectedText.ToLower
                Exit Sub
        End Select


        If (strType = "ET_AL_IND") Or (strType = "ANONYMOUS_IND") Then
            lvi.SubItems.Add(String.Empty)
            lvi.SubItems.Add("Y")
        Else
            lvi.SubItems.Add(String.Empty)
            lvi.SubItems.Add(strText)
        End If

        lvTags.Items.Add(lvi)
        lvi.EnsureVisible()
        rtbOCR.Focus()
    End Sub

    Private Shared Function Process_Date(strText As String) As String
        Dim i As Integer, j As Integer

        For i = 0 To strText.Length - 4
            If IsNumeric(strText.Substring(i, 4)) Then
                For j = i To i + 3
                    If Not IsNumeric(strText(j)) Then
                        Exit For
                    End If
                Next
                If j > i + 3 Then
                    Exit For
                End If
            End If
        Next

        Dim yr As String, dstr As String
        If strText.Length < 4 Then
            yr = ""
            dstr = strText
        ElseIf i > strText.Length - 4 Then
            yr = ""
            dstr = strText
        Else
            yr = strText.Substring(i, 4)
            dstr = strText.Replace(yr, "").Trim
        End If
        If yr <> "" And dstr = "" Then
            If Not IsNumeric(strText) Then Return String.Empty
            If CInt(strText) > Now.Year Then Return String.Empty
            strText = "<DATE><YEAR>" & strText & "</YEAR></DATE>"
        ElseIf dstr <> "" And yr = "" Then
            strText = "<DATE><DATE_STRING>" & strText & "</DATE_STRING></DATE>"
        ElseIf ((yr & dstr) = strText.Replace(" ", "")) Or ((dstr & yr) = strText.Replace(" ", "")) Then
            If dstr.Contains("-") Then
                If IsNumeric(dstr) And ((yr & dstr) = strText.Replace(" ", "")) Then
                    If dstr(0) = "-" Then dstr = dstr.Substring(1, dstr.Length - 1)
                    If dstr.Last = "-" Then dstr = dstr.Substring(0, dstr.Length - 1)
                    If Not IsNumeric(strText) Then Return String.Empty
                    If CInt(strText) > Now.Year Then Return String.Empty
                    strText = "<DATE><YEAR>" & strText & "</YEAR></DATE>"
                ElseIf yr = "" Then
                    If dstr(0) = "-" Then dstr = dstr.Substring(1, dstr.Length - 1)
                    If dstr.Last = "-" Then dstr = dstr.Substring(0, dstr.Length - 1)
                    strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING></DATE>"
                Else
                    If Not IsNumeric(yr) Then Return String.Empty
                    If CInt(yr) > Now.Year Then Return String.Empty
                    strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
                End If
            ElseIf dstr.Contains("/") Then
                If dstr(0) = "/" Then dstr = dstr.Substring(1, dstr.Length - 1)
                If dstr.Last = "/" Then dstr = dstr.Substring(0, dstr.Length - 1)
                If yr = "" Then
                    strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING></DATE>"
                Else
                    If IsNumeric(dstr) And (dstr.Length = 4) Then
                        If Not IsNumeric(strText) Then Return String.Empty
                        If CInt(strText) > Now.Year Then Return String.Empty
                        strText = "<DATE><YEAR>" & strText & "</YEAR></DATE>"
                    Else
                        If Not IsNumeric(yr) Then Return String.Empty
                        If CInt(yr) > Now.Year Then Return String.Empty
                        strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
                    End If
                End If
            ElseIf yr <> "" And dstr = "" Then
                If Not IsNumeric(strText) Then Return String.Empty
                If CInt(strText) > Now.Year Then Return String.Empty
                strText = "<DATE><YEAR>" & strText & "</YEAR></DATE>"
            ElseIf dstr <> "" And yr = "" Then
                strText = "<DATE><DATE_STRING>" & strText & "</DATE_STRING></DATE>"
            Else
                If Not IsNumeric(yr) Then Return String.Empty
                If CInt(yr) > Now.Year Then Return String.Empty
                strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
            End If
        ElseIf dstr.Contains(" to ") Then
            If Not IsNumeric(yr) Then Return String.Empty
            If CInt(yr) > Now.Year Then Return String.Empty
            strText = "<DATE><DATE_STRING>" & dstr.Split(" to ").First.Trim & "</DATE_STRING><YEAR>" & yr & "</YEAR><DATE_STRING> to " & dstr.Split(New String() {" to "}, StringSplitOptions.None).Last.Trim & "</DATE_STRING></DATE>"
        Else
            If dstr = "" Then
                If Not IsNumeric(yr) Then Return String.Empty
                If CInt(yr) > Now.Year Then Return String.Empty
                strText = "<DATE><YEAR>" & yr & "</YEAR></DATE>"
            ElseIf yr = "" Then
                strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING></DATE>"
            ElseIf dstr.Contains("-") Then
                If dstr(0) = "-" Then dstr = dstr.Substring(1, dstr.Length - 1)
                If dstr.Last = "-" Then dstr = dstr.Substring(0, dstr.Length - 1)
                If dstr.Contains("-") Then
                    Dim tsplt() As String
                    tsplt = dstr.Split("-")
                    Dim bYr As Boolean = False
                    Dim ts As String = ""
                    For Each ts In tsplt
                        If IsNumeric(ts.Trim) Then
                            If ts.Trim.Length = 4 Then
                                bYr = True
                                Exit For
                            End If
                        End If
                    Next
                    If bYr Then
                        yr = yr & "-" & ts
                        dstr = dstr.Replace(ts, "")
                        dstr = dstr.Replace("-", "")
                    End If
                Else
                    If dstr.Contains(" ") Then
                        Dim tsplt() As String
                        tsplt = dstr.Split(" ")
                        Dim bYr As Boolean = False
                        Dim ts As String = ""
                        For Each ts In tsplt
                            If IsNumeric(ts.Trim) Then
                                If ts.Trim.Length = 4 Then
                                    bYr = True
                                    Exit For
                                End If
                            End If
                        Next
                        If bYr Then
                            yr = yr & "-" & ts
                            dstr = dstr.Replace(ts, "")
                            dstr = dstr.Replace("-", "")
                        End If
                    End If
                End If
                If dstr.Trim = "" Then
                    If Not IsNumeric(yr) Then Return String.Empty
                    If CInt(yr) > Now.Year Then Return String.Empty
                    strText = "<DATE><YEAR>" & yr & "</YEAR></DATE>"
                ElseIf yr.Trim = "" Then
                    strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING></DATE>"
                Else
                    If Not IsNumeric(yr) Then Return String.Empty
                    If CInt(yr) > Now.Year Then Return String.Empty
                    strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
                End If
            ElseIf dstr.Contains("/") Then
                If dstr(0) = "/" Then dstr = dstr.Substring(1, dstr.Length - 1)
                If dstr.Last = "/" Then dstr = dstr.Substring(0, dstr.Length - 1)
                If Not IsNumeric(yr) Then Return String.Empty
                If CInt(yr) > Now.Year Then Return String.Empty
                strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
            Else
                If Not IsNumeric(yr) Then Return String.Empty
                If CInt(yr) > Now.Year Then Return String.Empty
                strText = "<DATE><DATE_STRING>" & dstr & "</DATE_STRING><YEAR>" & yr & "</YEAR></DATE>"
            End If
        End If

        Return strText
    End Function

    Private Sub ToolStripSubMenuItem_Click(sender As Object, e As EventArgs)
        Dim strText As String = frmCharMap.ReplaceSymbols(rtbOCR.SelectedText.Trim)
        Dim strType As String = CType(sender.owneritem, ToolStripMenuItem).Text
        Dim strSubItem As String = CType(sender, ToolStripMenuItem).Text

        If strType.Contains("LINK") Then
            strText = strText.Replace(" ", String.Empty)
        End If
        If strType = "PUB_ID" Then
            strText = strText.Replace(" ", String.Empty)
        End If
        If strSubItem = "DOI" Then
            'strText = strText.Replace(" ", String.Empty)
            If Not isValidDOI(strText) Then
                MsgBox("Invalid DOI format")
                Exit Sub
            End If
        End If
        Select Case strType
            Case "NAME"
                process_name(CType(sender, ToolStripMenuItem).Text, rtbOCR.SelectedText.Trim)
            Case Else
                AddTagToList(strType, CType(sender, ToolStripMenuItem).Text, strText)
        End Select
        rtbOCR.Focus()
    End Sub

    Private Function isValidDOI(doi As String) As Boolean
        Dim regex = New System.Text.RegularExpressions.Regex("^(10[.][0-9]{4,5}/(\S+))\b")
        'Dim regex = New Regex("\b(10[.][0-9]{4,}/(\S+))\b")
        If regex.Match(doi).Success Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub subInitContextMenu()
        For Each ctl As Object In Me.components.Components
            If ctl.GetType.ToString = "System.Windows.Forms.ContextMenuStrip" Then
                If ctl.Name = "cmsTag" Then
                    Continue For
                End If
                For Each Item As ToolStripMenuItem In ctl.Items
                    If Item.HasDropDownItems Then
                        DoSubItems(Item)
                    Else
                        AddHandler Item.Click, AddressOf ToolStripMenuItem_Click
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub DoSubItems(ByVal Item As ToolStripMenuItem)
        For Each SubItem As ToolStripMenuItem In Item.DropDownItems
            If SubItem.HasDropDownItems Then
                DoSubItems(SubItem)
            Else
                AddHandler SubItem.Click, AddressOf ToolStripSubMenuItem_Click
            End If
        Next
    End Sub

    Private Function makeParent(elem As XmlElement, parname As String) As XmlElement
        Dim xeTemp As XmlElement
        xeTemp = xdDoc.CreateElement(parname)
        xeTemp.AppendChild(elem)
        Return xeTemp
    End Function

    Private Sub FileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.Click
        btnNext.PerformClick()
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        btnPrev.PerformClick()
    End Sub

    Private Sub FirstToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FirstToolStripMenuItem.Click
        If SaveReference() = -1 Then Exit Sub
        subLoadRefs(xnlCitations, lstRefNum(0), "Next")
    End Sub

    Private Sub LastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LastToolStripMenuItem.Click
        If SaveReference() = -1 Then Exit Sub
        If xnlCitations Is Nothing Then
            tsslStatus.Text = "No citations to load"
            Exit Sub
        End If
        subLoadRefs(xnlCitations, lstRefNum(lstRefNum.Count - 1), "Prev")
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        If MsgBox("Do you want to quit", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = vbYes Then
            If File.Exists(strRootPath & strCurDir & Path.GetFileName(strInFile)) Then
                If bIsPrior Then
                    Try
                        File.Delete(strRootPath & strPriorDir & Path.GetFileName(strInFile))
                    Catch ex As Exception
                    End Try
                    File.Move(strRootPath & strCurDir & Path.GetFileName(strInFile), strRootPath & strPriorDir & Path.GetFileName(strInFile))
                Else
                    Try
                        File.Delete(strRootPath & strInDir & Path.GetFileName(strInFile))
                    Catch ex As Exception
                    End Try
                    File.Move(strRootPath & strCurDir & Path.GetFileName(strInFile), strRootPath & strInDir & Path.GetFileName(strInFile))
                End If
            End If
            Me.Close()
        End If
    End Sub

    Public Function getCitations() As XmlNodeList
        Return xnlCitations
    End Function

    Private Sub GoToPageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoToPageToolStripMenuItem.Click
        Try
            frmGoTo.Visible = False
        Catch ex As Exception
        End Try
        frmGoTo.Show(Me)
    End Sub

    Private Sub DoneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DoneToolStripMenuItem.Click
        If ValidateFCIInfo() = False Then
            Exit Sub
        End If
        If MsgBox("Do you Done this item?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = vbYes Then
            If cmbElements.SelectedItem IsNot Nothing Then
                If SaveReference() = -1 Then
                    If lblCurRef.Text.Trim <> lblTotRef.Text.Trim Then
                        MsgBox("Error saving current reference.")
                        Exit Sub
                    End If
                End If
            End If
            If String.IsNullOrEmpty(strInFile) Then
                tsslStatus.Text = "File not loaded. Could not Save now."
                Exit Sub
            End If
            '--------------------------------------------

            Dim TempOutDir As String = strOutDir

            'todo review

            If frmChoose.chosenTool = ToolMode.OP Then
                Try
                    If strInFile.Contains(strPriorDir) Then
                        File.Copy(Replace(strInFile, strPriorDir, strTempDir), Replace(strInFile, strPriorDir, "QC\" & strPriorDir), True)
                    Else
                        If RQC_Enabled.ToUpper = "TRUE" Then
                            If QC_Control <= 0 Then
                                File.Copy(Replace(strInFile, strInDir, strTempDir), Replace(strInFile, strInDir, strOutDir), True)
                                QC_Control = 1
                            Else
                                TempOutDir = strOutDir.Replace("Output\", "QC\Output\")
                                File.Copy(Replace(strInFile, strInDir, strTempDir), Replace(strInFile, strInDir, TempOutDir), True)
                                Dim csvText As String
                                csvText = Date.Now.ToShortDateString
                                csvText &= ", " & Date.Now.ToLongTimeString
                                csvText &= ", " & Path.GetFileNameWithoutExtension(strInFile) & vbCrLf
                                File.AppendAllText(RQCFile, csvText)
                                If QC_Control >= 3 Then
                                    QC_Control = 0
                                Else
                                    QC_Control += 1
                                End If
                            End If
                        Else
                            File.Copy(Replace(strInFile, strInDir, strTempDir), Replace(strInFile, strInDir, strOutDir), True)
                        End If
                    End If
                Catch ex As Exception
                    sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                End Try
            Else
                Try
                    If strInFile.Contains(strPriorDir) Then
                        File.Copy(Replace(strInFile, strPriorDir, strTempDir), Replace(strInFile, strPriorDir, strOutDir), True)
                    Else
                        File.Copy(Replace(strInFile, strInDir, strTempDir), Replace(strInFile, strInDir, strOutDir), True)
                    End If
                Catch ex As Exception
                    sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                End Try
            End If

            Dim FileLocation As DirectoryInfo = New DirectoryInfo(strRootPath & strInDir)
            Dim fi As FileInfo() = FileLocation.GetFiles(Path.GetFileNameWithoutExtension(strInFile) & "*.TIF")
            For Each file In fi
                If Not file.Name.Contains("backup") Then
                    file.CopyTo(Replace(file.FullName, strInDir, TempOutDir), True)
                    Try
                        file.Delete()
                    Catch ex As Exception
                        sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                    End Try
                End If
            Next file

            '-------------------------------
            Try
                If strInFile.Contains(strPriorDir) Then
                    File.Delete(Replace(strInFile, strPriorDir, strTempDir))
                Else
                    File.Delete(Replace(strInFile, strInDir, strTempDir))
                End If
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
            End Try
            Try
                If strInFile.Contains(strPriorDir) Then
                    File.Delete(Replace(strInFile, strPriorDir, strCurDir))
                Else
                    File.Delete(Replace(strInFile, strInDir, strCurDir))
                End If
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
            End Try
            subClearFields()

            Dim TotRefs As Integer = xnlCitations.Count
            Dim DelRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@D='Y']").Count
            Dim InsRefs As Integer = xnlCitations(0).ParentNode.SelectNodes("CI_CITATION[@I='Y']").Count

            swCurProd.Stop()
            swCurSusp.Stop()

            lvRefs.Items.Clear()
            strInFile = String.Empty
            xdDoc = Nothing
            xnlCitations = Nothing
            iCompItems += 1
            iCompRefs += lstRefNum.Count
            lstRefNum.Clear()

            'LogEntry.references_Close = TotRefs + InsRefs - DelRefs
            'LogEntry.references_Deleted = DelRefs
            'LogEntry.references_Inserted = InsRefs
            LogEntry.Session_Production = swCurProd.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_Suspend = swCurSusp.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_End_Date = Format(Today, "dd-MM-yyyy")
            LogEntry.Session_End_Time = Now.ToLongTimeString
            LogEntry.Status = "CO"
            UpdateLogEntry(strLogPath & frmLogin.UserName & ".XML", LogEntry)



            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                'tryagain_Done:
                strInFile = GetXMLFile(strRootPath & strPriorDir)
                If strInFile = "" Then
                    bIsPrior = False
                    strInFile = GetXMLFile(strRootPath & strInDir)
                End If
                If strInFile = "" Then
                    MsgBox("No input files found")
                    'If MsgBox("No files for input. Copy files to input, then retry", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then
                    'GoTo tryagain_Done
                    'End If
                Else
                    bIsPrior = False
                    LoadFile(strInFile)
                End If
            End If
        End If
    End Sub

    Private Sub subClearFields()
        lblCurRef.Text = String.Empty
        lblTotRef.Text = String.Empty
        lvTags.Items.Clear()
        pbImage.Image = Nothing
        cmbElements.SelectedItem = Nothing
        rtbOCR.Clear()
        'txtAuthor.Clear()
        'txtVolume.Clear()
        'txtPage.Clear()
        'txtYear.Clear()
        'txtTitle.Clear()
        'txtARTN1.Clear()
        'cmbARTN1.SelectedItem = Nothing
    End Sub

    Private Sub getAllTags(citnode As XmlNode)
        If citnode Is Nothing Then
            tsslStatus.Text = "No citations to load"
            Exit Sub
        End If
        For Each xnnode As XmlNode In citnode.ChildNodes
            If lstTags.FindIndex(Function(value As String)
                                     Return value = xnnode.Name
                                 End Function) <> -1 Then
                If xnnode.Name = "NAME" Then
                    If xnnode.Attributes("role").Value = "Other" Then
                        Dim lvi As ListViewItem
                        If xnnode.FirstChild.Name = "ANONYMOUS_IND" Then
                            lvi = New ListViewItem("ANONYMOUS_IND")
                        Else
                            lvi = New ListViewItem("ET_AL_IND")
                        End If
                        lvi.SubItems.Add(String.Empty)
                        'lvi.SubItems.Add("et al")
                        lvi.SubItems.Add("Y")                    'etal anony
                        lvTags.Items.Add(lvi)
                        lvi.EnsureVisible()
                    Else
                        Dim lvi As ListViewItem
                        If xnnode.ParentNode.Name = "COLLAB" Then
                            lvi = New ListViewItem("G_" & xnnode.Name)
                        Else
                            lvi = New ListViewItem(xnnode.Name)
                        End If
                        Dim strSubType As String
                        If xnnode.Attributes.Count > 0 Then
                            strSubType = xnnode.Attributes(0).Value
                        Else
                            strSubType = String.Empty
                        End If
                        Dim strText As String = xnnode.InnerXml
                        If strText.Contains("<![CDATA[") And strText.Contains("]]>") Then
                            strText = Replace(strText, "<![CDATA[", String.Empty)
                            strText = Replace(strText, "]]>", String.Empty)
                        End If
                        lvi.SubItems.Add(strSubType)
                        lvi.SubItems.Add(strText)
                        lvTags.Items.Add(lvi)
                        lvi.EnsureVisible()
                    End If
                Else
                    Dim lvi As ListViewItem
                    If (xnnode.Name = "EXTERNAL_LINK") And (xnnode.ParentNode.Name = "COLLAB") Then
                        lvi = New ListViewItem("G_" & xnnode.Name)
                    Else
                        lvi = New ListViewItem(xnnode.Name)
                    End If

                    Dim strSubType As String
                    If xnnode.Attributes.Count > 0 Then
                        strSubType = xnnode.Attributes(0).Value
                    Else
                        strSubType = String.Empty
                    End If
                    Dim strText As String = xnnode.InnerXml
                    If strText.Contains("<![CDATA[") And strText.Contains("]]>") Then
                        strText = Replace(strText, "<![CDATA[", String.Empty)
                        strText = Replace(strText, "]]>", String.Empty)
                    End If
                    lvi.SubItems.Add(strSubType)
                    lvi.SubItems.Add(strText)
                    lvTags.Items.Add(lvi)
                    lvi.EnsureVisible()
                    lvi.EnsureVisible()
                End If
            Else
                getAllTags(xnnode)
            End If
        Next
    End Sub

    Private Sub ShortcutKeysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShortcutKeysToolStripMenuItem.Click
        Dim sbHelp As New System.Text.StringBuilder, tabs As String
        tabs = "                     "
        sbHelp.AppendLine("       Shortcut keys Reference       ")
        sbHelp.AppendLine("-------------------------------------")
        sbHelp.AppendLine("Ctrl + F1 " & tabs & "- Opens this window")
        sbHelp.AppendLine("Ctrl + P  " & tabs & "- Conf Sponsor")
        sbHelp.AppendLine("Alt  + A  " & tabs & "- Name Author")
        sbHelp.AppendLine("Alt  + B  " & tabs & "- Collab Group Name")
        sbHelp.AppendLine("Alt  + C  " & tabs & "- Conf Name")
        sbHelp.AppendLine("Alt  + D  " & tabs & "- Done")
        sbHelp.AppendLine("Alt  + E  " & tabs & "- Name Editor")
        sbHelp.AppendLine("Alt  + F  " & tabs & "- First Reference")
        sbHelp.AppendLine("Alt  + G  " & tabs & "- Goto Reference")
        sbHelp.AppendLine("Alt  + I  " & tabs & "- Issue ID")
        sbHelp.AppendLine("Alt  + L  " & tabs & "- Last Reference")
        sbHelp.AppendLine("Alt  + M  " & tabs & "- MISC")
        sbHelp.AppendLine("Alt  + N  " & tabs & "- Next Reference")
        sbHelp.AppendLine("Alt  + O  " & tabs & "- Publisher Location")
        sbHelp.AppendLine("Alt  + P  " & tabs & "- Previous Reference")
        sbHelp.AppendLine("Alt  + Q  " & tabs & "- Quit")
        sbHelp.AppendLine("Alt  + R  " & tabs & "- Page Range")
        sbHelp.AppendLine("Alt  + S  " & tabs & "- Source Pub Title(General)")
        sbHelp.AppendLine("Alt  + T  " & tabs & "- Item Title")
        sbHelp.AppendLine("Alt  + U  " & tabs & "- Publisher Name")
        sbHelp.AppendLine("Alt  + V  " & tabs & "- Volume")
        sbHelp.AppendLine("Alt  + W  " & tabs & "- Pub ID/ General")
        sbHelp.AppendLine("Alt  + X  " & tabs & "- G External Link")
        sbHelp.AppendLine("Alt  + Y  " & tabs & "- Year")
        sbHelp.AppendLine("Alt  + Z  " & tabs & "- Pub ID General")
        sbHelp.AppendLine("PgUp      " & tabs & "- Alter Name forward")
        sbHelp.AppendLine("Shift+PgUp" & tabs & "- Alter Name backward")
        sbHelp.AppendLine("PgDown    " & tabs & "- Switch Name Tag")
        sbHelp.AppendLine("F2        " & tabs & "- Remove Spaces")
        sbHelp.AppendLine("F3        " & tabs & "- Next Reference")
        sbHelp.AppendLine("F4        " & tabs & "- Previous Reference")
        sbHelp.AppendLine("F6        " & tabs & "- Superscript")
        sbHelp.AppendLine("F7        " & tabs & "- Subscript")
        sbHelp.AppendLine("F9        " & tabs & "- Change case")
        sbHelp.AppendLine()

        Dim ufHelp As New Form, lblHelp As New Label, chkOnTop As New CheckBox
        chkOnTop.Text = "Show on top"
        lblHelp.Text = sbHelp.ToString
        lblHelp.Font = New Font("Times New Roman", 12)
        lblHelp.AutoSize = True
        'ufHelp.AutoSize = True
        'ufHelp.AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink
        ufHelp.Size = New Size(500, 800)
        ufHelp.AutoScroll = True
        ufHelp.Controls.Add(chkOnTop)
        chkOnTop.Location = New Point(10, 0)
        AddHandler chkOnTop.CheckStateChanged, AddressOf chkOnTop_CheckStateChanged
        ufHelp.Controls.Add(lblHelp)
        lblHelp.Location = New Point(10, 20)
        ufHelp.Text = "Shortcuts"
        ufHelp.Show()
        ufHelp.StartPosition = FormStartPosition.CenterParent
    End Sub

    Private Sub lvRefs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvRefs.SelectedIndexChanged
        If lvRefs.SelectedIndices.Count > 0 Then
            If cmbElements.SelectedItem = String.Empty Then
                MsgBox("Select an element first")
                Exit Sub
            End If
            If SaveReference() = -1 Then Exit Sub
            subLoadRefs(xnlCitations, lstRefNum(lvRefs.SelectedIndices(0)), "Next")
        End If
    End Sub

    Private Sub lvTags_KeyDown(sender As Object, e As KeyEventArgs) Handles lvTags.KeyDown
        If lvTags.SelectedItems.Count = 0 Then Exit Sub
        If e.KeyData = Keys.Delete Then
            If MsgBox("Do you want to remove selected Tag", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                lvTags.SelectedItems(0).Remove()
            End If
            e.Handled = True
        ElseIf e.KeyData = Keys.PageDown Then
            If lvTags.SelectedItems(0).SubItems(0).Text = "NAME" Then
                Dim strText As String = lvTags.SelectedItems(0).SubItems(2).Text
                Dim strTemp() As String = Split(strText, "</SURNAME><GIVEN_NAMES>")
                Dim sname As String = Replace(strTemp(0), "<SURNAME>", String.Empty)
                If strTemp.Length > 1 Then
                    Dim gname As String = Replace(strTemp(1), "</GIVEN_NAMES>", String.Empty)
                    lvTags.SelectedItems(0).SubItems(2).Text = "<SURNAME>" & gname & "</SURNAME><GIVEN_NAMES>" & sname & "</GIVEN_NAMES>"
                End If
                e.Handled = True
            End If
        ElseIf (e.KeyCode = Keys.PageUp) And (e.Modifiers = Keys.Shift) Then
            If lvTags.SelectedItems(0).SubItems(0).Text = "NAME" Then
                Dim strText As String = lvTags.SelectedItems(0).SubItems(2).Text
                Dim strTemp() As String = Split(strText, "</SURNAME><GIVEN_NAMES>")
                Dim sname As String = Replace(strTemp.First, "<SURNAME>", String.Empty).Trim
                If strTemp.Length > 1 Then
                    Dim gname As String = Replace(strTemp(1), "</GIVEN_NAMES>", String.Empty).Trim
                    If Not sname.Contains(" ") Then
                        e.Handled = True
                        Exit Sub
                    End If
                    Dim snarr() As String = Split(sname, " ")
                    gname = snarr.Last & " " & gname
                    sname = Replace(sname, snarr.Last, String.Empty).Trim
                    lvTags.SelectedItems(0).SubItems(2).Text = "<SURNAME>" & sname & "</SURNAME><GIVEN_NAMES>" & gname & "</GIVEN_NAMES>"
                End If
                e.Handled = True
            End If

        ElseIf (e.KeyCode = Keys.PageUp) And (e.Modifiers <> Keys.Shift) Then
            If lvTags.SelectedItems(0).SubItems(0).Text = "NAME" Then
                Dim strText As String = lvTags.SelectedItems(0).SubItems(2).Text
                Dim strTemp() As String = Split(strText, "</SURNAME><GIVEN_NAMES>")
                Dim sname As String = Replace(strTemp.First, "<SURNAME>", String.Empty).Trim
                If strTemp.Length > 1 Then
                    Dim gname As String = Replace(strTemp(1), "</GIVEN_NAMES>", String.Empty).Trim
                    If Not gname.Contains(" ") Then
                        e.Handled = True
                        Exit Sub
                    End If

                    Dim gnarr() As String = Split(gname, " ")
                    sname = sname & " " & gnarr.First
                    gname = Replace(gname, gnarr.First, String.Empty).Trim
                    lvTags.SelectedItems(0).SubItems(2).Text = "<SURNAME>" & sname & "</SURNAME><GIVEN_NAMES>" & gname & "</GIVEN_NAMES>"
                End If
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub lvTags_MouseClick(sender As Object, e As MouseEventArgs) Handles lvTags.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If lvTags.SelectedItems.Count = 0 Then
                tsslStatus.Text = "No tags selected"
                Exit Sub
            End If
            Dim cmsContext As New ContextMenuStrip()
            Dim saMenus() As String
            saMenus = {""}

            Select Case lvTags.SelectedItems(0).SubItems(0).Text
                Case "SOURCE_PUB_TITLE"
                    Select Case cmbElements.SelectedItem
                        Case "Book-Series", "DCI", "General", "Non-Traditional-Reference", "InPress", "Journal",
                                "Magazine-Newspaper", "Meeting", "Report", "Unpublished"
                            saMenus = {"General", "Translated", "Series", "Delete"}
                        Case "Communication", "Patent", "Thesis"
                            saMenus = {"General", "Translated", "Delete"}
                    End Select

                Case "SOURCE_ID"
                    saMenus = {"ISSN", "ISBN", "Other", "Delete"}

                Case "ITEM_TITLE"
                    Select Case cmbElements.SelectedItem
                        Case "Book-Series", "DCI"
                            saMenus = {"Chapter", "Translated", "Other", "Delete"}
                        Case "Communication", "Journal", "Magazine-Newspaper", "Meeting",
                            "Patent", "Thesis"
                            saMenus = {"Article", "Translated", "Other", "Delete"}
                        Case "General", "Non-Traditional-Reference", "InPress", "Report", "Unpublished"
                            saMenus = {"Article", "Chapter", "Translated", "Other", "Delete"}
                    End Select

                Case "PUB_ID"
                    Select Case cmbElements.SelectedItem
                        Case "Book-Series", "DCI", "Journal", "Magazine-Newspaper"
                            saMenus = {"ARTN", "DOI", "Elocation-ID", "Other", "Delete"}
                        Case "Communication", "Meeting", "Thesis"
                            saMenus = {"ARTN", "DOI", "Abst-Num", "Elocation-ID", "Other", "Delete"}
                        Case "General", "Non-Traditional-Reference", "InPress", "Report", "Unpublished"
                            saMenus = {"ARTN", "DOI", "Abst-Num", "Elocation-ID", "Patent-Number", "Other", "Delete"}
                        Case "Patent"
                            saMenus = {"ARTN", "DOI", "Elocation-ID", "Patent-Number", "Other", "Delete"}
                    End Select

                Case "SIZE"
                    saMenus = {"B", "KB", "MB", "GB", "TB", "PB", "PP", "Delete"}

                Case "NAME"
                    Select Case cmbElements.SelectedItem
                        Case "Book-Series", "DCI", "Communication", "Journal", "Magazine-Newspaper",
                                "Meeting", "Report", "Thesis"
                            saMenus = {"Author", "Editor", "Translator", "Delete"}
                        Case "General", "Non-Traditional-Reference", "InPress", "Unpublished"
                            saMenus = {"Author", "Editor", "Translator", "Inventor", "Assignee", "Delete"}
                        Case "Patent"
                            saMenus = {"Translator", "Inventor", "Assignee", "Delete"}
                        Case ""
                            MsgBox("Select Element first")
                            saMenus = Nothing
                    End Select

                Case "EXTERNAL_LINK"
                    saMenus = {"URI", "FTP", "EMAIL", "Other", "Delete"}

                Case "G_NAME"
                    Select Case cmbElements.SelectedItem
                        Case "Book-Series", "DCI", "Communication", "Journal", "Magazine-Newspaper",
                                "Meeting", "Report", "Thesis"
                            saMenus = {"Author", "Editor", "Translator", "Delete"}
                        Case "General", "Non-Traditional-Reference", "InPress", "Unpublished"
                            saMenus = {"Author", "Editor", "Translator", "Inventor", "Assignee", "Delete"}
                        Case "Patent"
                            saMenus = {"Translator", "Inventor", "Assignee", "Delete"}
                    End Select

                Case "G_EXTERNAL_LINK"
                    saMenus = {"URI", "FTP", "EMAIL", "Other", "Delete"}

            End Select
            If saMenus Is Nothing Then
                lvTags.ContextMenuStrip = Nothing
            ElseIf saMenus(0) = "" Then
                lvTags.ContextMenuStrip = cmsTag
            Else
                For i As Integer = 0 To saMenus.Count - 1
                    Dim menu1 As New ToolStripMenuItem() With {.Text = saMenus(i), .Name = "mnu" & saMenus(i)}
                    AddHandler menu1.Click, AddressOf mnuItem_Clicked
                    cmsContext.Items.Add(menu1)
                Next
                lvTags.ContextMenuStrip = cmsContext
            End If
        End If
    End Sub

    Private Sub mnuItem_Clicked(sender As Object, e As EventArgs)
        Dim item As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
        If lvTags.SelectedItems.Count = 0 Then
            tsslStatus.Text = "No tags selected"
            Exit Sub
        End If
        If item IsNot Nothing Then
            If item.Name = "mnuDelete" Then
                If MsgBox("Do you want to remove selected Tag", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    lvTags.SelectedItems(0).Remove()
                End If
            ElseIf item.Text = "DOI" Then
                Dim strText As String = lvTags.SelectedItems(0).SubItems(2).Text.Replace(" ", "").Trim
                If Not isValidDOI(strText) Then
                    MsgBox("Invalid DOI Format")
                    Exit Sub
                Else
                    lvTags.SelectedItems(0).SubItems(1).Text = item.Text
                End If
            Else
                lvTags.SelectedItems(0).SubItems(1).Text = Replace(item.Name, "mnu", String.Empty)
            End If
        End If
    End Sub

    Private Sub ShowErrorLogsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowErrorLogsToolStripMenuItem.Click
        ufLogs.lblLogs.Text = sbLogs.ToString
        ufLogs.Show()
        ufLogs.StartPosition = FormStartPosition.CenterParent
    End Sub

    Private Function GetMethodName() As String
        Dim st As StackTrace = New StackTrace()
        Dim sf As StackFrame = st.GetFrame(1)
        Dim mb As System.Reflection.MethodBase = sf.GetMethod()
        Return mb.Name
    End Function

    Private Sub chkOnTop_CheckStateChanged(sender As Object, e As EventArgs)
        If CType(sender, CheckBox).CheckState = CheckState.Checked Then
            CType(sender.parent, Form).TopMost = True
        Else
            CType(sender.parent, Form).TopMost = False
        End If
    End Sub

    Private Sub SpecialCharToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecialCharToolStripMenuItem.Click
        frmCharMap.Show()
    End Sub

    Public Shared Sub insertText(strText As String)
        'Dim rtbText As RichTextBox, sel As Integer
        frmFMain.rtbOCR.SelectedText = strText
        'rtbText = frmFMain.rtbOCR
        'sel = rtbText.SelectionStart
        'rtbText.Text = rtbText.Text.Insert(sel, strText)
        'rtbText.SelectionStart = sel + strText.Length

    End Sub

    Public Shared Function CharsInTag(strTag As String) As Integer
        If Not strTag.Trim.StartsWith("<") Then
            Return strTag.Trim.Length
        End If

        Dim tagStart As Integer
        Dim tagEnd As Integer
        Dim strText As String = strTag

        If strTag.Contains("<![CDATA[") Then
            strTag.Replace("<![CDATA[", "")
            strTag.Replace("]]>", "")
        End If
        tagStart = strTag.IndexOf("<")
        tagEnd = strTag.IndexOf(">")
        If tagStart < tagEnd Then
            Do Until ((tagStart = -1) Or (tagEnd = -1))
                strText = strText.Replace(strText.Substring(tagStart, tagEnd - tagStart + 1), "")
                tagStart = strText.IndexOf("<")
                tagEnd = strText.IndexOf(">")
            Loop
        End If
        Return strText.Trim.Length
    End Function

    Public Shared Function TextInTag(strTag As String) As String
        If Not strTag.Trim.StartsWith("<") Then
            Return strTag.Trim
        End If

        Dim tagStart As Integer
        Dim tagEnd As Integer
        Dim strText As String = strTag

        If strText.Contains("<![CDATA[") Then
            strText = strText.Replace("<![CDATA[", "")
            strText = strText.Replace("]]>", "")
        End If
        tagStart = strTag.IndexOf("<")
        tagEnd = strTag.IndexOf(">")
        If tagStart < tagEnd Then
            Do Until ((tagStart = -1) Or (tagEnd = -1))
                strText = strText.Replace(strText.Substring(tagStart, tagEnd - tagStart + 1), "")
                tagStart = strText.IndexOf("<")
                tagEnd = strText.IndexOf(">")
            Loop
        End If
        Return strText.Trim
    End Function

    Private Sub process_name(strType As String, strText As String)
        If cmbElements.SelectedItem = "Patent" Then strType = "Assignee"

        If strText.ToLower.Contains(", et al.") Then
            If strText.Contains(";") Then
                strText = Replace(strText, ", et al.", "; et al", Compare:=CompareMethod.Text)
            Else
                strText = Replace(strText, ", et al.", ", et al", Compare:=CompareMethod.Text)
            End If
        ElseIf strText.ToLower.Contains(", Anonymous.") Then
            If strText.Contains(";") Then
                strText = Replace(strText, ", Anonymous.", "; Anonymous", Compare:=CompareMethod.Text)
            Else
                strText = Replace(strText, ", Anonymous.", ", Anonymous", Compare:=CompareMethod.Text)
            End If
        End If
        If strText.Contains(";") Then
            If strText.ToLower.Contains(" and ") Then
                strText = Replace(strText, " and ", ";", Compare:=CompareMethod.Text)
            End If
            Dim names() As String = Split(strText, ";")
            For Each fName In names
                fName = fName.Trim
                If fName = String.Empty Then Continue For
                If fName.ToLower.Contains(", et al.") Then
                    AddTagToList("ET_AL_IND", String.Empty, "Y")
                ElseIf fName.ToLower.Contains(", Anonymous.") Then
                    AddTagToList("ANONYMOUS_IND", String.Empty, "Y")
                ElseIf fName.ToLower.Contains(" of ") Or fName.ToLower.Contains(" the ") Then
                    AddTagToList("NAME", strType, "<SURNAME>" & fName.Trim & "</SURNAME>")
                Else
                    Dim nms() As String = fName.Split(",")
                    Dim sname As String = nms.First ' Split(fName, ",").First.Trim
                    Dim gname As String = "" '= fName.Replace(sname, String.Empty).Trim
                    'gname = gname.Replace(",", String.Empty).Trim
                    'sname = System.Text.RegularExpressions.Regex.Replace(sname, "[^a-zA-Z ]+", "")
                    'gname = System.Text.RegularExpressions.Regex.Replace(gname, "[^a-zA-Z ]+", "").Trim
                    If gname <> "" Then
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME><GIVEN_NAMES>" & gname.Trim & "</GIVEN_NAMES>"
                    Else
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME>"
                    End If
                    AddTagToList("NAME", strType, strText)
                End If
            Next
        ElseIf strText.Contains(". ,") Then
            If strText.ToLower.Contains(" and ") Then
                strText = Replace(strText, " and ", ". ,", Compare:=CompareMethod.Text)
            ElseIf strText.ToLower.Contains(" & ") Then
                strText = Replace(strText, " & ", ". ,", Compare:=CompareMethod.Text)
            End If
            Dim names() As String = Split(strText, ". ,")
            For Each fName In names
                fName = fName.Trim
                If fName = String.Empty Then Continue For
                'If fName.ToLower.Contains(", et al.") Then
                If fName = "et al" Then
                    AddTagToList("ET_AL_IND", String.Empty, "Y")
                ElseIf fName = "Anonymous" Then
                    AddTagToList("ANONYMOUS_IND", String.Empty, "Y")
                ElseIf fName.ToLower.Contains(" of ") Or fName.ToLower.Contains(" the ") Then
                    AddTagToList("NAME", strType, "<SURNAME>" & fName.Trim & "</SURNAME>")
                Else
                    Dim nms() As String
                    If fName.Contains(",") Then
                        nms = fName.Split(",")
                    Else
                        nms = fName.Split(" ")
                    End If
                    Dim sname As String = nms.First 'Split(fName, ",").First.Trim
                    Dim gname As String = "" '= fName.Replace(sname, String.Empty).Trim
                    'gname = gname.Replace(",", String.Empty).Trim
                    For i As Integer = 1 To nms.Count - 1
                        gname = gname & nms(i) & " "
                    Next
                    'sname = System.Text.RegularExpressions.Regex.Replace(sname, "[^a-zA-Z ]+", "")
                    'gname = System.Text.RegularExpressions.Regex.Replace(gname, "[^a-zA-Z ]+", "").Trim
                    If gname <> "" Then
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME><GIVEN_NAMES>" & gname.Trim & "</GIVEN_NAMES>"
                    Else
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME>"
                    End If
                    AddTagToList("NAME", strType, strText)
                End If
            Next
        Else
            If strText.ToLower.Contains(" and ") Then
                strText = Replace(strText, " and ", ",", Compare:=CompareMethod.Text)
            End If
            Dim names() As String = Split(strText, ",")
            For Each fName In names
                fName = fName.Trim
                If fName = String.Empty Then Continue For
                If fName.ToLower.Contains(", et al.") Then
                    AddTagToList("ET_AL_IND", String.Empty, "Y")
                ElseIf fName.ToLower.Contains(", Anonymous.") Then
                    AddTagToList("ANONYMOUS_IND", String.Empty, "Y")
                ElseIf fName.ToLower.Contains(" of ") Or fName.ToLower.Contains(" the ") Then
                    AddTagToList("NAME", strType, "<SURNAME>" & fName.Trim & "</SURNAME>")
                Else
                    Dim nms() As String = fName.Split(" ")
                    Dim sname As String = nms.First 'Split(fName, " ").First.Trim
                    Dim gname As String = "" '= fName.Replace(sname, String.Empty).Trim
                    For i As Integer = 1 To nms.Count - 1
                        gname = gname & nms(i) & " "
                    Next
                    'sname = System.Text.RegularExpressions.Regex.Replace(sname, "[^a-zA-Z ]+", "")
                    'gname = System.Text.RegularExpressions.Regex.Replace(gname, "[^a-zA-Z ]+", "").Trim
                    If gname <> "" Then
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME><GIVEN_NAMES>" & gname.Trim & "</GIVEN_NAMES>"
                    Else
                        strText = "<SURNAME>" & sname.Trim & "</SURNAME>"
                    End If
                    AddTagToList("NAME", strType, strText)
                End If
            Next
        End If
    End Sub

    Private Function fnValidateTags() As Integer
        If lblCurRef.Text.Trim = lblTotRef.Text.Trim Then
            Return -1
        End If
        Dim bValid As Boolean = True, strXPath As String
        If lvTags.Items.Count = 0 Then
            If cmbElements.SelectedItem <> "Non-Traditional-Reference" Then
                MsgBox("Tags list empty.")
                Return -1
            End If
        End If
        For Each item As ListViewItem In lvTags.Items
            Dim xeTemp As XmlElement = Nothing
            strXPath = vbNullString
            Try
                strXPath = "Elements/" & cmbElements.SelectedItem.ToString
                strXPath = strXPath & "/" & item.SubItems(0).Text
                If item.SubItems(1).Text <> "" Then strXPath = strXPath & "/" & item.SubItems(1).Text
                xeTemp = xdElements.SelectSingleNode(strXPath)
            Catch ex As Exception
            End Try
            If xeTemp Is Nothing Then
                bValid = False
                item.ForeColor = Color.Red
            Else
                item.ForeColor = Color.Black
            End If
            If item.SubItems(2).Text.Contains("></") Then
                If ((item.SubItems(2).Text.Contains("</YEAR></DATE>")) Or (item.SubItems(2).Text.Contains("</DATE_STRING></DATE>"))) Then 'Or (item.SubItems(2).Text.Contains("<GIVEN_NAMES></GIVEN_NAMES>"))) Then
                    If ((item.SubItems(2).Text.Contains("<YEAR></YEAR>") Or (item.SubItems(2).Text.Contains("<DATE_STRING></DATE_STRING>")))) Then
                        bValid = False
                        item.ForeColor = Color.Red
                    End If
                Else
                    bValid = False
                    item.ForeColor = Color.Red
                End If
            End If
        Next

        For i As Integer = 0 To lvTags.Items.Count - 2
            For j As Integer = i + 1 To lvTags.Items.Count - 1
                If lvTags.Items(i).SubItems(0).Text = lvTags.Items(j).SubItems(0).Text Then
                    If lvTags.Items(i).SubItems(1).Text = lvTags.Items(j).SubItems(1).Text Then
                        If lvTags.Items(i).SubItems(2).Text = lvTags.Items(j).SubItems(2).Text Then
                            bValid = False
                            lvTags.Items(j).ForeColor = Color.Red
                        End If
                    End If
                End If
            Next
        Next
        If bValid Then
            If fnCheckReqTag() Then
                Return 1
            Else
                MsgBox("Required Tag Missing. Check for SPT, Conf Name or Pub ID tags")
                Return -1
            End If
            Return 1
        Else
            Return -1
        End If
    End Function

    Private Function fnCheckReqTag() As Boolean
        Dim element As String = cmbElements.Text, bValid As Boolean = False
        Dim TextLine As String

        If (element = "Journal") Or (element = "Book-Series") Or (element = "Meeting") Then
            If File.Exists(strPolicyFile) = True Then
                Dim objReader As New StreamReader(strPolicyFile)

                Do While (objReader.Peek() <> -1)
                    TextLine = objReader.ReadLine().ToLower.Trim
                    If rtbOCR.Text.ToLower.Contains(TextLine) Then
                        Return True
                    End If
                Loop
            End If
        End If
        Select Case element
            Case "Journal", "Book-Series"
                If rtbOCR.Text.Contains("") Then

                End If
                For Each item As ListViewItem In lvTags.Items
                    If item.SubItems(0).Text = "SOURCE_PUB_TITLE" Then
                        bValid = True
                    End If
                Next
            Case "Meeting"
                For Each item As ListViewItem In lvTags.Items
                    If (item.SubItems(0).Text = "SOURCE_PUB_TITLE") Or (item.SubItems(0).Text = "CONF_NAME") Then
                        bValid = True
                    End If
                Next
            Case "Report", "Magazine-Newspaper"
                For Each item As ListViewItem In lvTags.Items
                    If item.SubItems(0).Text = "SOURCE_PUB_TITLE" Then
                        bValid = True
                    End If
                Next
            Case "Patent"
                For Each item As ListViewItem In lvTags.Items
                    If item.SubItems(0).Text = "PUB_ID" Then
                        bValid = True
                    End If
                Next
            Case Else
                bValid = True
        End Select
        If bValid Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub AddTagToList(p1 As String, p2 As String, p3 As String)
        Dim lvi As New ListViewItem(p1)
        lvi.SubItems.Add(p2)
        lvi.SubItems.Add(p3)
        lvTags.Items.Add(lvi)
        lvi.EnsureVisible()
        rtbOCR.Focus()
    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        If MsgBox("Do you want to reject this item?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = vbYes Then

            Dim lstInTif As IEnumerable(Of String) = Directory.EnumerateFiles(strRootPath & strInDir, Replace(Split(strInFile, "\").Last, ".XML", "*.TIF", Compare:=CompareMethod.Text), SearchOption.TopDirectoryOnly)

            For Each inTif As String In lstInTif
                File.Copy(inTif, Replace(inTif, strInDir, strQueryDir), True)
                File.Copy(inTif, Replace(inTif, strInDir, strQueryDir), True)
                Try
                    File.Delete(inTif)
                Catch ex As Exception
                    sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
                End Try
            Next inTif
            Try
                File.Copy(Replace(strInFile, strInDir, strCurDir), Replace(strInFile, strInDir, strQueryDir), True)
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
            End Try

            Try
                File.Delete(Replace(strInFile, strInDir, strCurDir))
            Catch ex As Exception
                sbLogs.AppendLine(Erl() & " :: " & GetMethodName() & " :: " & ex.Message)
            End Try
            subClearFields()

            swCurProd.Stop()
            swCurSusp.Stop()

            LogEntry.Session_Production = swCurProd.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_Suspend = swCurSusp.Elapsed.ToString("hh\:mm\:ss")
            LogEntry.Session_End_Date = Format(Today, "dd-MM-yyyy")
            LogEntry.Session_End_Time = Now.ToShortTimeString
            LogEntry.Status = "QR"
            UpdateLogEntry(strLogPath & frmLogin.UserName & ".XML", LogEntry)

            lvRefs.Items.Clear()
            strInFile = String.Empty
            xdDoc = Nothing
            xnlCitations = Nothing

            tabMain.SelectedTab = tpFCR
            If MsgBox("Do you want to load new item?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                strInFile = GetXMLFile(strRootPath & strPriorDir)

                If strInFile = "" Then
                    bIsPrior = False
                    strInFile = GetXMLFile(strRootPath & strInDir)
                End If
                If strInFile = "" Then
                    MsgBox("No files for input", MsgBoxStyle.Information)
                Else
                    bIsPrior = False
                    LoadFile(strInFile)
                End If
            End If
        End If
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        Dim sbMsg As New System.Text.StringBuilder

        If xdDoc IsNot Nothing Then
            If xdDoc.BaseURI <> "" Then
                frmReport.lblAccession.Text = xdDoc.SelectSingleNode("//ID_ACCESSION").InnerText
                frmReport.lblItemNo.Text = xdDoc.SelectSingleNode("//ITEM").Attributes("ITEMNO").InnerText
            End If
        Else
            frmReport.lblAccession.Text = ""
            frmReport.lblItemNo.Text = ""
        End If
        frmReport.lblCompItems.Text = iCompItems
        frmReport.lblCompRefs.Text = iCompRefs
        frmReport.Show()
    End Sub

    Private Sub MakeLogFile(filename As String)
        If Not File.Exists(filename) Then
            Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 3
            writer.WriteStartElement("Logs")
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
        Else
            Dim tdoc As New XmlDocument
            Try
                tdoc.Load(filename)
            Catch ex As Exception
                Dim tsb As New Text.StringBuilder
                tsb.AppendLine("DB file is either corrupted or invalid")
                tsb.AppendLine("It will be backed up and New DB file will be created")
                MsgBox(tsb.ToString)
                Try
                    My.Computer.FileSystem.RenameFile(filename, Path.GetFileName(filename) & ".backup")
                Catch ex1 As Exception
                    Dim i As Integer = 1
                    While (File.Exists(Path.GetFileName(filename) & ".backup" & i))
                        i = i + 1
                    End While
                    My.Computer.FileSystem.RenameFile(filename, Path.GetFileName(filename) & ".backup" & i)
                End Try

                Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 3
                writer.WriteStartElement("Logs")
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()
            End Try
        End If
        tsslStatus.Text = "Created"
    End Sub

    Private Sub InitLog(ByRef LE As Log)
        LE.Session_Start_Date = ""
        LE.Session_Start_Time = ""
        LE.Session_End_Date = ""
        LE.Session_End_Time = ""
        LE.Session_Production = ""
        LE.Session_Suspend = ""
        LE.Item_Accession = ""
        LE.Item_Itemno = ""
        LE.Status = ""
        LE.references_Open = ""
        'LE.references_Close = ""
        'LE.references_Inserted = ""
        'LE.references_Deleted = ""
    End Sub

    Private Sub MakeLogEntry(ByVal filename As String)
        Dim xdDoc As New XmlDocument
        Dim nod1 As XmlNode, nod2 As XmlNode, attr As XmlAttribute

        Try
            xdDoc.Load(strLogPath & frmLogin.UserName & ".XML")
        Catch ex As Exception
            MsgBox("Could not open DB file")
            Exit Sub
        End Try

        nod1 = xdDoc.DocumentElement
        nod2 = xdDoc.CreateElement("Log")
        attr = xdDoc.CreateAttribute("seq")
        attr.Value = xdDoc.DocumentElement.SelectNodes("Log").Count + 1
        nod2.Attributes.Append(attr)
        nod1.AppendChild(nod2)
        nod1 = nod2

        nod2 = xdDoc.CreateElement("System")
        attr = xdDoc.CreateAttribute("type")
        attr.Value = LogEntry.System_Type
        nod2.Attributes.Append(attr)
        nod2.InnerText = LogEntry.System
        nod1.AppendChild(nod2)

        nod2 = xdDoc.CreateElement("Session")
        nod1.AppendChild(nod2)
        nod1 = nod2

        nod2 = xdDoc.CreateElement("Start_Time")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Date")
        nod2.InnerText = LogEntry.Session_Start_Date
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Time")
        nod2.InnerText = LogEntry.Session_Start_Time
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("End_Time")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Date")
        nod2.InnerText = LogEntry.Session_End_Date
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Time")
        nod2.InnerText = LogEntry.Session_End_Time
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Production")
        nod2.InnerText = LogEntry.Session_Production
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("Suspend")
        nod2.InnerText = LogEntry.Session_Suspend
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Item")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Accession")
        nod2.InnerText = LogEntry.Item_Accession
        nod1.AppendChild(nod2)
        nod2 = xdDoc.CreateElement("ItemNo")
        nod2.InnerText = LogEntry.Item_Itemno
        nod1.AppendChild(nod2)

        nod1 = nod1.ParentNode
        nod2 = xdDoc.CreateElement("Status")
        nod2.InnerText = LogEntry.Status
        nod1.AppendChild(nod2)

        nod2 = xdDoc.CreateElement("references")
        nod1.AppendChild(nod2)
        nod1 = nod2
        nod2 = xdDoc.CreateElement("Open")
        nod2.InnerText = LogEntry.references_Open
        nod1.AppendChild(nod2)
        'nod2 = xdDoc.CreateElement("Close")
        'nod2.InnerText = LogEntry.references_Close
        'nod1.AppendChild(nod2)
        'nod2 = xdDoc.CreateElement("Inserted")
        'nod2.InnerText = LogEntry.references_Inserted
        'nod1.AppendChild(nod2)
        'nod2 = xdDoc.CreateElement("Deleted")
        'nod2.InnerText = LogEntry.references_Deleted
        'nod1.AppendChild(nod2)

        xdDoc.Save(filename)
    End Sub

    Private Sub UpdateLogEntry(ByVal filename As String, LE As Log)
        Dim xddoc As New XmlDocument, lChild As XmlNode
        Try
            xddoc.Load(filename)
        Catch ex As Exception
            MsgBox("Could not open DB file")
            Exit Sub
        End Try

        lChild = xddoc.DocumentElement.LastChild
        If lChild Is Nothing Then
            MsgBox("Could not update. Check DB file")
            Exit Sub
        End If
        lChild.SelectSingleNode("System").InnerText = LE.System
        lChild.SelectSingleNode("System").Attributes("type").Value = LE.System_Type
        lChild.SelectSingleNode("Session/Start_Time/Date").InnerText = LE.Session_Start_Date
        lChild.SelectSingleNode("Session/Start_Time/Time").InnerText = LE.Session_Start_Time
        lChild.SelectSingleNode("Session/End_Time/Date").InnerText = LE.Session_End_Date
        lChild.SelectSingleNode("Session/End_Time/Time").InnerText = LE.Session_End_Time
        lChild.SelectSingleNode("Session/Production").InnerText = LE.Session_Production
        lChild.SelectSingleNode("Session/Suspend").InnerText = LE.Session_Suspend
        lChild.SelectSingleNode("Item/Accession").InnerText = LE.Item_Accession
        lChild.SelectSingleNode("Item/ItemNo").InnerText = LE.Item_Itemno
        lChild.SelectSingleNode("Status").InnerText = LE.Status
        lChild.SelectSingleNode("references/Open").InnerText = LE.references_Open
        'lChild.SelectSingleNode("references/Close").InnerText = LE.references_Close
        'lChild.SelectSingleNode("references/Inserted").InnerText = LE.references_Inserted
        'lChild.SelectSingleNode("references/Deleted").InnerText = LE.references_Deleted
        xddoc.Save(filename)
    End Sub

    Private Sub lvTags_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvTags.SelectedIndexChanged
        rtbOCR.Text = rtbOCR.Text
        If lvTags.SelectedItems.Count > 0 Then
            If lvTags.SelectedItems(0).SubItems.Count >= 3 Then
                ShowInBlurrb(lvTags.SelectedItems(0).SubItems(2).Text)
            End If
        End If
    End Sub

    Private Sub ShowInBlurrb(fulltag As String)
        Dim tagtext As String = TextInTag(fulltag)
        Dim blurrb As String = rtbOCR.Text
        Dim iTPos As Integer, iBPos As Integer, curPos As Integer
        Dim bumbTagText As String = tagtext.Replace(" ", String.Empty)
        Dim selSt As Integer, selEnd As Integer

        If tagtext = "" Then Exit Sub
        selSt = -1
        selEnd = -1
        curPos = 0
        blurrb = blurrb.ToLower
        bumbTagText = bumbTagText.ToLower

        For iBPos = 0 To blurrb.Length - 1
            If Not System.Text.RegularExpressions.Regex.Match(blurrb(iBPos), "[0-9a-zA-Z]").Success Then Continue For

            For iTPos = curPos To bumbTagText.Length - 1
                If System.Text.RegularExpressions.Regex.Match(bumbTagText(iTPos), "[0-9a-zA-Z]").Success Then Exit For
            Next

            If iTPos >= bumbTagText.Length Then
                rtbOCR.SelectAll()
                rtbOCR.SelectionBackColor = Color.White
                rtbOCR.SelectionColor = Color.Black

                rtbOCR.SelectionStart = selSt
                rtbOCR.SelectionLength = selEnd - selSt + 1
                rtbOCR.SelectionBackColor = Color.DarkBlue
                rtbOCR.SelectionColor = Color.White
                Exit Sub
            End If

            If blurrb(iBPos) = bumbTagText(iTPos) Then
                If selSt = -1 Then
                    selSt = iBPos
                End If
                selEnd = iBPos
                curPos = iTPos + 1
            Else
                selSt = -1
                selEnd = -1
                curPos = 0
                If blurrb(iBPos) = bumbTagText(0) Then
                    selSt = iBPos
                    selEnd = iBPos
                    iTPos = 0
                    curPos = iTPos + 1
                End If
            End If
        Next
        If selSt <> -1 And selEnd <> -1 Then
            rtbOCR.SelectAll()
            rtbOCR.SelectionBackColor = Color.White
            rtbOCR.SelectionColor = Color.Black
            rtbOCR.SelectionStart = selSt
            rtbOCR.SelectionLength = selEnd - selSt + 1
            rtbOCR.SelectionBackColor = Color.DarkBlue
            rtbOCR.SelectionColor = Color.White
        End If
    End Sub

    Private Function GetValidRefsCount() As String
        If xnlCitations Is Nothing Then Return ""
        If xnlCitations.Count = 0 Then Return "0"
        Dim count As Integer = 0
        For Each nod As XmlNode In xnlCitations
            Dim attr As XmlAttribute = nod.Attributes("D")
            If attr IsNot Nothing Then Continue For
            Dim tnod As XmlNode = nod.SelectSingleNode("UT_CODE")
            If tnod IsNot Nothing Then Continue For
            count = count + 1
        Next
        Return count.ToString
    End Function

    Private Sub SuspendToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SuspendToolStripMenuItem.Click
        Me.Enabled = False
        'frmLogin.txtUserName.Text = lblUName.Text
        'frmLogin.txtUserName.Enabled = False
        'focusToBox()
        frmLogin.Show(Me)
    End Sub

    Private Sub frmFMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.Modifiers = Keys.Alt Then
            e.SuppressKeyPress = True
            e.Handled = True
        End If
    End Sub

    Private Sub ReadCIInfo(fromCINode As XmlNode, ByRef toCIStruct As CI_INFO)
        If fromCINode Is Nothing Then
            Exit Sub
        End If
        Try
            toCIStruct.Author = fromCINode.SelectSingleNode("CI_JOURNAL/CI_AUTHOR").InnerText.Trim
        Catch ex As Exception
        End Try
        Try
            toCIStruct.Volume = fromCINode.SelectSingleNode("CI_JOURNAL/CI_VOLUME").InnerText.Trim
        Catch ex As Exception
        End Try
        Try
            toCIStruct.Page = fromCINode.SelectSingleNode("CI_JOURNAL/CI_PAGE").InnerText.Trim
        Catch ex As Exception
        End Try
        Try
            toCIStruct.Year = fromCINode.SelectSingleNode("CI_JOURNAL/CI_YEAR").InnerText.Trim
        Catch ex As Exception
        End Try
        Try
            toCIStruct.Title = fromCINode.ParentNode.SelectSingleNode("CI_CAPTURE").SelectSingleNode("CI_CAPTURE_TITLE").InnerText.Trim
        Catch ex As Exception
        End Try

        Dim xnlArtn As XmlNodeList
        xnlArtn = fromCINode.ParentNode.SelectNodes("RI_CITATIONIDENTIFIER")

        Try
            toCIStruct.ARTN1 = xnlArtn(0).InnerText
            toCIStruct.ARTNType1 = xnlArtn(0).Attributes("seq").Value
        Catch ex As Exception
            toCIStruct.ARTN1 = ""
            toCIStruct.ARTNType1 = ""
        End Try
        Try
            toCIStruct.ARTN2 = xnlArtn(1).InnerText
            toCIStruct.ARTNType2 = xnlArtn(1).Attributes("seq").Value
        Catch ex As Exception
            toCIStruct.ARTN2 = ""
            toCIStruct.ARTNType2 = ""
        End Try
        Try
            toCIStruct.ARTN3 = xnlArtn(2).InnerText
            toCIStruct.ARTNType3 = xnlArtn(2).Attributes("seq").Value
        Catch ex As Exception
            toCIStruct.ARTN3 = ""
            toCIStruct.ARTNType3 = ""
        End Try
        Try
            toCIStruct.ARTN4 = xnlArtn(3).InnerText
            toCIStruct.ARTNType4 = xnlArtn(3).Attributes("seq").Value
        Catch ex As Exception
            toCIStruct.ARTN4 = ""
            toCIStruct.ARTNType4 = ""
        End Try

    End Sub

    Private Function ValidateFCIInfo() As Boolean
        If lblCurRef.Text.Trim = lblTotRef.Text.Trim Then
            Return True
        End If
        Dim FCI As XmlNode, i As Integer = 0, UTC As XmlNode, DEL As XmlNode
        For Each cit As XmlNode In xnlCitations
            DEL = Nothing
            Try
                DEL = cit.Attributes("D")
            Catch ex As Exception
            End Try
            If DEL IsNot Nothing Then
                Continue For
            End If

            UTC = Nothing
            Try
                UTC = cit.SelectSingleNode("UT_CODE")
            Catch ex As Exception
            End Try
            If UTC IsNot Nothing Then
                Continue For
            End If

            FCI = Nothing
            Try
                FCI = cit.SelectSingleNode("FULL_CI_INFO")
            Catch ex As Exception
            End Try

            If FCI Is Nothing Then
                subLoadRefs(xnlCitations, lstRefNum(i), "Next")
                MsgBox("Complete this Citation")
                Return False
            End If

            If String.IsNullOrWhiteSpace(FCI.InnerText) Then
                subLoadRefs(xnlCitations, lstRefNum(i), "Next")
                MsgBox("Complete this Citation")
                Return False
            End If

            i += 1
        Next
        Return True
    End Function

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim sbAbout As New Text.StringBuilder

        sbAbout.AppendLine(ProductName)
        sbAbout.AppendLine("Version " & ProductVersion)
        sbAbout.AppendLine("Build : 2001092235")
        MsgBox(sbAbout.ToString)
    End Sub

    Private Sub ThemeColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThemeColorToolStripMenuItem.Click
        If ColorDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.BackColor = ColorDialog.Color
            tpFCR.BackColor = ColorDialog.Color
            tpQuery.BackColor = ColorDialog.Color
            My.Settings.MainBackColor = Me.BackColor
            My.Settings.Save()
        End If
    End Sub

    Private Sub ReorderTags()
        Dim TagPos As New Dictionary(Of Integer, Integer), XtraItems As Integer = 0
        'Debug.Print("Start time : " & Now.ToString("hh:mm:ss"))
        For Each item As ListViewItem In lvTags.Items
            Dim add As Integer = 0
            Select Case item.Text
                Case "SOURCE_PUB_TITLE"
AddSPT:
                    Try
                        TagPos.Add(0 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddSPT
                    End Try
                Case "SOURCE_ID"
AddSrcID:
                    Try
                        TagPos.Add(5 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddSrcID
                    End Try
                Case "PUBLISHER_NAME"
AddPubName:
                    Try
                        TagPos.Add(10 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddPubName
                    End Try
                Case "PUBLISHER_LOC"
AddPubLoc:
                    Try
                        TagPos.Add(15 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddPubLoc
                    End Try
                Case "ITEM_TITLE"
AddItemTitle:
                    Try
                        TagPos.Add(20 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddItemTitle
                    End Try
                Case "EDITION"
AddEdition:
                    Try
                        TagPos.Add(25 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddEdition
                    End Try
                Case "VOLUME_ID"
AddVolID:
                    Try
                        TagPos.Add(30 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddVolID
                    End Try
                Case "ISSUE_ID"
AddIsID:
                    Try
                        TagPos.Add(35 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddIsID
                    End Try
                Case "ISSUE_TITLE"
AddIsTitle:
                    Try
                        TagPos.Add(40 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddIsTitle
                    End Try
                Case "SUPPLEMENT"
AddSupp:
                    Try
                        TagPos.Add(45 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddSupp
                    End Try
                Case "PUB_DATE"
AddPubDate:
                    Try
                        TagPos.Add(50 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddPubDate
                    End Try
                Case "PAGE_RANGE"
AddPageRange:
                    Try
                        TagPos.Add(55 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddPageRange
                    End Try
                Case "SIZE"
AddSize:
                    Try
                        TagPos.Add(60 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddSize
                    End Try
                Case "EXTERNAL_LINK"
AddExtLink:
                    Try
                        TagPos.Add(65 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddExtLink
                    End Try
                Case "PUB_ID"
AddPubID:
                    Try
                        TagPos.Add(70 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddPubID
                    End Try
                Case "COLLAB_GRP_NAME"
AddCGN:
                    Try
                        TagPos.Add(75 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddCGN
                    End Try
                Case "INSTITUTION"
AddIns:
                    Try
                        TagPos.Add(80 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddIns
                    End Try
                Case "G_NAME"
AddGName:
                    Try
                        TagPos.Add(85 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddGName
                    End Try
                Case "G_EXTERNAL_LINK"
AddGExtLink:
                    Try
                        TagPos.Add(90 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddGExtLink
                    End Try
                Case "NAME"
AddName:
                    Try
                        TagPos.Add(95 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 299 Then GoTo AddName
                    End Try
                Case "CONF_NAME"
AddConfName:
                    Try
                        TagPos.Add(395 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddConfName
                    End Try
                Case "CONF_DATE"
AddConfDate:
                    Try
                        TagPos.Add(400 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddConfDate
                    End Try
                Case "CONF_LOC"
AddConfLoc:
                    Try
                        TagPos.Add(405 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddConfLoc
                    End Try
                Case "CONF_SPONSOR"
AddConfSpon:
                    Try
                        TagPos.Add(410 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 4 Then GoTo AddConfSpon
                    End Try
                Case "MISC"
AddMisc:
                    Try
                        TagPos.Add(415 + add, lvTags.Items.IndexOf(item))
                    Catch ex As Exception
                        add += 1
                        If add <= 9 Then GoTo AddMisc
                    End Try
                Case Else
                    TagPos.Add(425 + XtraItems, lvTags.Items.IndexOf(item))
                    XtraItems += 1
            End Select
        Next
        'Debug.Print("Loop end time : " & Now.ToString("hh:mm:ss"))
        Dim ArrLVI(0 To TagPos.Count - 1) As ListViewItem
        Dim iKey As Integer, iValue As Integer, iCount As Integer
        Dim bResult As Boolean
        iKey = 0 : iCount = 0
        'Debug.Print("While Start time : " & Now.ToString("hh:mm:ss"))
        While True
            If iCount >= lvTags.Items.Count Then Exit While
            If iKey > 500 Then Exit While
            bResult = TagPos.TryGetValue(iKey, iValue)
            If bResult Then
                ArrLVI(iCount) = lvTags.Items(iValue)
                iCount += 1
            End If
            iKey += 1
        End While
        'Debug.Print("While end time : " & Now.ToString("hh:mm:ss"))
        lvTags.Items.Clear()
        Try
            lvTags.Items.AddRange(ArrLVI)
        Catch ex As Exception
            MsgBox("Tag index exceeds allowed limit(500)")
        End Try
        'Debug.Print("End time : " & Now.ToString("hh:mm:ss"))
    End Sub

End Class
