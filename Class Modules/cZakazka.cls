VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cZakazka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TClass
    pCislo As Long              'jedinèný èíselný identifikátor
    pStroj As Long              'èíslo stroja zákazky
    pTermin As Date             'termín spracovania zákazky
    pKvalita As String          'stupeò lepenia zákazky
    pMnozstvo As Double         'celkový poèet kusov zákazky
    pSirka As Double            'šírka zákazky v mm
    pDlzka As Double            'dåžka zákazky v mm
    pRyhovanie As String        'kód ryhovania zákazky
    pPriorita As Integer        'priorita zákazky
End Type

Private Type cJednotky
    Dlzka As String
End Type

Private this As TClass
Private oJednotky As cJednotky

Private Sub Class_Initialize()
    With oJednotky
        .Dlzka = "m"
    End With
End Sub

Public Property Get Cislo() As Long
    Cislo = this.pCislo
End Property

Public Property Let Cislo(ByVal CisloZakazky As Long)
    this.pCislo = CisloZakazky
End Property

Public Property Get Stroj() As Long
    Stroj = this.pStroj
End Property

Public Property Let Stroj(ByVal CisloStrojaZakazky As Long)
    this.pStroj = CisloStrojaZakazky
End Property

Public Property Get Termin() As Date
    Termin = this.pTermin
End Property

Public Property Let Termin(ByVal TerminZakazky As Date)
    this.pTermin = TerminZakazky
End Property

Public Property Get Kvalita() As String
    Kvalita = this.pKvalita
End Property

Public Property Let Kvalita(ByVal StupenLepeniaZakazky As String)
    this.pKvalita = StupenLepeniaZakazky
End Property

Public Property Get Mnozstvo() As Long
    Mnozstvo = this.pMnozstvo
End Property

Public Property Let Mnozstvo(ByVal CelkoveMnozstvoZakazky As Long)
    this.pMnozstvo = CelkoveMnozstvoZakazky
End Property

Public Property Get Sirka(Optional ByVal vJednotkach As String = "mm") As Double
    Sirka = WorksheetFunction.Convert(this.pSirka, oJednotky.Dlzka, vJednotkach)
End Property

Public Property Let Sirka(Optional ByVal vJednotkach As String = "mm", ByVal SirkaZakazky_v_mm As Double)
    this.pSirka = WorksheetFunction.Convert(SirkaZakazky_v_mm, vJednotkach, oJednotky.Dlzka)
End Property

Public Property Get Dlzka(Optional ByVal vJednotkach As String = "mm") As Double
    Dlzka = WorksheetFunction.Convert(this.pDlzka, oJednotky.Dlzka, vJednotkach)
End Property

Public Property Let Dlzka(Optional ByVal vJednotkach As String = "mm", ByVal DlzkaZakazky_v_mm As Double)
    this.pDlzka = WorksheetFunction.Convert(DlzkaZakazky_v_mm, vJednotkach, oJednotky.Dlzka)
End Property

Public Property Get Ryhovanie() As String
    Ryhovanie = this.pRyhovanie
End Property

Public Property Let Ryhovanie(ByVal TypRyhovania As String)
    this.pRyhovanie = TypRyhovania
End Property

Public Property Get Priorita() As Integer
    Priorita = this.pPriorita
End Property

Public Property Let Priorita(ByVal HodnotaPriority As Integer)
    this.pPriorita = HodnotaPriority
End Property

