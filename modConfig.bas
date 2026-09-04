Attribute VB_Name = "modConfig"

' modConfig ? central column maps + document configuration (DocConfig).
'   Single source of truth for log layouts so magic numbers disappear and a
'   layout change is a one-line edit. Used by the shared Save/Recall/Renumber
'   engine (Phase 0+) and by MedClaim (Phase 1).
'
'   PHASE 0: definitions only. Existing modules keep working unchanged; new
'   engine code and later phases consume these.
' ============================================================================
Option Explicit

' ---------------- InvoiceLog columns (1-based) ----------------
Public Const IL_NO As Long = 1          ' A InvNo
Public Const IL_DEPT As Long = 2        ' B Dept
Public Const IL_RECIP As Long = 3       ' C RecipientType
Public Const IL_DATE As Long = 4        ' D InvDate
Public Const IL_DUE As Long = 5         ' E DueDate
Public Const IL_CUST As Long = 6        ' F Bill-To / CustID
Public Const IL_PATIENT As Long = 7     ' G PatientName
Public Const IL_APPLIANCE As Long = 8   ' H ApplianceType
Public Const IL_SUBTOTAL As Long = 9    ' I  (H35 excl)
Public Const IL_DISC As Long = 10       ' J  (H34 disc)
Public Const IL_VAT As Long = 11        ' K  (H36 vat)
Public Const IL_TOTAL As Long = 12      ' L  (H37 total)
Public Const IL_PAID As Long = 14       ' N Paid
Public Const IL_BALANCE As Long = 15    ' O Balance
Public Const IL_STATUS As Long = 16     ' P Status
Public Const IL_SRCQUOTE As Long = 17   ' Q SourceQuoteNo
Public Const IL_CREATED As Long = 18    ' R CreatedAt
Public Const IL_MODIFIED As Long = 19   ' S ModifiedAt
Public Const IL_MEDAID As Long = 20     ' T MedAid
Public Const IL_MEDNO As Long = 21      ' U MedNo
Public Const IL_MAINMEM As Long = 22    ' V MainMember
Public Const IL_DOCTOR As Long = 23     ' W Doctor
Public Const IL_BHF As Long = 24        ' X BHF
Public Const IL_DISCPCT As Long = 25    ' Y Discount %
Public Const IL_DISCFIX As Long = 26    ' Z Discount Fixed
Public Const IL_NOTES As Long = 27      ' AA Notes (Phase 2)

' ---------------- QuoteLog columns ----------------
Public Const QL_NO As Long = 1          ' A QuoteNo
Public Const QL_DEPT As Long = 2        ' B Dept
Public Const QL_RECIP As Long = 3       ' C RecipientType
Public Const QL_DATE As Long = 4        ' D QuoteDate
Public Const QL_CUST As Long = 5        ' E Bill-To / CustID
Public Const QL_PATIENT As Long = 6     ' F PatientName
Public Const QL_APPLIANCE As Long = 7   ' G ApplianceType
Public Const QL_SUBTOTAL As Long = 8    ' H
Public Const QL_DISC As Long = 9        ' I
Public Const QL_VAT As Long = 10        ' J
Public Const QL_TOTAL As Long = 11      ' K
Public Const QL_KIND As Long = 12       ' L (K3 kind)
Public Const QL_STATUS As Long = 13     ' M Status
Public Const QL_CONVINV As Long = 14    ' N ConvertedInvNo
Public Const QL_CREATED As Long = 15    ' O CreatedAt
Public Const QL_MODIFIED As Long = 16   ' P ModifiedAt
Public Const QL_MEDAID As Long = 17     ' Q MedAid
Public Const QL_MEDNO As Long = 18      ' R MedNo
Public Const QL_MAINMEM As Long = 19    ' S MainMember
Public Const QL_DOCTOR As Long = 20     ' T Doctor
Public Const QL_BHF As Long = 21        ' U BHF
Public Const QL_DISCPCT As Long = 22    ' V Discount %
Public Const QL_DISCFIX As Long = 23    ' W Discount Fixed
Public Const QL_NOTES As Long = 24      ' X Notes (Phase 2)

' ---------------- MedAidLog columns (Phase 1) ----------------
Public Const ML_NO As Long = 1          ' A InvNo
Public Const ML_DEPT As Long = 2        ' B Dept
Public Const ML_RECIP As Long = 3       ' C RecipientType
Public Const ML_DATE As Long = 4        ' D InvDate
Public Const ML_DUE As Long = 5         ' E DueDate
Public Const ML_PATIENT As Long = 6     ' F PatientName
Public Const ML_APPLIANCE As Long = 7   ' G ApplianceType
Public Const ML_SUBTOTAL As Long = 8    ' H SubTotal
Public Const ML_DISC As Long = 9        ' I DiscountAmt
Public Const ML_VAT As Long = 10        ' J VatAmt
Public Const ML_TOTAL As Long = 11      ' K TotalDue
Public Const ML_PAID As Long = 12       ' L Paid
Public Const ML_CUST As Long = 13       ' M Bill To
Public Const ML_PATIENTID As Long = 14  ' N Patient ID
Public Const ML_CREATED As Long = 15    ' O CreatedAt
Public Const ML_MODIFIED As Long = 16   ' P ModifiedAt
Public Const ML_MEDAID As Long = 17     ' Q Medical Aid
Public Const ML_MEDNO As Long = 18      ' R Med no
Public Const ML_MAINMEM As Long = 19    ' S main member
Public Const ML_BHF As Long = 20        ' T BHF
Public Const ML_DOCTOR As Long = 21     ' U Treating Doctor
Public Const ML_DEPCODE As Long = 22    ' V Dependent Code
Public Const ML_DISCPCT As Long = 23    ' W Discount %
Public Const ML_DISCFIX As Long = 24    ' X Discount Fixed
Public Const ML_NOTES As Long = 25      ' Y Notes
Public Const ML_BALANCE As Long = 26    ' Z Balance
Public Const ML_STATUS As Long = 27     ' AA Status

' ---------------- Lines columns (Invoice/Quote/MedAid share this) ----------------
Public Const LN_DOCNO As Long = 1       ' A
Public Const LN_LINENO As Long = 2      ' B
Public Const LN_CODE As Long = 3        ' C
Public Const LN_ZCODE As Long = 4       ' D
Public Const LN_DESC As Long = 5        ' E
Public Const LN_QTY As Long = 6         ' F
Public Const LN_EXCL As Long = 7        ' G
Public Const LN_VAT As Long = 8         ' H
Public Const LN_INCL As Long = 9        ' I
Public Const LN_TOTAL As Long = 10      ' J

' ---------------- Payments columns ----------------
Public Const PY_ID As Long = 1          ' A PayID
Public Const PY_INV As Long = 2         ' B InvNo
Public Const PY_DATE As Long = 3        ' C Date
Public Const PY_AMT As Long = 4         ' D Amount
Public Const PY_METHOD As Long = 5      ' E Method
Public Const PY_REF As Long = 6         ' F Reference
Public Const PY_NOTES As Long = 7       ' G Notes

' ---------------- Settings counter cells ----------------
Public Const SET_QUOTE_WA As String = "B10"
Public Const SET_QUOTE_WD As String = "B11"
Public Const SET_INV_WA As String = "B12"
Public Const SET_INV_WD As String = "B13"
Public Const SET_CN As String = "B14"
Public Const SET_PAYMENT As String = "B18"
Public Const SET_MC As String = "B20"
Public Const SET_BACKUP_DIR As String = "B21"   ' optional backup path override

' ============================================================================
' DocConfig ? describes one document class so shared engine code can operate
' generically. docKind: "INV" | "QTE" | "CN" | "MC"
' ============================================================================
Public Type DocConfig
    docKind As String
    sheetName As String
    logName As String
    linesName As String
    prefix As String          ' e.g. "INV-", "Q-", "CN-", "MC-INV-"
    hasDeptSegment As Boolean  ' True for INV/QTE/MC, False for CN
    counterWA As String
    counterWD As String
    counterFlat As String      ' for CN/MC single-counter cases
    colNo As Long
    colDept As Long
    colDate As Long
    colTotal As Long
    colPaid As Long
    colBalance As Long
    colStatus As Long
End Type

' ---- Build a DocConfig for a given kind ----
Public Function GetDocConfig(ByVal docKind As String) As DocConfig
    Dim c As DocConfig
    Select Case UCase$(Trim$(docKind))
        Case "INV"
            c.docKind = "INV": c.sheetName = "Invoice"
            c.logName = "InvoiceLog": c.linesName = "InvoiceLines"
            c.prefix = "INV-": c.hasDeptSegment = True
            c.counterWA = SET_INV_WA: c.counterWD = SET_INV_WD
            c.colNo = IL_NO: c.colDept = IL_DEPT: c.colDate = IL_DATE
            c.colTotal = IL_TOTAL: c.colPaid = IL_PAID
            c.colBalance = IL_BALANCE: c.colStatus = IL_STATUS
        Case "QTE"
            c.docKind = "QTE": c.sheetName = "Quote"
            c.logName = "QuoteLog": c.linesName = "QuoteLines"
            c.prefix = "Q-": c.hasDeptSegment = True
            c.counterWA = SET_QUOTE_WA: c.counterWD = SET_QUOTE_WD
            c.colNo = QL_NO: c.colDept = QL_DEPT: c.colDate = QL_DATE
            c.colTotal = QL_TOTAL: c.colPaid = 0
            c.colBalance = 0: c.colStatus = QL_STATUS
        Case "CN"
            c.docKind = "CN": c.sheetName = "CreditNote"
            c.logName = "CreditNotes": c.linesName = "CreditNoteLines"
            c.prefix = "CN-": c.hasDeptSegment = False
            c.counterFlat = SET_CN
            c.colNo = 1: c.colDept = 0: c.colDate = 3
            c.colTotal = 5: c.colPaid = 0: c.colBalance = 0: c.colStatus = 8
        Case "MC"
            c.docKind = "MC": c.sheetName = "Med Claim"
            c.logName = "MedAidLog": c.linesName = "MedAidLines"
            c.prefix = "MC-INV-": c.hasDeptSegment = True
            c.counterFlat = SET_MC   ' single counter, dept in the number string
            c.colNo = ML_NO: c.colDept = ML_DEPT: c.colDate = ML_DATE
            c.colTotal = ML_TOTAL: c.colPaid = ML_PAID
            c.colBalance = ML_BALANCE: c.colStatus = ML_STATUS
    End Select
    GetDocConfig = c
End Function
