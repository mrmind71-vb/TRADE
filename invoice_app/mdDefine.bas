Attribute VB_Name = "inv_Define"
Public idSrvBaseUrl As String
Public TokenUrl As String
Public client_id As String
Public client_secret As String
Public apiBaseUrl As String
Public SubmitUrl As String
Public inv_version As String
Public taxpayerActivityCode As String
Public taxRate As Integer
Public subType As String
Public showSign As Boolean
Public Const errorFlag  As String = "@@error@@"
Public Function DefineUrl(con As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select top 1 * From SettingInvoice", con)
If Not loctable.EOF Then
    idSrvBaseUrl = loctable!idSrvBaseUrl & ""
    TokenUrl = loctable!TokenUrl & ""
    client_id = loctable!client_id & ""
    client_secret = loctable!client_secret & ""
    inv_version = loctable!Version & ""
    apiBaseUrl = loctable!apiBaseUrl & ""
    SubmitUrl = loctable!SubmitUrl & ""
    taxpayerActivityCode = loctable!ActivityCode
    taxRate = mRound(loctable!taxRate)
    subType = loctable!subType & ""
    showSign = loctable!showSign
    DefineUrl = True
End If
Set loctable = Nothing
End Function
Public Sub eInform(Mcaption As String, Optional mCaption2 As String, Optional nInterval As Integer = 900)
On Error Resume Next
eInformfrm.sLabel1 = Mcaption
eInformfrm.sLabel2 = mCaption2
eInformfrm.nInterval = nInterval
eInformfrm.Show 1
DoEvents
Err.Clear
End Sub
Public Function myFormat_z(pDate As String, Optional pFormat As String = "YYYY-MM-DDThh:mm:00Z") As String
myFormat_z = Format(pDate, pFormat)
End Function

