Attribute VB_Name = "MxConst"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxConst."
':OfLimPmSpec$ = "It is lin with 3 Pm: [-Sw xx -Sng xx -Mul xx].  Where -Sw xx are allowed switch.  " & _
"Where -Sng xx are allowed single-value Pm.  " & _
"Where -Mul xx are allowed multiple-value Pm.  " & _
"Where xx are PmNm."
Public Const vbOpnBkt$ = "("
Public Const vbDblQ$ = """"
Public Const vb2CrLf$ = vbCrLf & vbCrLf
Public Const vb2DblQ$ = vbDblQ & vbDblQ
Public Const vbDblQAsc As Byte = 34
Public Const vbSngQ$ = "'"
Public Const vbExcM$ = "!"
Public Const vbPround$ = "#"
Public Const vbOpnSqBkt$ = "["
Public Const vbOpnBigBkt$ = "{"
Public Fso As New Scripting.FileSystemObject
Const H$ = "C:\Users\User\Desktop\MHD\SAPAccessReports\"
Const H1$ = "C:\Users\User\Desktop\"
'------------------------------------------
'From:
'https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql/sql-server-express-user-instances
Public Const SampCnStr_SQLEXPR$ = "Data Source=.\\SQLExpress;Integrated Security=true;" & _
"User Instance=true;AttachDBFilename=|DataDirectory|\InstanceDB.mdf;" & _
"Initial Catalog=InstanceDB;"
'------------------------------------------
'From:
'https://social.msdn.microsoft.com/Forums/vstudio/en-US/61d45bef-eea7-4366-a8ad-e15a1fa3d544/vb6-to-connect-with-sqlexpress?forum=vbgeneral
Public Const SampCnStr_SQLEXPR_NotWrk3$ = _
"Provider=SQLNCLI.1;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
Public Const GetCnStr_ADO_SampSQL_EXPR_NOT_WRK$ = _
"Provider=LoSqleDb;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
'--------------------------------
'From https://social.msdn.microsoft.com/Forums/en-US/a73a838b-ec3f-419b-be65-8b1732fbf4d0/connect-to-a-remote-sql-server-db?forum=isvvba
Public Const SampCnStr_SQLEXPR_NotWrk1$ = "driver={SQL Server};" & _
      "server=LAPTOP-SH6AEQSO;uid=MyUserName;pwd=;database=pubs"
   
Public Const SampCnStr_SQLEXPR_NotWrk2$ = "driver={SQL Server};" & _
      "server=127.0.0.1;uid=MyUserName;pwd=;database=pubs"
   
Public Const SampCnStr_SQLEXPR_NotWrk$ = ".\SQLExpress;AttachDbFilename=c:\mydbfile.mdf;Database=dbname;" & _
"Trusted_Connection=Yes;"
'"Typical normal SQL Server connection string: Data Source=myServerAddress;
'"Initial Catalog=myDataBase;Integrated Security=SSPI;"

'From VisualStudio
Public Const SampSqlCnStr_NotWrk$ = _
    "Data Source=LAPTOP-SH6AEQSO\ProjectsV13;Initial Catalog=master;Integrated Security=True;Connect Timeout=30;" & _
    "Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"

Public Const SampFbzDutyDta$ _
                                    = H & "DutyPrepay5\DutyPrepay5_Data.mdb"
Public Const SampFbzDutyPgm$ _
                                    = H & "DutyPrepay5\DutyPrepay5.accdb"
Public Const SampFxzKE24 _
                                    = H & "DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Public Const SampFbzDutyzPgmBackup$ _
                                    = H & "DutyPrepay5\DutyPrepay5_BackupFfn.accdb"
Public Const SampFbzTaxCmp$ _
                                    = H1 & "QFinalSln\TaxExpCmp v1.3.accdb"
Public Const SampFbzShpRate$ _
                                    = H1 & "QFinalSln\StockShipRate (ver 1.0).accdb"
Public Const SampFbzShpCst$ = "C:\Users\user\Documents\Projects\Vba\ShpCst\ShpCstApp.accdb"
Public Const SampFx$ = SampFxzKE24
Property Get SampDbShpCst() As Database
Set SampDbShpCst = Db(SampFbzShpCst)
End Property

Property Get DbEng() As DBEngine
Set DbEng = DAO.DBEngine
End Property

Property Get SampCnzDutyDta() As ADODB.Connection
Set SampCnzDutyDta = CnzFb(SampFbzDutyDta)
End Property
Property Get SampFb$()
SampFb = SampFbzDutyDta
End Property
Property Get SampDb() As Database
Set SampDb = Db(SampFb)
End Property
Property Get SampDbDutyDta() As Database
Set SampDbDutyDta = Db(SampFbzDutyDta)
End Property

Sub AAAAA()
Dim A
'{00024500-0000-0000-C000-000000000046}
Set A = Interaction.CreateObject("{00024500-0000-0000-C000-000000000046}", "Excel.Application")
Stop
End Sub
