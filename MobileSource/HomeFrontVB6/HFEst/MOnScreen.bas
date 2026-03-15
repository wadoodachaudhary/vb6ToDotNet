Attribute VB_Name = "MOnScreen"
Option Explicit
Option Compare Text
Private Const SRCFILE = "MOnScreen::"


Public OnScreenConnection  As ADODB.Connection

Public Function OpenOnScreen() As Boolean
    Dim s As String
    Set OnScreenConnection = New ADODB.Connection
    With HFApp.Options
        If .ValueByName("OnScreenType") <> "SQL" Then
            s = ""
            If .ValueByName("OnScreenMDB") <> "" Then
                s = s & "Provider=Microsoft.Jet.OLEDB.4.0;"
                s = s & "Data Source=" & .ValueByName("OnScreenMDB") & ";"
            End If
        Else
            s = ""
            If .ValueByName("OnScreenServer") <> "" Then
                s = s & "Provider=sqloledb;"
                s = s & "Data Source=" & .ValueByName("OnScreenServer") & ";"
                s = s & "Initial Catalog=" & .ValueByName("OnScreenDatabase") & ";"
                If .ValueByName("OnScreenUID") = "" Then
                    s = s & "Integrated Security=SSPI;"
                Else
                    s = s & "User Id=" & .ValueByName("OnScreenUID") & ";"
                    s = s & "Password=" & .ValueByName("OnScreenPWD") & ";"
                End If
            End If
        End If
    End With
    
    OpenOnScreen = False
    If s <> "" Then
        On Error Resume Next
        OnScreenConnection.Open s
        If Err.Number = 0 Then
            Call InitializeDB
        Else
            MsgBox "Error connecting to OnScreen Takeoff Database" & vbCrLf & vbCrLf & Err.Description, vbExclamation, App.ProductName
        End If
        OpenOnScreen = OnScreenConnection.State = adStateOpen
    End If
    
End Function
Public Sub CloseOnScreen()
    Set OnScreenConnection = Nothing
End Sub


Private Sub InitializeDB()
    Dim c As ADODB.Connection
    Set c = OnScreenConnection

On Error Resume Next
    
    c.Execute "drop table HF_Conversions"
    c.Execute "create table HF_Conversions(FromUnit varchar(4),ToUnit varchar(4),Factor float);"
    c.Execute "create unique index PK_HF_Conversions on HF_Conversions(fromunit,tounit);"
    
    c.Execute "drop table HF_QtyCodes"
    c.Execute "CREATE TABLE HF_QtyCodes(Code integer,Description varchar(50));"
    c.Execute "CREATE UNIQUE INDEX PK_HF_QtyCodes on HF_QtyCodes(Code);"
    
    c.Execute "drop table HF_UnitCodes"
    c.Execute "CREATE TABLE HF_UnitCodes(Code integer, Conversion float,Description varchar(4));"
    c.Execute "CREATE UNIQUE INDEX PK_HF_UnitCodes on HF_UnitCodes(Code);"
    
    
Conversions:
On Error GoTo QTYCODES
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN','LF',0.08333333333333330000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN','LY',0.02777777777777780000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN','m',0.02540000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN','mm',25.40000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LF','IN',12.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LF','LY',0.33333333333333333333);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LF','m',0.30480000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LF','mm',304.80000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LY','IN',36.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LY','LF',3);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LY','m',0.91440000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LY','mm',914.40000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m','IN',39.37000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m','LF',3.28100000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m','LY',1.09400000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m','mm',1000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm','IN',0.03937000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm','LF',0.00328000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm','LY',0.00109000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm','m',0.00100000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','IN²',14400.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','SF',100.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','SY',11.11111111111110000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','m²',9.29000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','mm²',9290000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','SQ',0.00006944444444444440);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','SF',0.00694000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','SY',0.00077200000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','m²',0.00064520000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','mm²',645.20000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','IN²',144.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','SY',0.11110000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','SQ',0.01000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','m²',0.09290000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','mm²',92900.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','IN²',1296.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','SF',9.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','SQ',0.09000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','m²',0.83610000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','mm²',836000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','IN²',1550.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','SF',10.76000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','SY',1.19600000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','SQ',0.10760000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','mm²',1000000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','IN²',0.00155000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','SF',0.00001080000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','SY',0.00000120000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','SQ',0.00000010800000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','m²',0.00000100000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CF','CY',0.03704000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CF','m³',0.02832000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CF','mm³',28316846.59200000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CY','CF',27.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CY','m³',0.76460000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CY','mm³',764600000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m³','CF',35.31000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m³','CY',1.30795100000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m³','mm³',1000000000.00000000000000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm³','CF',0.00000003531466672149);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm³','CY',0.00000000130787339786);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm³','m³',0.00000000100000000000);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('EA','EA',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LF','LF',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('LY','LY',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN','IN',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m','m',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm','mm',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SF','SF',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SY','SY',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('SQ','SQ',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('IN²','IN²',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m²','m²',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm²','mm²',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('m³','m³',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('mm³','mm³',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CF','CF',1);"
    c.Execute "insert into HF_Conversions(FromUnit,ToUnit,Factor) values('CY','CY',1);"
    
QTYCODES:
On Error GoTo UNITCODES
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(1,'Length');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(2,'Segment Count');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(3,'Surface Area (Single Side)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(4,'Surface Area (Both Sides)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(5,'Surface Area (Top OR Bottom)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(6,'Surface Area (Top AND Bottom)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(7,'Surface Area (Single End)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(8,'Surface Area (Both Ends)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(9,'Surface Area (All Side/Duct)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(11,'Area');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(12,'Area (Minus Attachments)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(13,'Perimeter');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(14,'Perimeter (Without Backout Perimeters)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(15,'Perimeter (plus Attachments Perimeter) ');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(16,'Grid Length (Visible)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(17,'Tile Count (Average)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(18,'Tile Count (Visible)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(19,'Area Counts');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(20,'Volume');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(21,'Volume (Ignore Blackout Volume)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(22,'Volume (Minus Attachments)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(23,'Count');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(24,'Total Height');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(25,'Surface Area (Single Width Side)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(26,'Surface Area (Both Width Sides) ');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(27,'Surface Area (Single Depth Side)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(28,'Surface Area (Both Depth Sides)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(29,'Surface Area (All Sides) ');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(30,'Surface Area (All Sides +Top and Bottom) ');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(31,'Perimeter (Four Sides)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(32,'Perimeter (Top + Both Sides) ');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(33,'Perimeter (Plus Attachment Perimeters)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(34,'Area (Ignore Backout Areas)');"
    c.Execute "insert into HF_QtyCodes(Code,Description) Values(35,'Volume (Ignore Backout Volume)');"

UNITCODES:
On Error GoTo EXITSUB
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(0,  1                                    ,'EA');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(1,  1                                    ,'IN');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(2,  0.083333333333333333333333333333333  ,'LF');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(3,  0.027777777777777777777777777777778  ,'LY');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(4,  1                                    ,'IN²');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(5,  0.0069444444444444444444444444444444 ,'SF');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(6,  0.0007716049382716049382716049382716 ,'SY');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(7,  0.0000694444444444444444444444444444 ,'SQ');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(8,  0.0005787037037037037037037037037037 ,'CF');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(9,  0.0000214334705075445816186556927297 ,'CY');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(11, 25.4000000001016000000004064         ,'mm');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(12, 645.16000000258064000001032256       ,'mm²');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(13, 0.0254000000001016000000004064       ,'m');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(14, 0.0006451612903225806451612903225806 ,'m²');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(15, 16387.06509958255590365323397091     ,'mm³');"
    c.Execute "insert into HF_UnitCodes(Code,Conversion,Description) Values(16, 0.0000163872638185602150009012995100 ,'m³');"

EXITSUB:
End Sub




