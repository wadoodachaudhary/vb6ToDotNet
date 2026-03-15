VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DragDropAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' DragDropAttachment Class
'
' Class to get dropped Outlook attachments
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media,  without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows XP Pro SP1
'
' History:
'          06/12/2003 * Added SaveAttachment method
'          12/30/2002 * Code was released
'
'*********************************************************************************************
Option Explicit

Private CF_FILECONTENTS As Integer
Private CF_FILEGROUPDESCRIPTOR As Integer
Private CF_FILEGROUPDESCRIPTORW As Integer

Private m_oSource As IDataObject

'
' Attachment
'
' Returns the attachment contents
'
' Parameters:
'
'  Index - Zero based index of the attachment
'
Function Attachment(ByVal Index As Long, Optional ByVal SaveAs As String) As Variant
Dim oStorage As IStorage
Dim oStream As IStream
Dim tFMT As FORMATETC
Dim tSTGM As STGMEDIUM
Dim tStat As STATSTG
Dim abData() As Byte
  
   With tFMT
      .cfFormat = CF_FILECONTENTS
      .dwAspect = DVASPECT_CONTENT
      .lindex = Index
      .TYMED = TYMED_ISTORAGE Or TYMED_ISTREAM
   End With
   
   If m_oSource.GetData(tFMT, tSTGM) = S_OK Then
   
      If (tSTGM.TYMED And TYMED_ISTORAGE) = TYMED_ISTORAGE Then
      
         Dim oILB As ILockBytes, oStgMem As IStorage
         Dim tIID As UUID
         
         ' Create a storage object in memory
         Set oILB = CreateILockBytesOnHGlobal(0, 1)
         Set oStgMem = StgCreateDocfileOnILockBytes(oILB, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0)
         
         ' Get the storage from the data object
         Set oStorage = pvStorageFromSTGMEDIUM(tSTGM)
                       
         ' Copy storage contents from
         ' data object to memory
         oStorage.CopyTo 0, tIID, vbNullString, oStgMem
   
         ' Save changes
         oStgMem.Commit STGC_DEFAULT
         
         ' Get memory stream info
         oILB.Stat tStat, STATFLAG_NONAME
         
         ReDim abData(0 To tStat.cbSize * 10000 - 1)
         
         ' Get the data
         oILB.ReadAt 0, abData(0), tStat.cbSize * 10000
         
         ' Return the array
         Attachment = abData()
         
         ' Release all objects
         Set oStorage = Nothing
         Set oStgMem = Nothing
         Set oILB = Nothing
      
      Else
         
         ' Get the stream
         Set oStream = pvStreamFromSTGMEDIUM(tSTGM)
         
         ' Get stream info
         oStream.Stat tStat, STATFLAG_NONAME
         
         ReDim abData(0 To tStat.cbSize * 10000 - 1)
         
         ' Read all data from the stream
         oStream.Read abData(0), tStat.cbSize * 10000
         
         ' Return the array
         Attachment = abData()
         
         ' Release the stream
         Set oStream = Nothing
         
      End If
            
      ' Release the storage medium
      ReleaseStgMedium tSTGM
   
   Else
      Err.Raise 5
   End If

End Function

'
' SaveAttachment
'
' Saves the attachment to a file
'
' Parameters:
'
'  Index - Zero based index of the attachment
'  Filename - Name of the destination file
'
Sub SaveAttachment(ByVal Index As Long, ByVal Filename As String)
Dim oStorage As IStorage
Dim oStream As IStream
Dim tFMT As FORMATETC
Dim tSTGM As STGMEDIUM
Dim tStat As STATSTG
Dim abData() As Byte
  
   With tFMT
      .cfFormat = CF_FILECONTENTS
      .dwAspect = DVASPECT_CONTENT
      .lindex = Index
      .TYMED = TYMED_ISTORAGE Or TYMED_ISTREAM
   End With
   
   If m_oSource.GetData(tFMT, tSTGM) = S_OK Then
   
      If (tSTGM.TYMED And TYMED_ISTORAGE) = TYMED_ISTORAGE Then
      
         Dim oStgMem As IStorage
         Dim tIID As UUID
         
         ' Create a the new storage file
         Set oStgMem = StgCreateDocfile(Filename, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE)
         
         ' Get the storage from the data object
         Set oStorage = pvStorageFromSTGMEDIUM(tSTGM)
                       
         ' Copy storage contents from
         ' data object to memory
         oStorage.CopyTo 0, tIID, vbNullString, oStgMem
         
         ' Release all objects
         Set oStorage = Nothing
         Set oStgMem = Nothing
      
      Else
         
         Dim lFF As Long
         
         ' Get the stream
         Set oStream = pvStreamFromSTGMEDIUM(tSTGM)
         
         ' Get stream info
         oStream.Stat tStat, STATFLAG_NONAME
         
         ReDim abData(0 To tStat.cbSize * 10000 - 1)
         
         ' Read all data from the stream
         oStream.Read abData(0), tStat.cbSize * 10000
         
         ' Save the data
         lFF = FreeFile
         Open Filename For Binary As lFF
         Put lFF, , abData()
         Close
         
         ' Release the stream
         Set oStream = Nothing
         
      End If
            
      ' Release the storage medium
      ReleaseStgMedium tSTGM
   
   Else
      Err.Raise 5
   End If

End Sub


'
' FileName
'
' Return the attachment file name
'
' Parameters:
'
'  Index - Zero based index of the attachment
'
Function Filename(ByVal Index As Long) As String
Dim tFMT As FORMATETC
Dim tSTGM As STGMEDIUM
Dim tFGD As FILEGROUPDESCRIPTORA
Dim tFD As FILEDESCRIPTORA
Dim lPtr As Long
      
   With tFMT
      .cfFormat = CF_FILEGROUPDESCRIPTOR
      .dwAspect = DVASPECT_CONTENT
      .lindex = -1
      .TYMED = TYMED_HGLOBAL
   End With
   
   ' Get the FILEGROUPDESCRIPTOR
   If m_oSource.GetData(tFMT, tSTGM) <> S_OK Then
      
      ' Try with CF_FILEGROUPDESCRIPTOR
      tFMT.cfFormat = CF_FILEGROUPDESCRIPTORW
      
      If m_oSource.GetData(tFMT, tSTGM) <> S_OK Then
         Exit Function
      End If
      
   End If
 
   ' Get the pointer to the data
   lPtr = GlobalLock(tSTGM.Data)
   
   ' Copy the data to the UDT
   MoveMemory tFGD, ByVal lPtr, Len(tFGD)
   
   If tFGD.cItems < Index Then
      
      ' Release the pointer
      GlobalUnlock tSTGM.Data
      
      ' Release the data
      ReleaseStgMedium tSTGM
      
      Err.Raise 5
      
   End If
   
   ' Get the FILEDESCRIPTOR for the attachment
   MoveMemory tFD, ByVal lPtr + 4& + Len(tFD) * Index, Len(tFD)
      
   If tFMT.cfFormat = CF_FILEGROUPDESCRIPTOR Then
      
      Filename = StrConv(tFD.cFileName, vbUnicode)
   
      ' Remove null chars from the end of the string
      Filename = Left$(Filename, InStr(Filename, vbNullChar) - 1)
   
   Else
      Filename = Left$(tFD.cFileName, InStr(tFD.cFileName, vbNullChar) - 1)
   End If
   
   ' Release the pointer
   GlobalUnlock tSTGM.Data
   
   ' Release the data
   ReleaseStgMedium tSTGM
   
End Function

'
' Count
'
' Returns the # of attachments in the DataObject
'
Function Count() As Long
Dim tFMT As FORMATETC
Dim tSTGM As STGMEDIUM
Dim tFGD As FILEGROUPDESCRIPTORA
Dim lPtr As Long
      
   With tFMT
      .cfFormat = CF_FILEGROUPDESCRIPTOR
      .dwAspect = DVASPECT_CONTENT
      .lindex = -1
      .TYMED = TYMED_HGLOBAL
   End With
   
   ' Get the FILEGROUPDESCRIPTOR
   If m_oSource.GetData(tFMT, tSTGM) <> S_OK Then
      
      ' Try with CF_FILEGROUPDESCRIPTORW
      tFMT.cfFormat = CF_FILEGROUPDESCRIPTORW
      
      If m_oSource.GetData(tFMT, tSTGM) <> S_OK Then
         Exit Function
      End If
      
   End If
    
   ' Get the pointer to the data
   lPtr = GlobalLock(tSTGM.Data)
   
   ' Copy the data to the UDT
   MoveMemory tFGD, ByVal lPtr, Len(tFGD)
   
   ' Return the number of messages
   Count = tFGD.cItems
   
   ' Release the pointer
   GlobalUnlock tSTGM.Data
   
   ' Release the data
   ReleaseStgMedium tSTGM
   
End Function

Private Function pvStorageFromSTGMEDIUM(tSTGM As STGMEDIUM) As IStorage
Dim oUnk As IUnknown

   MoveMemory oUnk, tSTGM.Data, 4&
   Set pvStorageFromSTGMEDIUM = oUnk
   MoveMemory oUnk, 0&, 4&
   
End Function

Private Function pvStreamFromSTGMEDIUM(tSTGM As STGMEDIUM) As IStream
Dim oUnk As IUnknown

   MoveMemory oUnk, tSTGM.Data, 4&
   Set pvStreamFromSTGMEDIUM = oUnk
   MoveMemory oUnk, 0&, 4&
   
End Function

'
' IDataObjectFromData
'
' Returns the IDataObject interface from
' a DataObject object
'
Private Function pvIDataObjectFromData(ByVal Data As DataObject) As IDataObject
Dim oTmpData As IDataObject
   
   ' Get an uncounted reference
   ' to the IDataObject interface
   MoveMemory oTmpData, ByVal ObjPtr(Data) + 16, 4
   
   ' Get a counted reference
   Set pvIDataObjectFromData = oTmpData
   
   ' Release the uncounted reference
   MoveMemory oTmpData, 0&, 4
   
End Function

Property Set Source(ByVal New_Source As stdole.IUnknown)

   If New_Source Is Nothing Then
      Set m_oSource = OleGetClipboard()
   ElseIf TypeOf New_Source Is IDataObject Then
      Set m_oSource = New_Source
   ElseIf TypeOf New_Source Is VBRUN.DataObject Then
      Set m_oSource = pvIDataObjectFromData(New_Source)
   Else
      Err.Raise 5
   End If
   
End Property

Private Sub Class_Initialize()
   
   ' Get the clipboard formats
   CF_FILECONTENTS = RegisterClipboardFormat("FileContents") And &H7FFF& Or &H8000
   CF_FILEGROUPDESCRIPTOR = RegisterClipboardFormat("FileGroupDescriptor") And &H7FFF& Or &H8000
   CF_FILEGROUPDESCRIPTORW = RegisterClipboardFormat("FileGroupDescriptorW") And &H7FFF& Or &H8000
   
   ' Initialize the source object
   Set m_oSource = OleGetClipboard()
   
End Sub
