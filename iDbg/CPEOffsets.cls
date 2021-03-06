VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPEOffsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'License:   GPL
'Copyright: David Zimmer 2002
'Site:      http://sandsprite.com
'
'Author:    David Zimmer <dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA


Private Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Private Type IMAGE_NT_HEADERS
    Signature As String * 4
    'Public Type IMAGE_FILE_HEADER
        Machine As Integer
        NumberOfSections As Integer
        TimeDateStamp As Long
        PointerToSymbolTable As Long
        NumberOfSymbols As Long
        SizeOfOptionalHeader As Integer
        Characteristics As Integer
    'End Type
    'Public Type IMAGE_OPTIONAL_HEADER
        Magic As Integer
        MajorLinkerVersion As Byte
        MinorLinkerVersion As Byte
        SizeOfCode As Long
        SizeOfInitializedData As Long
        SizeOfUninitializedData As Long
        AddressOfEntryPoint As Long
        BaseOfCode As Long
        BaseOfData As Long
        ImageBase As Long
        SectionAlignment As Long
        FileAlignment As Long
        MajorOperatingSystemVersion As Integer
        MinorOperatingSystemVersion As Integer
        MajorImageVersion As Integer
        MinorImageVersion As Integer
        MajorSubsystemVersion As Integer
        MinorSubsystemVersion As Integer
        Win32VersionValue As Long
        SizeOfImage As Long
        SizeOfHeaders As Long
        CheckSum As Long
        Subsystem As Integer
        DllCharacteristics As Integer
        SizeOfStackReserve As Long
        SizeOfStackCommit As Long
        SizeOfHeapReserve As Long
        SizeOfHeapCommit As Long
        LoaderFlags As Long
        NumberOfRvaAndSizes As Long
        rvaExportTable As Long
            ExportTableSize As Long
        rvaImportTable As Long
        DataDirectoryX(1 To 116) As Byte
    'End Type
End Type

Private Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Private DOSHEADER As IMAGEDOSHEADER
Private NTHEADER As IMAGE_NT_HEADERS
Private Section() As IMAGE_SECTION_HEADER
Private myFile As String

Public errMessage  As String
Public ImageBase As Long
Public AddressOfEntryPoint As Long
Public RvaImportDirectory As Long
Public RvaExportDirectory As Long

Function LoadFile(fpath As String) As Boolean
    On Error GoTo hell
        
        Dim i As Long
        Dim f As Long
        Dim udtSH As IMAGE_SECTION_HEADER
        Dim sectionStart As Long
        
        errMessage = Empty
        myFile = Empty
        If Not FileExists(fpath) Then Err.Raise 1

        f = FreeFile
        Open fpath For Binary Access Read As f
        Get f, , DOSHEADER
        Get f, DOSHEADER.e_lfanew + 1, NTHEADER
                    
        sectionStart = DOSHEADER.e_lfanew + 1 + _
                       NTHEADER.SizeOfOptionalHeader + 24
                       ' + Len(NTHEADER.FileHeader) +  Len(NTHEADER.Signature)
                       
        Seek f, sectionStart
        
        ReDim Section(1 To NTHEADER.NumberOfSections)
        For i = 1 To NTHEADER.NumberOfSections
            Get f, , udtSH
            Section(i) = udtSH
        Next

        ImageBase = NTHEADER.ImageBase
        AddressOfEntryPoint = NTHEADER.AddressOfEntryPoint
        RvaImportDirectory = NTHEADER.rvaImportTable
        RvaExportDirectory = NTHEADER.rvaExportTable
        
        Close f
        LoadFile = True
        myFile = fpath
        
Exit Function
hell:
    
    Close f
    
    Select Case Err.Number
        Case 1: errMessage = "File not found: " & fpath
        Case Else: errMessage = Err.Description
    End Select
    
End Function

Function RvaToOffset(ByVal rva As Long, Optional ByRef VirtualSectionBase As Long, Optional sectName As String)
        
    If UBound(Section) = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As IMAGE_SECTION_HEADER
    Dim i As Integer
    Dim vBase As Long, vSize As Long, vMax As Long
        
    For i = 1 To UBound(Section)
       c = Section(i)
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            rva = rva - vBase
            rva = rva + c.PointerToRawData
            RvaToOffset = rva
            VirtualSectionBase = vBase
            sectName = c.nameSec
            Exit Function
       End If
    Next
       
End Function

Function OffsetToRVA(ByVal fOffset As Long, Optional sectName As String)
    
    If UBound(Section) = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As IMAGE_SECTION_HEADER
    Dim rBase As Long, rSize As Long, rMax As Long, i As Long
    
    For i = 1 To UBound(Section)
       c = Section(i)
       rBase = c.PointerToRawData
       rSize = c.SizeOfRawData
       rMax = rBase + rSize
       
       If fOffset >= rBase And fOffset < rMax Then 'right section
            OffsetToRVA = c.VirtualAddress + fOffset - rBase
            sectName = c.nameSec
            Exit Function
       End If
    Next

End Function

Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function
