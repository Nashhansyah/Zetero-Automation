Attribute VB_Name = "ZoteroAutomate"
Option Explicit

' =====================================================================
'  Jalankan dari Word:
'  Alt+F11 -> File -> Import File -> pilih file ini
'  Lalu: Developer -> Macros -> UpdateZoteroCitations
'
'  Dokumen aktif disimpan, diproses oleh zotero_cite.py, lalu hasilnya
'  dibuka sebagai dokumen baru. Dokumen asli tidak diubah isinya.
' =====================================================================

Private Const SCRIPT_DIR As String = "C:\ZOTERO AUTOMATION"

Sub UpdateZoteroCitations()
    Dim docPath As String
    Dim outPath As String
    Dim cmd As String
    Dim rc As Long

    If ActiveDocument.Path = "" Then
        MsgBox "Simpan dokumen ini dulu sebelum dijalankan.", vbExclamation, "Zotero Automation"
        Exit Sub
    End If

    If Dir(SCRIPT_DIR & "\zotero_cite.py") = "" Then
        MsgBox "Tidak menemukan zotero_cite.py di:" & vbCrLf & SCRIPT_DIR, vbCritical
        Exit Sub
    End If

    ActiveDocument.Save
    docPath = ActiveDocument.FullName
    outPath = Left$(docPath, InStrRev(docPath, ".") - 1) & " - Zotero.docx"

    cmd = "cmd /c cd /d """ & SCRIPT_DIR & """ && python zotero_cite.py"
    cmd = cmd & " --docx """ & docPath & """"
    cmd = cmd & " --out """ & outPath & """"

    rc = CreateObject("WScript.Shell").Run(cmd, 1, True)

    If Dir(outPath) <> "" Then
        Documents.Open outPath
        MsgBox "Selesai." & vbCrLf & vbCrLf & _
               "Langkah berikutnya di tab Zotero:" & vbCrLf & _
               "  1. Refresh" & vbCrLf & _
               "  2. Add/Edit Bibliography" & vbCrLf & vbCrLf & _
               "Jangan lupa hapus daftar pustaka lama yang diketik manual.", _
               vbInformation, "Zotero Automation"
    Else
        MsgBox "Proses gagal (exit code " & rc & ")." & vbCrLf & _
               "Lihat jendela console untuk pesan errornya.", vbExclamation, "Zotero Automation"
    End If
End Sub
