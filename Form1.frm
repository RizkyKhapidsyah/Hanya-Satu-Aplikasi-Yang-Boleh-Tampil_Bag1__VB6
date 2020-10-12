VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hanya Satu Aplikasi yang Boleh Tampil (1)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jalankan atau double click file exe yang Anda buat 'ini, lalu minimize-kan.
'Jalankan atau double click lagi file exe tadi...
'Karena aplikasi ini sudah dijalankan sebelumnya, maka 'aplikasi yang sama tidak dapat dijalankan lagi, dan 'akan muncul pesan konfirmasi.

Public Sub CheckSoftware(x As Form)
On Error GoTo Pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Program ini sedang dijalankan!", _
               vbCritical, "Sedang Dijalankan"
        App.Title = ""
        x.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
Pesan:
    End
    Exit Sub
End Sub

'Gunakan di saat event form_load
Private Sub Form_Load()
   Call CheckSoftware(Form1)
End Sub


