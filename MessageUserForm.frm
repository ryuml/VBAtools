VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageUserForm 
   Caption         =   "マクロ実行前"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MessageUserForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MessageUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub UserForm_Activate()
    Me.Caption = "マクロ実行中..."
    Me.StatusLabel.Caption = "マクロ実行中です..."
    Call GetLinksDirAllFiles
    Me.Caption = "マクロ実行完了"
    Me.StatusLabel.Caption = "マクロ実行完了しました"
End Sub
