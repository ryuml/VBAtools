VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageUserForm 
   Caption         =   "�}�N�����s�O"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MessageUserForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MessageUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub UserForm_Activate()
    Me.Caption = "�}�N�����s��..."
    Me.StatusLabel.Caption = "�}�N�����s���ł�..."
    Call GetLinksDirAllFiles
    Me.Caption = "�}�N�����s����"
    Me.StatusLabel.Caption = "�}�N�����s�������܂���"
End Sub
