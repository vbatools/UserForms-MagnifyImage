VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMagnify 
   Caption         =   "Magnify Image"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   OleObjectBlob   =   "FMagnify.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : FMagnify
'* Created    : 16-03-2021 19:16
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_ZoomFactor As Double
Private Sub CheckBox1_Click()

    If CheckBox1.Value Then
        Image2.AutoSize = False
        Image2.PictureSizeMode = fmPictureSizeModeStretch
        Image2.Width = Image2.Width * 2
        Image2.Height = Image2.Height
    Else
        Image2.PictureSizeMode = fmPictureSizeModeClip
        Image2.AutoSize = True
    End If

End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 1 Then
        With Frame1
            .Left = X
            .Top = Y
            .Visible = True
        End With
    End If

End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim RatioX      As Double
    Dim RatioY      As Double

    If Button = 1 Then
        Frame1.Left = Image1.Left + X - (Frame1.Width / 2)
        Frame1.Top = Image1.Top + Y - (Frame1.Height / 2)

        RatioX = X / Image1.Width
        RatioY = Y / Image1.Height

        Image2.Left = -(Image2.Width * RatioX) + (Frame1.Width / 2)
        Image2.Top = -(Image2.Height * RatioY) + (Frame1.Height / 2)

    End If
End Sub


Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 1 Then
        Frame1.Left = Image1.Left + Image1.Width
        Frame1.Top = Image1.Top + Image1.Height
        Frame1.Visible = False
    End If
End Sub


Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()

    Image2.Picture = Image1.Picture
    Image2.AutoSize = True
    Frame1.SpecialEffect = fmSpecialEffectRaised
    Frame1.Visible = False

    m_ZoomFactor = 1
    ZoomFactor.List = Array("1.0", "1.5", "2.0", "2.5", "3.0", "3.5", "4.0", "5.0", "10.0")
    ZoomFactor.ListIndex = 0

End Sub


Private Sub ZoomFactor_Change()

    m_ZoomFactor = CDbl(Val(ZoomFactor.Value))

    Image2.AutoSize = True
    If ZoomFactor.ListIndex > 0 Then
        Image2.AutoSize = False
        Image2.PictureSizeMode = fmPictureSizeModeStretch
        Image2.Width = Image2.Width * m_ZoomFactor
        Image2.Height = Image2.Height * m_ZoomFactor
    Else
        Image2.PictureSizeMode = fmPictureSizeModeClip
    End If

End Sub

