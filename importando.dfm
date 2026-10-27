object frmImportando: TfrmImportando
  Left = 0
  Top = 0
  Caption = 'Importando Dados'
  ClientHeight = 196
  ClientWidth = 309
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 96
    Top = 16
    Width = 128
    Height = 25
    Caption = 'Importando...'
    StyleElements = [seClient, seBorder]
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHotLight
    Font.Height = -21
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object Label2: TLabel
    Left = 60
    Top = 72
    Width = 52
    Height = 24
    Align = alCustom
    Alignment = taCenter
    Caption = '0/100'
    StyleElements = [seClient, seBorder]
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -20
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    Transparent = True
    Layout = tlCenter
  end
  object Label3: TLabel
    Left = 136
    Top = 112
    Width = 36
    Height = 16
    Caption = 'Status'
    StyleElements = [seClient, seBorder]
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object BitBtn1: TBitBtn
    Left = 113
    Top = 152
    Width = 80
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 0
    OnClick = BitBtn1Click
  end
end
