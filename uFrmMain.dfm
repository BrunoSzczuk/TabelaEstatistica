object frmMain: TfrmMain
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  Caption = 'Trabalho de Estat'#237'stica - Bruno'
  ClientHeight = 97
  ClientWidth = 293
  Color = clWindow
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 43
    Top = 11
    Width = 59
    Height = 13
    Caption = 'Valor inicial: '
  end
  object Label2: TLabel
    Left = 46
    Top = 38
    Width = 51
    Height = 13
    Caption = 'Amplitude:'
  end
  object Label3: TLabel
    Left = 34
    Top = 65
    Width = 67
    Height = 13
    Caption = 'Nr Intervalos:'
  end
  object edVlInicial: TEdit
    Left = 105
    Top = 8
    Width = 73
    Height = 21
    Alignment = taCenter
    NumbersOnly = True
    TabOrder = 0
    Text = '150'
  end
  object edAmplitude: TEdit
    Left = 104
    Top = 35
    Width = 73
    Height = 21
    Alignment = taCenter
    NumbersOnly = True
    TabOrder = 1
    Text = '4'
  end
  object edNrIntervalos: TEdit
    Left = 103
    Top = 62
    Width = 73
    Height = 21
    Alignment = taCenter
    NumbersOnly = True
    TabOrder = 2
    Text = '6'
  end
  object Button1: TButton
    Left = 210
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Gerar Tabela'
    TabOrder = 3
    OnClick = Button1Click
  end
end
