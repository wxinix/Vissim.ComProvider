object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'EmbeddedVissim Example'
  ClientHeight = 545
  ClientWidth = 949
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  DesignSize = (
    949
    545)
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 48
    Top = 32
    Width = 97
    Height = 25
    Caption = 'StartVissim'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 48
    Top = 80
    Width = 97
    Height = 25
    Caption = 'ShowVissim'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 48
    Top = 128
    Width = 97
    Height = 25
    Caption = 'HideVissim'
    TabOrder = 2
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 48
    Top = 176
    Width = 97
    Height = 25
    Caption = 'EmbeddedVissim'
    TabOrder = 3
    OnClick = Button4Click
  end
  object Panel1: TPanel
    Left = 176
    Top = 23
    Width = 765
    Height = 514
    Anchors = [akLeft, akTop, akRight, akBottom]
    BorderStyle = bsSingle
    Caption = 'Panel1'
    Color = clWindow
    ParentBackground = False
    TabOrder = 4
  end
end
