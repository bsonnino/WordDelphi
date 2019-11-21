object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Open Word File'
  ClientHeight = 471
  ClientWidth = 917
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 513
    Top = 41
    Height = 430
    ExplicitLeft = 464
    ExplicitTop = 208
    ExplicitHeight = 100
  end
  object Memo1: TMemo
    Left = 0
    Top = 41
    Width = 513
    Height = 430
    Align = alLeft
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Consolas'
    Font.Style = []
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 0
    WordWrap = False
  end
  object ValueListEditor1: TValueListEditor
    Left = 516
    Top = 41
    Width = 401
    Height = 430
    Align = alClient
    TabOrder = 1
    ColWidths = (
      150
      245)
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 917
    Height = 41
    Align = alTop
    TabOrder = 2
    ExplicitLeft = 192
    ExplicitTop = 152
    ExplicitWidth = 185
    object Button1: TButton
      Left = 0
      Top = 10
      Width = 75
      Height = 25
      Caption = 'Open'
      TabOrder = 0
      OnClick = Button1Click
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Word Files|*.docx|All Files|*.*'
    Left = 112
    Top = 8
  end
  object XMLDocument1: TXMLDocument
    Options = [doNodeAutoCreate, doNodeAutoIndent, doAttrNull, doAutoPrefix, doNamespaceDecl]
    Left = 160
    Top = 8
  end
  object XMLDocument2: TXMLDocument
    Left = 200
    Top = 8
  end
end
