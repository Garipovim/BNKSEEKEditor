object Form1: TForm1
  Left = 122
  Top = 147
  Width = 1363
  Height = 692
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object DBGrid2: TDBGrid
    Left = 0
    Top = 512
    Width = 1345
    Height = 73
    DataSource = DataSource2
    TabOrder = 10
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Visible = False
  end
  object Button1: TButton
    Left = 472
    Top = 8
    Width = 75
    Height = 25
    Caption = #1042#1099#1073#1088#1072#1090#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 8
    Top = 8
    Width = 449
    Height = 21
    TabOrder = 1
  end
  object Button2: TButton
    Left = 560
    Top = 8
    Width = 75
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 2
    OnClick = Button2Click
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 88
    Width = 1345
    Height = 369
    DataSource = DataSource1
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    TabOrder = 3
    TitleFont.Charset = OEM_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = Button5Click
  end
  object LabeledEdit1: TLabeledEdit
    Left = 24
    Top = 63
    Width = 121
    Height = 21
    EditLabel.Width = 22
    EditLabel.Height = 13
    EditLabel.Caption = #1041#1048#1050
    TabOrder = 4
    OnKeyPress = LabeledEdit1KeyPress
  end
  object LabeledEdit2: TLabeledEdit
    Left = 160
    Top = 63
    Width = 121
    Height = 21
    EditLabel.Width = 44
    EditLabel.Height = 13
    EditLabel.Caption = #1056#1045#1043#1048#1054#1053
    TabOrder = 5
  end
  object DBLookupComboBox1: TDBLookupComboBox
    Left = 291
    Top = 63
    Width = 145
    Height = 21
    KeyField = 'NAME'
    ListFieldIndex = 1
    ListSource = DataSource2
    TabOrder = 6
  end
  object Button3: TButton
    Left = 450
    Top = 63
    Width = 119
    Height = 22
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100' '#1092#1080#1083#1100#1090#1088
    Enabled = False
    TabOrder = 7
    OnClick = Button3Click
  end
  object BitBtn1: TBitBtn
    Left = 573
    Top = 472
    Width = 153
    Height = 25
    Hint = #1044#1086#1073#1072#1074#1083#1077#1085#1080#1077' '#1085#1086#1074#1086#1081' '#1089#1090#1088#1086#1082#1080
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1085#1086#1074#1091#1102' '#1089#1090#1088#1086#1082#1091
    TabOrder = 8
    Visible = False
    OnClick = BitBtn1Click
  end
  object Button4: TButton
    Left = 586
    Top = 63
    Width = 119
    Height = 22
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1092#1080#1083#1100#1090#1088
    Enabled = False
    TabOrder = 9
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 457
    Top = 472
    Width = 105
    Height = 25
    Hint = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100' '#1090#1077#1082#1091#1097#1091#1102' '#1089#1090#1088#1086#1082#1091
    Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
    TabOrder = 11
    Visible = False
    OnClick = Button5Click
  end
  object Button6: TButton
    Left = 737
    Top = 472
    Width = 153
    Height = 25
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100
    TabOrder = 12
    Visible = False
    OnClick = Button6Click
  end
  object CNDBF: TADOConnection
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 640
    Top = 8
  end
  object ADOQuery1: TADOQuery
    Connection = CNDBF
    Parameters = <>
    Left = 704
    Top = 8
  end
  object DataSource1: TDataSource
    AutoEdit = False
    DataSet = ADOQuery1
    Left = 672
    Top = 8
  end
  object DataSource2: TDataSource
    DataSet = ADOQuery2
    Left = 776
    Top = 8
  end
  object ADOQuery2: TADOQuery
    Connection = CNDBF
    Parameters = <>
    Left = 808
    Top = 8
  end
  object DataSource3: TDataSource
    DataSet = ADOQuery3
    Left = 872
    Top = 8
  end
  object ADOQuery3: TADOQuery
    Connection = CNDBF
    Parameters = <>
    Left = 904
    Top = 8
  end
  object DataSource4: TDataSource
    DataSet = ADOQuery4
    Left = 976
    Top = 8
  end
  object ADOQuery4: TADOQuery
    Connection = CNDBF
    Parameters = <>
    Left = 1008
    Top = 8
  end
  object ADOQuery5: TADOQuery
    Connection = CNDBF
    Parameters = <>
    Left = 1072
    Top = 8
  end
  object DataSource5: TDataSource
    DataSet = ADOQuery5
    Left = 1048
    Top = 8
  end
end
