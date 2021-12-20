inherited frmFrameDavayPodkl: TfrmFrameDavayPodkl
  Width = 751
  Height = 582
  ExplicitWidth = 751
  ExplicitHeight = 582
  object sPanel1: TsPanel [0]
    Left = 0
    Top = 0
    Width = 751
    Height = 41
    Align = alTop
    TabOrder = 0
    object btnDownloadUslugy: TsBitBtn
      Left = 17
      Top = 10
      Width = 225
      Height = 25
      Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1092#1072#1081#1083' '#1079' '#1087#1086#1089#1083#1091#1075#1072#1084#1080
      Layout = blGlyphRight
      ModalResult = 12
      NumGlyphs = 2
      Spacing = 40
      TabOrder = 0
      OnClick = btnDownloadUslugyClick
      Grayed = True
      TextAlignment = taLeftJustify
    end
    object btnOtchet: TsBitBtn
      Left = 512
      Top = 10
      Width = 225
      Height = 25
      Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1079#1074#1110#1090#1085#1080#1081' '#1092#1072#1081#1083' '
      TabOrder = 1
      OnClick = btnOtchetClick
    end
  end
  object sPanel2: TsPanel [1]
    Left = 0
    Top = 41
    Width = 751
    Height = 500
    Align = alClient
    Caption = 'sPanel1'
    TabOrder = 1
    object Splitter1: TSplitter
      Left = 1
      Top = 163
      Width = 749
      Height = 3
      Cursor = crVSplit
      Align = alTop
      ExplicitWidth = 336
    end
    object sDBGrid1: TsDBGrid
      Left = 1
      Top = 1
      Width = 749
      Height = 162
      Align = alTop
      Color = 15921906
      DataSource = DM.dsTemaDP
      DrawingStyle = gdsGradient
      GradientEndColor = 13353918
      GradientStartColor = 14539223
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'Data'
          Width = 60
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'Tema'
          Width = 600
          Visible = True
        end>
    end
    object StringGrid: TJvStringGrid
      Left = 1
      Top = 166
      Width = 749
      Height = 333
      Align = alClient
      ColCount = 2
      RowCount = 1
      FixedRows = 0
      TabOrder = 1
      Alignment = taLeftJustify
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = []
    end
  end
  object sPanel3: TsPanel [2]
    Left = 0
    Top = 541
    Width = 751
    Height = 41
    Align = alBottom
    TabOrder = 2
    object btnCreateExcel: TsBitBtn
      Left = 512
      Top = 8
      Width = 225
      Height = 25
      Action = abtnCreateExcel
      Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1092#1072#1081#1083's'
      TabOrder = 0
    end
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 256
    Top = 8
  end
  object ActionList: TActionList
    Left = 336
    Top = 8
    object abtnCreateExcel: TAction
      Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1092#1072#1081#1083's'
      OnExecute = btnCreateExcelClick
      OnUpdate = abtnCreateExcelUpdate
    end
  end
end
