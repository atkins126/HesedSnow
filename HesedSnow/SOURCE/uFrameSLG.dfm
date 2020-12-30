inherited frmSLG: TfrmSLG
  Width = 895
  Height = 556
  ExplicitWidth = 895
  ExplicitHeight = 556
  object sPanel1: TsPanel [0]
    Left = 0
    Top = 0
    Width = 895
    Height = 41
    Align = alTop
    TabOrder = 0
    ExplicitLeft = 232
    ExplicitTop = 72
    ExplicitWidth = 185
    DesignSize = (
      895
      41)
    object SpeedButton2: TSpeedButton
      Left = 8
      Top = 8
      Width = 193
      Height = 27
      Caption = #1048#1084#1087#1086#1088#1090' '#1091#1089#1083#1091#1075' '#1057#1051#1043
      OnClick = SpeedButton2Click
    end
    object SpeedButton3: TSpeedButton
      Left = 705
      Top = 10
      Width = 23
      Height = 22
    end
    object SpeedButton1: TSpeedButton
      Left = 734
      Top = 8
      Width = 153
      Height = 27
      Anchors = [akTop, akRight]
      Caption = #1048#1084#1087#1086#1088#1090' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1072' '#1057#1051#1043
      OnClick = SpeedButton1Click
    end
  end
  object sPanel2: TsPanel [1]
    Left = 0
    Top = 41
    Width = 895
    Height = 474
    Align = alClient
    TabOrder = 1
    ExplicitLeft = 208
    ExplicitTop = 304
    ExplicitWidth = 185
    ExplicitHeight = 41
    object Splitter1: TSplitter
      Left = 653
      Top = 1
      Height = 472
      Align = alRight
      ExplicitLeft = 849
      ExplicitTop = -4
    end
    object Splitter2: TSplitter
      Left = 241
      Top = 1
      Height = 472
      ExplicitLeft = 696
      ExplicitTop = -4
    end
    object Panel1: TPanel
      Left = 1
      Top = 1
      Width = 240
      Height = 472
      Align = alLeft
      Caption = 'Panel1'
      TabOrder = 0
      ExplicitHeight = 462
      object sDBGrid1: TsDBGrid
        Left = 1
        Top = 22
        Width = 238
        Height = 449
        Align = alClient
        Color = 15921906
        DataSource = DM.dsUslugy
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
        OnDragDrop = sDBGrid1DragDrop
        OnKeyPress = sDBGrid1KeyPress
      end
      object sEdit1: TsEdit
        Left = 1
        Top = 1
        Width = 238
        Height = 21
        Align = alTop
        TabOrder = 1
        OnChange = sEdit1Change
        ExplicitLeft = 16
        ExplicitTop = 48
        ExplicitWidth = 121
      end
    end
    object Panel2: TPanel
      Left = 656
      Top = 1
      Width = 238
      Height = 472
      Align = alRight
      Caption = 'Panel1'
      TabOrder = 1
      object sEdit2: TsEdit
        Left = 1
        Top = 1
        Width = 236
        Height = 21
        Align = alTop
        TabOrder = 0
        OnChange = sEdit2Change
        ExplicitLeft = 117
        ExplicitTop = 40
        ExplicitWidth = 121
      end
      object sDBGrid2: TsDBGrid
        Left = 1
        Top = 22
        Width = 236
        Height = 449
        Align = alClient
        Color = 15921906
        DataSource = DM.dsQslg
        DrawingStyle = gdsGradient
        GradientEndColor = 13353918
        GradientStartColor = 14539223
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnKeyPress = sDBGrid2KeyPress
      end
    end
    object Panel3: TPanel
      Left = 244
      Top = 1
      Width = 409
      Height = 472
      Align = alClient
      Caption = 'Panel3'
      TabOrder = 2
      ExplicitLeft = 0
      ExplicitTop = 41
      ExplicitWidth = 895
      ExplicitHeight = 474
      object Panel4: TPanel
        Left = 1
        Top = 1
        Width = 407
        Height = 29
        Align = alTop
        TabOrder = 0
        ExplicitWidth = 893
        DesignSize = (
          407
          29)
        object Button3: TButton
          Left = 3
          Top = 2
          Width = 97
          Height = 25
          Action = btnDelStroka
          TabOrder = 0
        end
        object Button4: TButton
          Left = 304
          Top = 2
          Width = 101
          Height = 25
          Action = btnClearList
          Anchors = [akTop, akRight]
          TabOrder = 1
          ExplicitLeft = 790
        end
        object Button5: TButton
          Left = 103
          Top = 2
          Width = 75
          Height = 25
          Action = btnSaveTempFile
          TabOrder = 2
        end
        object Button6: TButton
          Left = 181
          Top = 2
          Width = 75
          Height = 25
          Action = btnLoadTempFile
          TabOrder = 3
        end
      end
      object StringGrid1: TJvStringGrid
        Left = 1
        Top = 30
        Width = 407
        Height = 441
        Align = alClient
        ColCount = 7
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing]
        TabOrder = 1
        Alignment = taLeftJustify
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = clWindowText
        FixedFont.Height = -11
        FixedFont.Name = 'Tahoma'
        FixedFont.Style = []
        ExplicitWidth = 893
        ExplicitHeight = 443
      end
    end
  end
  object sPanel3: TsPanel [2]
    Left = 0
    Top = 515
    Width = 895
    Height = 41
    Align = alBottom
    TabOrder = 2
    ExplicitLeft = 216
    ExplicitTop = 256
    ExplicitWidth = 185
    DesignSize = (
      895
      41)
    object Button2: TButton
      Left = 697
      Top = 5
      Width = 193
      Height = 31
      Action = btnExportExcel
      Anchors = [akRight, akBottom]
      TabOrder = 0
    end
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 40
    Top = 132
  end
  object acList: TActionList
    Left = 39
    Top = 237
    object btnDelStroka: TAction
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1089#1090#1088#1086#1082#1091
      OnExecute = Button3Click
      OnUpdate = btnDelStrokaUpdate
    end
    object btnClearList: TAction
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1089#1087#1080#1089#1086#1082
      OnExecute = Button4Click
      OnUpdate = btnClearListUpdate
    end
    object btnExportExcel: TAction
      Caption = #1045#1082#1089#1087#1086#1088#1090' '#1074' Excel'
      OnExecute = Button2Click
      OnUpdate = btnExportExcelUpdate
    end
    object btnLoadTempFile: TAction
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
      OnExecute = Button6Click
      OnUpdate = btnLoadTempFileUpdate
    end
    object btnSaveTempFile: TAction
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      OnExecute = Button5Click
      OnUpdate = btnSaveTempFileUpdate
    end
  end
  object OpenDialog: TOpenDialog
    Filter = 'xlsx|*.xlsx'
    InitialDir = #1047#1072#1074#1072#1085#1090#1072#1078#1077#1085#1085#1103
    Left = 40
    Top = 181
  end
end
