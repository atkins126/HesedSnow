inherited frmVidomost: TfrmVidomost
  Width = 864
  Height = 553
  ExplicitWidth = 864
  ExplicitHeight = 553
  object sGradientPanel1: TsGradientPanel [0]
    Left = 0
    Top = 0
    Width = 864
    Height = 41
    Align = alTop
    Caption = 'sGradientPanel1'
    ShowCaption = False
    TabOrder = 0
    ExplicitWidth = 1063
    DesignSize = (
      864
      41)
    object labInfoStatus: TsLabelFX
      Left = 422
      Top = 14
      Width = 434
      Height = 21
      Alignment = taCenter
      Anchors = [akTop, akBottom]
      Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1074#1110#1076#1086#1084#1110#1089#1090#1110' '#1059#1057#1055#1030#1064#1053#1040'!!!'
      Color = clBtnFace
      ParentColor = False
      Angle = 0
      Shadow.OffsetKeeper.LeftTop = -3
      Shadow.OffsetKeeper.RightBottom = 5
      Shadow.Mode = smCustom
    end
    object sButton1: TsButton
      Left = 5
      Top = 8
      Width = 137
      Height = 25
      Caption = #1042#1090#1103#1085#1091#1090#1080' '#1076#1072#1085#1085#1110
      TabOrder = 0
      OnClick = sButton1Click
    end
  end
  object sPanel1: TsPanel [1]
    Left = 0
    Top = 41
    Width = 864
    Height = 512
    Align = alClient
    Caption = 'sPanel1'
    TabOrder = 1
    ExplicitWidth = 1063
    object DBGridEhVedomist: TDBGridEh
      Left = 1
      Top = 1
      Width = 862
      Height = 469
      Align = alClient
      DynProps = <>
      Flat = True
      FooterRowCount = 1
      IndicatorOptions = [gioShowRowIndicatorEh, gioShowRecNoEh]
      MinAutoFitWidth = 80
      OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghFitRowHeightToText, dghDialogFind, dghShowRecNo, dghColumnResize, dghColumnMove, dghAutoFitRowHeight, dghExtendVertLines]
      RowHeight = 2
      RowLines = 1
      RowSizingAllowed = True
      SumList.Active = True
      TabOrder = 0
      Columns = <
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Num_Uslugy'
          Footers = <>
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'JDCID'
          Footers = <>
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'FIO'
          Footers = <>
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Count_usl'
          Footer.ValueType = fvtCount
          Footers = <>
          Width = 20
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Cost_usl'
          Footer.ValueType = fvtSum
          Footers = <>
          Width = 60
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Programma'
          Footers = <>
          Width = 100
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Curator'
          Footers = <>
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Osobie_Proecty'
          Footers = <>
          Width = 20
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Zhertva'
          Footers = <>
          Width = 40
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Mobila'
          Footers = <>
          Width = 100
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Adress'
          Footers = <>
          Width = 200
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'S_Kem'
          Footers = <>
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'SABA'
          Footers = <>
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Type_Uchasnika'
          Footers = <>
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'Data_Plan'
          Footers = <>
          Width = 80
        end>
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
    object sPanel2: TsPanel
      Left = 1
      Top = 470
      Width = 862
      Height = 41
      Align = alBottom
      TabOrder = 1
      ExplicitLeft = 109
      ExplicitTop = 464
      ExplicitWidth = 746
      DesignSize = (
        862
        41)
      object sLabelFX1: TsLabelFX
        Left = 16
        Top = 12
        Width = 87
        Height = 17
        Caption = #1053#1072#1079#1074#1072' '#1074#1110#1076#1086#1084#1086#1089#1090#1110':'
        Angle = 0
        Shadow.OffsetKeeper.LeftTop = -1
        Shadow.OffsetKeeper.RightBottom = 3
      end
      object btnCreateVidomist: TsButton
        Left = 684
        Top = 8
        Width = 171
        Height = 25
        Anchors = [akRight, akBottom]
        Caption = 'btnCreateVidomist'
        TabOrder = 0
        OnClick = btnCreateVidomistClick
      end
      object DBComboBoxEh1: TDBComboBoxEh
        Left = 128
        Top = 9
        Width = 534
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        DynProps = <>
        EditButtons = <>
        Items.Strings = (
          #1097#1097#1097
          #1076#1076#1076
          #1073#1073#1073#1073)
        TabOrder = 1
        Visible = True
      end
    end
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 264
    Top = 8
  end
  object OpenDialog: TOpenDialog
    Left = 360
    Top = 9
  end
end
