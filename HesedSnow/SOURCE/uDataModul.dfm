object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 250
  Width = 355
  object tVedomost: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'Vidomist'
    Left = 16
    Top = 64
  end
  object myConnection: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\GitCopy\HesedSno' +
      'w\BIN\Win32\Debug\Snow.mdb;Persist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 16
    Top = 8
  end
  object dsVedomist: TDataSource
    DataSet = tVedomost
    Left = 16
    Top = 112
  end
  object qQuery: TADOQuery
    Connection = myConnection
    Parameters = <>
    Left = 16
    Top = 168
  end
  object ADOTable1: TADOTable
    Left = 304
    Top = 64
  end
  object DataSource1: TDataSource
    Left = 304
    Top = 112
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 304
    Top = 168
  end
  object tUslugy: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'Uslugy'
    Left = 80
    Top = 64
  end
  object dsUslugy: TDataSource
    DataSet = qUslugy
    Left = 80
    Top = 168
  end
  object qUslugy: TADOQuery
    Connection = myConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select Uslugy.[FIO], Uslugy.[JDCID],Uslugy.[RITM],Uslugy.[Number' +
        '],Uslugy.[SABA],Uslugy.[City] from Uslugy')
    Left = 80
    Top = 120
  end
end
