object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 250
  Width = 125
  object tVedomost: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'Vidomist'
    Left = 16
    Top = 64
  end
  object myConnection: TADOConnection
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
end
