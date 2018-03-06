# VBA_Excel

Sub test_dict_of_dicts()

  Dim arr_BR_TR, arr_TR_Tcode As Variant

  arr_BR_TR = make_array_from_two_col_mapping(sh_BR_TR)
  arr_TR_Tcode = make_array_from_two_col_mapping(sh_TR_Tcode)

  Dim test_dict_BR As New Scripting.Dictionary
  Set test_dict_BR = make_dict_of_dicts(arr_BR_TR)

  Dim test_dict_TR As New Scripting.Dictionary
  Set test_dict_TR = make_dict_of_dicts(arr_TR_Tcode)

  Debug.Print test_dict_BR.Count
  Debug.Print test_dict_TR.Count

End Sub
