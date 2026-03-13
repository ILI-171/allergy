Option Explicit

Public Function settings_allergens() As Variant

    settings_allergens = Array("えび", "かに", "くるみ", "小麦", "そば", "卵", "乳", "落花生", "アーモンド", "あわび", "いか", "いくら", "オレンジ", "カシューナッツ", "キウイフルーツ", "牛肉", "ごま", "さけ", "さば", "大豆", "鶏肉", "バナナ", "豚肉", "まつたけ", "もも", "やまいも", "りんご", "ゼラチン")

End Function

Public Function settings_vendor_sheets(include_part_recipe As Boolean) As Variant

    If include_part_recipe Then
        settings_vendor_sheets = Array("業者１", "業者２", "業者３", "業者４", "業者５", "業者６", "業者７", "業者８", "パートレシピ")
    Else
        settings_vendor_sheets = Array("業者１", "業者２", "業者３", "業者４", "業者５", "業者６", "業者７", "業者８")
    End If

End Function

Public Function settings_menu_category_sheets() As Variant

    settings_menu_category_sheets = Array("冷菜　魚", "冷菜　肉", "冷菜　その他", "温製　魚", "温製　肉", "温製　その他", "デザート　冷製", "デザート　温製")

End Function

Public Function settings_menu_category_colors() As Variant

    settings_menu_category_colors = Array(rgbSkyBlue, rgbSkyBlue, rgbSkyBlue, rgbSandyBrown, rgbSandyBrown, rgbSandyBrown, rgbPaleGreen, rgbPaleGreen)

End Function

Public Function settings_recipe_clear_columns() As Variant

    settings_recipe_clear_columns = Array(2, 7, 8, 9, 10, 11, 12, 15)

End Function

Public Function settings_recipe_source_columns() As Variant

    settings_recipe_source_columns = Array(2, 6, 7, 10, 12, "AS")

End Function

Public Function settings_recipe_target_columns() As Variant

    settings_recipe_target_columns = Array(2, 7, 9, 10, 12, 15)

End Function

Public Function settings_marker_columns() As Variant

    settings_marker_columns = Array(8, 9, 10, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38)

End Function

Public Function settings_checkbox_rows() As Variant

    settings_checkbox_rows = Array(22, 25, 28)

End Function

Public Function settings_checkbox_label_rows() As Variant

    settings_checkbox_label_rows = Array(21, 24, 27)

End Function

Public Function settings_checkbox_end_columns() As Variant

    settings_checkbox_end_columns = Array(11, 13, 13)

End Function

Public Function settings_menu_display_sheets() As Variant

    settings_menu_display_sheets = Array("メニュー表示(1-50)", "メニュー表示(51-100)", "メニュー表示(101-150)", "メニュー表示(151-200)")

End Function

Public Function settings_merge_col_ranges() As Variant

    settings_merge_col_ranges = Array("A:K", "L:V")

End Function

Public Function settings_alergy_copy_base_source_columns() As Variant

    settings_alergy_copy_base_source_columns = Array("A", "B", "C", "D", "E")

End Function

Public Function settings_alergy_copy_base_target_columns() As Variant

    settings_alergy_copy_base_target_columns = Array("B", "E", "K", "Q", "AA")

End Function
