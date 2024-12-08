let
    ソース = Csv.Document(
        File.Contents("path/to/csv"), [
            Delimiter = ",",
            Columns = 12,
            Encoding = 65001,
            QuoteStyle = QuoteStyle.None
        ]
    ),
    変更された型 = Table.TransformColumnTypes(
        ソース,
        {
            {"Column1", type text},
            {"Column2", type text},
            {"Column3", type text},
            {"Column4", type text},
            {"Column5", type text},
            {"Column6", type text},
            {"Column7", type text},
            {"Column8", type text},
            {"Column9", type text},
            {"Column10", type text},
            {"Column11", type text},
            {"Column12", type text}
        }
    ),
    昇格されたヘッダー数 = Table.PromoteHeaders(変更された型, [PromoteAllScalars = true]),
    変更された型1 = Table.TransformColumnTypes(昇格されたヘッダー数, {{"検査カテゴリ", type text}}),
    追加されたインデックス = Table.AddIndexColumn(変更された型1, "インデックス", 1, 1, Int64.Type),
    // すべての列で「〃〃」を null に置き換え
    null置換 = Table.ReplaceValue(追加されたインデックス, "〃〃", null, Replacer.ReplaceValue, Table.ColumnNames(昇格されたヘッダー数)),
    // null を上の値で埋める
    埋められた値 = Table.FillDown(null置換, Table.ColumnNames(null置換))
in
    埋められた値
