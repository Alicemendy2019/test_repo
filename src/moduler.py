
######### merge
import pandas as pd
from pathlib import Path

def combine_files(p, extension, output_ex):
    try:
        # パスを設定
        # p = Path(r"D:\temp\travel")

        # 全Excelファイルを取得
        if extension == "csv":
            files = list(p.glob("*.csv"))
            files += list(p.glob("*.CSV"))
        else:
            files = list(p.glob(f"*.{extension}"))

        # 最初のファイルを読み込む
        if extension == "csv":
            merged_df = pd.read_csv(files[0])
        else:
            merged_df = pd.read_excel(files[0])
        print(2)
        # 2つ目のファイル以降を読み込む
        for file in files[1:]:
            if extension == "csv":
                temp_df = pd.read_csv(file, header=None).iloc[1:]
            else:
                temp_df = pd.read_excel(file, header=None).iloc[1:]
            temp_df.columns = merged_df.columns  # 最初に読み込んだファイルのカラム名を指定
            merged_df = pd.concat([merged_df, temp_df])

        print(3)
        # 結合したデータフレームを新しいExcelファイルに保存
        if output_ex == "EXCEL":
            out_p = p / "combine.xlsx"
            merged_df.to_excel(out_p, index=False)
        else:
            out_p = p / "combine.csv"
            merged_df.to_csv(out_p, index=False)
        return out_p
    except Exception as e:
        print(e)
        raise Exception(e)