# coding: utf-8

# Dummy Japanese character: あ（意味はないが確実に日本語を含むファイルにする）



# Gurobi Optimizerのパッケージをインポートし、
# "gp"という短縮名を設定（違う短縮名にしてもよい）
import gurobipy as gp

# Excelファイルを読み込むためにpandasをインポート
import pandas as pd



# クラスを定義
# 社員クラス
class Shain():
    def __init__(self, row):
        self.code = row[0]
        self.name = row[1]

    def __repr__(self):
        return f"{self.code}_{self.name}"

# 仕事クラス
class Shigoto():
    def __init__(self, row):
        self.code = row[0]
        self.name = row[1]

    def __repr__(self):
        return f"{self.code}_{self.name}"



# データを読み込み
excel_file = pd.ExcelFile("./Data.xlsx")

# 「社員」シートをDataFrameとして読み込み
df_shain = excel_file.parse("社員")
# 「社員コード」「社員名」の列を抽出
df_shain_part = df_shain[["社員コード", "社員名"]]
# listに変換
list_shain = df_shain_part.values.tolist()
# 社員オブジェクト集合を定義（※便宜上、setの代わりにlistを使用）
I = list()
# 「社員コード : 社員オブジェクト」の辞書を定義
dict_I = dict()
# 社員オブジェクト集合にShainクラスのオブジェクトを追加し、
# 「社員コード : 社員オブジェクト」の辞書にも追加
for row in list_shain:
    shain = Shain(row)
    I.append(shain)
    dict_I[row[0]] = shain
print(f"[集合・定数]")
print(f"    社員集合: {I}")

# 「仕事」シートについて同様の処理
J = list(
    [Shigoto(row) 
     for row 
     in excel_file.parse("仕事")[["仕事コード", "仕事名"]].values.tolist()])
dict_J = {shigoto.code : shigoto for shigoto in J}
print(f"    仕事集合: {J}")


# 「費用」シートをDataFrameとして読み込み、
# 「社員コード」「仕事コード」「費用」の列を抽出し、listに変換
list_hiyo = excel_file.parse("費用")[
    ["社員コード", "仕事コード", "費用"]].values.tolist()
# c[社員オブジェクト, 仕事オブジェクト] を 費用 を表す辞書として定義
c = {}
# c[社員オブジェクト, 仕事オブジェクト] に 費用 を設定
for (shain_code, shigoto_code, hiyo) in list_hiyo:
    shain = dict_I[shain_code]
    shigoto = dict_J[shigoto_code]
    c[shain, shigoto] = hiyo
print(f"    費用定数: {c}")
print()



# 問題を設定
model_2 = gp.Model(name = "GurobiSample2")


# 変数を設定（変数単体にかかる制約を含む）
# x[社員オブジェクト, 仕事オブジェクト] を gp.Var クラスのオブジェクト を
# 表す辞書として定義
x = {}
for i in I:
    for j in J:
        x[i, j] = model_2.addVar(vtype = gp.GRB.BINARY, name = f"x({i},{j})")


# 目的関数を設定
model_2.setObjective(gp.quicksum(c[i, j] * x[i, j] for i in I for j in J), 
                     sense = gp.GRB.MINIMIZE)


# 制約を設定
con_1 = {}
for i in I:
    con_1[i] = model_2.addConstr(gp.quicksum(x[i, j] for j in J) <= 1, 
                                 name = f"con_1({i})")
con_2 = {}
for j in J:
    con_1[j] = model_2.addConstr(gp.quicksum(x[i, j] for i in I) >= 1, 
                                 name = f"con_2({j})")


# 解を求める計算
print("[Gurobi Optimizerログ]")

model_2.optimize()

print()


# 最適解が得られた場合、結果を出力
print("[解]")
if model_2.Status == gp.GRB.OPTIMAL:
    print("    最適解: ")
    for i in I:
        for j in J:
            if x[i, j].X > 0.98:
                print(f"        {i.name}に{j.name}を割り当て")
    val_opt = model_2.ObjVal
    print(f"    最適値: {val_opt}")