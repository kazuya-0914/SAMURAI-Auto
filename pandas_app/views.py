from django.shortcuts import render
from django.views import View
import pandas as pd

# --- トップページ --- #
def index(request):
    return render(request, 'pandas.html')

# --- 7章 pandasでデータを操作しよう --- #
class Chap07(View):
    dir = 'pandas_app/static/excel/'
    template = "chap07.html"

    def dispatch(self, request, *args, **kwargs):
        if request.path.endswith('csv_read/'):
            return self.csv_read(request, *args, **kwargs)
        elif request.path.endswith('csv_write/'):
            return self.csv_write(request, *args, **kwargs)
        elif request.path.endswith('excel_read/'):
            return self.excel_read(request, *args, **kwargs)
        elif request.path.endswith('excel_write/'):
            return self.excel_write(request, *args, **kwargs)
        elif request.path.endswith('data_filter/'):
            return self.data_filter(request, *args, **kwargs)
        elif request.path.endswith('data_mean/'):
            return self.data_mean(request, *args, **kwargs)
        elif request.path.endswith('data_result/'):
            return self.data_result(request, *args, **kwargs)
        else:
            return self.top_page(request, *args, **kwargs)
    
    # トップページ
    def top_page(self, request, *args, **kwargs):
        context = {}
        return render(request, self.template, context)
    
    # CSVデータの読み込み方法
    def csv_read(self, request, *args, **kwargs):
        df = pd.read_csv(f"{self.dir}Chapter7_1.csv")
        pre = f"{df}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # CSVデータの書き込み方法
    def csv_write(self, request, *args, **kwargs):
        df = pd.DataFrame(data = {
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon']
        })
        df.to_csv(f"{self.dir}Chapter7_2.csv", index=False)
        context = { 'msg': 'CSVファイルを作成しました' }
        return render(request, self.template, context)
    
    # Excelファイルの読み込み方法
    def excel_read(self, request, *args, **kwargs):
        df = pd.read_excel(f"{self.dir}Chapter7_3.xlsx")
        pre = f"{df}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # Excelファイルの書き込み方法
    def excel_write(self, request, *args, **kwargs):
        df = pd.DataFrame(data = {
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon']
        })
        writer = pd.ExcelWriter(f"{self.dir}Chapter7_4.xlsx")
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()
        context = { 'msg': 'Excelファイルを作成しました' }
        return render(request, self.template, context)
    
    # データのフィルタリング方法
    def data_filter(self, request, *args, **kwargs):
        df = pd.DataFrame(data = {
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon']
        })
        filter = df[(df['ID'] >= 2)]
        pre = f"{filter}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # データの集計方法
    def data_mean(self, request, *args, **kwargs):
        df = pd.DataFrame(data = {
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon'],
            '値段': [100, 200, 600]  
        })
        mean = df['値段'].mean()
        pre = mean
        context = { 'pre': pre }
        return render(request, self.template, context)

    # データの変換方法
    def data_result(self, request, *args, **kwargs):
        df = pd.DataFrame(data = {
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon'],
            '値段': [100, 200, 600]  
        })
        def transform(cost):
            return cost * 10
        transform_result= df['値段'].apply(transform)
        pre = f"{transform_result}"
        context = { 'pre': pre }
        return render(request, self.template, context)

# --- 8章 pandasで型変換やデータフレーム結合をしよう --- #
class Chap08(View):
    dir = 'pandas_app/static/excel/'
    template = "chap08.html"

    def dispatch(self, request, *args, **kwargs):
        if request.path.endswith('data_iloc/'):
            return self.data_iloc(request, *args, **kwargs)
        elif request.path.endswith('to_numeric/'):
            return self.to_numeric(request, *args, **kwargs)
        elif request.path.endswith('to_string/'):
            return self.to_string(request, *args, **kwargs)
        elif request.path.endswith('loc/'):
            return self.loc(request, *args, **kwargs)
        elif request.path.endswith('concat/0'):
            return self.concat(request, *args, **kwargs)
        elif request.path.endswith('concat/1'):
            return self.concat(request, *args, **kwargs)
        else:
            return self.top_page(request, *args, **kwargs)

    # トップページ
    def top_page(self, request, *args, **kwargs):
        context = {}
        return render(request, self.template, context)

    # データの指定方法
    def data_iloc(self, request, *args, **kwargs):
        df = pd.read_csv(f"{self.dir}Chapter7_1.csv")
        df2 = df.iloc[1:3, [0, 2]]
        pre = f"{df2}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # データを数値型に変換する方法
    def to_numeric(self, request, *args, **kwargs):
        df = pd.DataFrame({
            'ID': ['1', '2', '3'],
            'point': ['80', '90', '70']
        })
        pre = f"{df.dtypes}"
        pre += "\n---------------------\n"

        df['point'] = pd.to_numeric(df['point'])
        pre += f"{df.dtypes}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # データを文字列型に変換する方法
    def to_string(self, request, *args, **kwargs):
        data = {
            'Name': ['Alice', 'Bob', 'Charlie', 'David'],
            'Age': [24, 28, 22, 35],
            'City': ['New York', 'Los Angeles', 'London', 'Tokyo']
        }
        df = pd.DataFrame(data)

        pre = "to_string()使用前:\n"
        pre += f"{df}\n"
        pre += f"{type(df)}\n"

        pre += "\n---------------------\n"

        df_string = df.to_string()
        pre += "to_string()使用後:\n"
        pre += f"{df_string}\n"
        pre += f"{type(df_string)}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
    # データの選択方法
    def loc(self, request, *args, **kwargs):
        df = pd.DataFrame({
            'ID': [1, 2, 3],
            '日本語': ['りんご', 'ぶどう', 'レモン'],
            '英語': ['apple', 'grape', 'lemon']
        })

        name_col = df.loc[:, '日本語']
        pre = f"{name_col}"
        context = { 'pre': pre }
        return render(request, self.template, context)

    # データの結合方法
    def concat(self, request, axis, *args, **kwargs):
        df1 = pd.DataFrame({'ID': [1, 2],
            '日本語': ['りんご', 'ぶどう'],
            '英語': ['apple', 'grape']
        })
        df2 = pd.DataFrame({'ID': [3, 4],
            '日本語': ['レモン', 'バナナ'],
            '英語': ['lemon', 'banana']
        })

        if axis == 0:
            result = pd.concat([df1, df2], axis=0) # 縦結合
        else:
            result = pd.concat([df1, df2], axis=1) # 横結合

        pre = f"{result}"
        context = { 'pre': pre }
        return render(request, self.template, context)
    
# --- 9章 pandasで実際にExcelファイルに書き出そう --- #
class Chap09(View):
    template = "chap09.html"
    dir = 'pandas_app/static/excel/'

    def dispatch(self, request, *args, **kwargs):
        if request.path.endswith('practice/'):
            return self.practice(request, *args, **kwargs)
        else:
            return self.top_page(request, *args, **kwargs)

    # トップページ
    def top_page(self, request, *args, **kwargs):
        context = {}
        return render(request, self.template, context)

    # スケジュール管理表を作成
    def practice(self, request, *args, **kwargs):
        # データフレームの作成
        df = pd.DataFrame({
            '日付': ['2023-05-17', '2023-05-18', '2023-05-19', '2023-05-20', '2023-05-21'],
            'スケジュール': ['設計', '開発', 'テスト', '運用', '保守'],
            '優先レベル': [1, 2, 5, 4, 3],
            '状況': ['完了', '作業中', '作業中', '未着手', '未着手'],
        })

        # 優先レベルの平均を求めて新しい列を作成
        df['平均レベル'] = df['優先レベル'].mean()

        # 「緊急度」列を作成し、関数「prioritize」を適用して値を設定
        df['緊急度'] = df['優先レベル'].apply(prioritize)

        # Excelファイルを作成
        writer = pd.ExcelWriter(f"{self.dir}スケジュール管理表.xlsx")

        # DataFrameオブジェクトをExcelファイルに書き込む
        df.to_excel(writer, sheet_name='Sheet1', index=False)

        # Excelファイルを閉じる
        writer.close()

        # コンテキスト
        msg = "スケジュール管理表を作成しました"
        context = { 'msg': msg }
        return render(request, self.template, context)

# 優先レベルで分岐して緊急度を求める関数「prioritize」を定義
def prioritize(level):
    result = ''
    if level >= 5:
        result = '高'
    elif level == 4 or level == 3:
        result = '中'
    else:
        result = '低'
    return result