"""
スコア集計プログラム
CSVファイルから参加者のスコアを読み込み、平均・最高点・最低点を表形式で表示します。
平均値が最も高い行は赤色太字、最も低い行は青色太字で表示します。
"""

# 必要なライブラリをインポート
import csv  # CSVファイルを読み込むためのライブラリ
from collections import defaultdict  # 辞書を便利に使うためのライブラリ
import subprocess  # コマンドを実行するためのライブラリ
import sys  # システム関連の機能を使用するためのライブラリ

# coloramaのインストール確認とインポート
try:
    # coloramaをインポート（ターミナルの色付けに使用）
    from colorama import Fore, Style, init
    # coloramaの初期化（Windowsでも色が正しく表示されるようにする）
    init(autoreset=True)
    COLORAMA_AVAILABLE = True  # coloramaが使用可能であることを示すフラグ
except ImportError:
    # coloramaがインストールされていない場合
    print("coloramaがインストールされていません。インストールを開始します...")
    
    try:
        # pipを使ってcoloramaをインストール
        subprocess.check_call([sys.executable, "-m", "pip", "install", "colorama"], 
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # インストール後に再度インポートを試みる
        from colorama import Fore, Style, init
        init(autoreset=True)
        COLORAMA_AVAILABLE = True  # coloramaが使用可能であることを示すフラグ
        print("coloramaのインストールが完了しました。")
    except Exception:
        # インストールに失敗した場合、色なしで動作するようにする
        print("coloramaのインストールに失敗しました。色なしで表示します。")
        # ダミーのクラスを作成（色付けなし）
        class Fore:
            RED = ""
            BLUE = ""
        class Style:
            BRIGHT = ""
            RESET_ALL = ""
        COLORAMA_AVAILABLE = False  # coloramaが使用不可であることを示すフラグ

# CSVファイルのパスを指定
CSV_FILE = "課題2.csv"


def load_csv_data(filename):
    """
    CSVファイルを読み込み、データをリストとして返す関数
    
    Args:
        filename: 読み込むCSVファイルのパス
        
    Returns:
        データのリスト
    """
    data = []  # データを格納するリストを初期化
    
    # CSVファイルを開く（UTF-8エンコーディングで）
    with open(filename, 'r', encoding='utf-8') as file:
        # CSVリーダーを作成（辞書形式で読み込む）
        csv_reader = csv.DictReader(file)
        
        # ファイルの各行を読み込む
        for row in csv_reader:
            data.append(row)  # 行をリストに追加
    
    return data  # 読み込んだデータを返す


def calculate_statistics(data):
    """
    参加者ごとのスコア統計を計算する関数
    
    Args:
        data: CSVから読み込んだデータのリスト
        
    Returns:
        参加者名をキー、統計情報（平均、最高点、最低点）を値とする辞書
    """
    # 参加者ごとのスコアを格納する辞書（初期値は空のリスト）
    scores_by_person = defaultdict(list)
    
    # 各データ行を処理
    for row in data:
        name = row['名前']  # 参加者名を取得
        score = int(row['スコア'])  # スコアを整数に変換して取得
        
        # 参加者名をキーとして、スコアをリストに追加
        scores_by_person[name].append(score)
    
    # 統計情報を格納する辞書を初期化
    statistics = {}
    
    # 各参加者の統計情報を計算
    for name, scores in scores_by_person.items():
        average = sum(scores) / len(scores)  # 平均値を計算（合計÷個数）
        max_score = max(scores)  # 最高点を取得
        min_score = min(scores)  # 最低点を取得
        
        # 統計情報を辞書に保存
        statistics[name] = {
            'average': average,
            'max': max_score,
            'min': min_score
        }
    
    return statistics  # 計算した統計情報を返す


def find_max_min_average(statistics):
    """
    平均値の最大値と最小値を見つける関数
    
    Args:
        statistics: 参加者ごとの統計情報が入った辞書
        
    Returns:
        最大平均値と最小平均値のタプル
    """
    # 全ての参加者の平均値をリストとして取得
    averages = [stat['average'] for stat in statistics.values()]
    
    # 平均値の最大値を取得
    max_average = max(averages)
    
    # 平均値の最小値を取得
    min_average = min(averages)
    
    return max_average, min_average  # 値を返す


def format_score(score):
    """
    スコアを小数点第1位まで表示する関数
    
    Args:
        score: スコア（数値）
        
    Returns:
        フォーマットされた文字列
    """
    return f"{score:.1f}"  # 小数点第1位まで表示


def display_table(statistics):
    """
    統計情報を表形式で表示する関数
    平均値が最も高い行は赤色太字、最も低い行は青色太字で表示します。
    
    Args:
        statistics: 参加者ごとの統計情報が入った辞書
    """
    # 平均値の最大値と最小値を事前に計算
    max_average, min_average = find_max_min_average(statistics)
    
    # 表の列幅を設定（余裕をもたせる）
    name_width = 20  # 参加者名の列幅
    avg_width = 15   # 平均点の列幅
    max_width = 15   # 最高点の列幅
    min_width = 15   # 最低点の列幅
    
    # 表のヘッダーを表示
    print("\n" + "=" * 70)  # 区切り線を表示
    # ヘッダー行（列名を表示）
    print(f"{'参加者名':<{name_width}} {'平均点':<{avg_width}} {'最高点':<{max_width}} {'最低点':<{min_width}}")
    print("=" * 70)  # 区切り線を表示
    
    # 参加者ごとにデータを表示（名前でソート）
    for name in sorted(statistics.keys()):
        stat = statistics[name]  # その参加者の統計情報を取得
        
        # 数値を文字列に変換
        avg_str = format_score(stat['average'])
        max_str = str(stat['max'])
        min_str = str(stat['min'])
        
        # 色付けの判定
        # デフォルトの色（なし）
        color_start = ""
        color_end = ""
        
        # 平均値が最も高い場合：赤色太字
        if stat['average'] == max_average:
            color_start = Fore.RED + Style.BRIGHT  # 赤色太字
            color_end = Style.RESET_ALL
        
        # 平均値が最も低い場合：青色太字
        elif stat['average'] == min_average:
            color_start = Fore.BLUE + Style.BRIGHT  # 青色太字
            color_end = Style.RESET_ALL
        
        # 行を表示（色付けあり、列幅を指定）
        print(f"{color_start}{name:<{name_width}} {avg_str:<{avg_width}} {max_str:<{max_width}} {min_str:<{min_width}}{color_end}")
    
    # 表の終了区切り線を表示
    print("=" * 70)
    print("\n凡例:")
    print(f"{Fore.RED}{Style.BRIGHT}■{Style.RESET_ALL} 平均値が最も高い行")
    print(f"{Fore.BLUE}{Style.BRIGHT}■{Style.RESET_ALL} 平均値が最も低い行")


def main():
    """
    メイン処理を行う関数
    """
    try:
        # CSVファイルからデータを読み込む
        print("CSVファイルを読み込んでいます...")
        data = load_csv_data(CSV_FILE)
        print(f"データを {len(data)} 件読み込みました。")
        
        # 統計情報を計算
        print("統計情報を計算しています...")
        statistics = calculate_statistics(data)
        print(f"{len(statistics)} 名の参加者の統計情報を計算しました。")
        
        # 表形式で結果を表示
        display_table(statistics)
        
    except FileNotFoundError:
        # ファイルが見つからない場合のエラーメッセージ
        print(f"エラー: ファイル '{CSV_FILE}' が見つかりません。")
        print("ファイルが同じディレクトリにあることを確認してください。")
    
    except Exception as e:
        # その他のエラーメッセージ
        print(f"エラーが発生しました: {e}")


# プログラムのエントリーポイント（このファイルが直接実行された場合）
if __name__ == "__main__":
    main()
