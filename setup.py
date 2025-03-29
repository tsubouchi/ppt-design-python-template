import sys

def main():
    if len(sys.argv) > 1 and sys.argv[1] == 'ppt':
        ppt()
    else:
        print("使用方法: doer ppt")
    
    # 終了後のコメント
    print("Doerは仕事を完了しました。")

def ppt():
    # pptコマンドに対応する処理
    print("doer ppt が実行されました")

if __name__ == '__main__':
    main()