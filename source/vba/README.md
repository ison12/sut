# ディレクトリ構成

    vba
      |
      + src
        + resource ... リソース（画像、設定ファイル、メタ定義取得SQLファイルなど）
        + manual   ... マニュアル
        + test     ... テストフォルダ
        + Sut.xlam ... Sut.xlam本体を格納
      + src_export ... Sut.xlamからエクスポートしたモジュールファイル
      |
      + tool
        + vbac.wsf VBAファイルのエクスポート・インポートツール
            http://igeta-diary.blogspot.com/2014/03/what-is-vbac.html
            https://github.com/vbaidiot/Ariawase
        + vba_module_export.bat - VBAモジュールをエクスポートするバッチ
