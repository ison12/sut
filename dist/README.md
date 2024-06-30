# ディレクトリ構成

    dist
      |
      + Sut     ... リリースフォルダ
      + Sut.zip ... リリースファイル
      + tool
        + logs ... ログフォルダ
        + zip  ... ZIP圧縮するためのツール
        + 00_Release_collect.bat ... リリースのための資材をコピーするツール
        + 01_Release_archive.bat ... ZIP圧縮するツール
        + 10_Release_vba.bat ... リリースバッチのサブモジュール
        + 20_Release_cpp.bat ... リリースバッチのサブモジュール

# リリースバッチの使い方

1. 00_Release_collect.bat を実行。
  本バッチ実行後に、Sut.xlamを開いて条件付きコンパイル引数のDEBUG_MODEを除去する。
1. 01_Release_archive.bat を実行。
  本バッチ実行後に、Sut.zipが生成されるのでこちらをリリースする。

# GitHubへのリリース

1. リリースバッチによって、Sut.zipを生成しておくこと。
1. GitHubの当該ページでリリースを実施。
  https://github.com/ison12/sut/releases
  ※リリース時に、Sut.zipファイルを添付
