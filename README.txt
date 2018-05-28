
ASConsole


1. 概要

ASConsole (Active script console) はActiveScript 用インタラクティブシェルです。
ActiveScript に対応した複数のスクリプト言語を切り替えて利用することができます。


2. 機能

ScriptControl を使用して eval するための簡単な GUI フロントエンドです。
このツールは Windows 上で動作する GUI アプリケーションであって、導入が簡単なこ
と、動作が軽いことが特徴です。
Windows ユーザがソースコードを記述する際のちょっとした文法確認、動作確認を手軽
に行うことができます。パス名の加工のような文字列操作の動作確認や複雑な正規表現
を記述する際の動作検証の際に特に有用と考えます。
また、スクリプト言語の対話的実行環境としてではなく、履歴付きの関数電卓としても
便利に使うことができます。

以下の機能があります。

  - 式の評価
  - 文の実行
  - 複数のスクリプトエンジンの切り替え
  - 複数行入力
  - 組込み関数(COM情報表示、16進数変換とか)
  - Web ライブラリサポート(jQueryなど)


3. 動作環境

動作確認済みプラットフォームは以下のものです。

   Windows 7 Pro + IE11
   Windows 10 Pro + IE11


4. インストール/アンインストール

配布のアーカイブファイルを任意のディレクトリに展開してください。
特別な設定は必要ありません。

アプリケーションの起動は、展開したフォルダをエクスプローラで参照し、ファイル
ascon.hta をダブルクリックします。

アンインストールは展開されたファイルをすべて削除してください。


5. スクリプトエンジン

標準で使用可能なスクリプトエンジンは JavaScript、VBScript の2つです。
他のスクリプトエンジンを使用するためには対応する言語の ActiveScript に対応した
パッケージを別途導入してください。
Ruby、Perl、Python で動作を確認しています。

以下、対応している(過去に動作確認に使用した)パッケージを示します。
同じ言語でも開発元、パッケージが異なると使用できない可能性が高いです。

なお、作者は JavaScript、Ruby 以外はほとんど使いません。以下に記載されているも
のについても、だいぶ前に1度試しただけのものがほとんどで、情報が古くなっていま
す。
うまく動作しない可能性がありますので、注意してください。

  - Ruby
    [1.8.7]
    ActiveScriptRuby(1.8.7-p330 --with-winsock2 --enable-tcltk-stub) Microsoft Installer Package (1.8.7.36)
    ActiveRuby.msi
    http://www.artonx.org/data/asr/

    [2.4.0]
    Ruby-2.4.0(i386-mswin32 100) Microsoft Installer Package(2016-12-24)
    Ruby-2.4.msi
    http://www.artonx.org/data/asr/

    ※ActiveScriptRuby 2.4 (2016-12-24版) はパッケージのインストーラに問題があ
      るようで、インストールしただけでは ActiveScript が有効になりません。

      > regsvr32 RScript22.dll

      として明示的に COM サーバを登録する必要があるようです。
      また ActiveX の"初期化安全"フラグを明示的に設定する必要があります。
      初期化安全の設定は作者の環境では、以下のレジストリ設定の追加でうまくいき
      ました。
      [HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID\{39D7243A-AF85-
46BB-B70C-200EE1021A9B}\Implemented Categories\{7DD95802-9882-11CF-9FA9-
00AA006C42C4}]
      ※ これは Windows 10 64bit に 32bit ActiveScriptRuby をインストールした
         場合の例です。
         設定するレジストリは Windows や ActiveScriptRuby のバージョン、
         32bit/64bit の違いによっても異なる可能性があります。

    ※ActiveScriptRuby では、1.8.x、2.4.x の複数のバージョンを同時にインストー
      ルできますが、インストールする順番に注意してください。
      当アプリケーションではデフォルトでは後からインストールされたもののみを認
      識します。
      先にインストールしたバージョン、または両方のバージョンを同時に有効にする
      にはスクリプトの修正が必要になります。
      詳細は「8. 使用上の注意点など」の該当部分を参照してください。

  - Perl
    ActivePerl 5.16 Windows(x86)
    ActivePerl-5.16.3.1603-MSWin32-x86-296746.msi
    http://www.activestate.com/activeperl/

  - Python
    ActivePython 2.7 Windows(x86)
    ActivePython-2.7.2.5-win32-x86.msi
    http://www.activestate.com/activepython/
    のページから「Download ActivePython Community Edition」のリンクをクリック
    インストール後、以下も合わせてインストール
    Python for Windows extensions
    pywin32-218.win32-py2.7.exe
    http://sourceforge.net/projects/pywin32/?source=dlp

  - JavaScript
    ASConsole の JavaScript は IE7.0 相当の JavaScript エンジンで動作します。

    JavaScript は Windows 標準のものを使用しますが、Windows には複数の
    JavaScript エンジンが搭載されています。また複数の互換モードが存在します。

    ASConsole では、ScriptControl を使用して JavaScript エンジンを有効化します。
    ScriptControl を使用する場合、JScript.dll が使用されます。JScript.dll には
    5.8 の JScript エンジンが搭載されますが、ScriptControl 経由では 5.7 の互換
    モード(IE7.0相当)でしか動作しません。
    これはアプリケーションを HTA で記述しているための制限になります。

後述する COM 情報取得の組込み関数を使用する場合、ActiveScriptRuby のインストー
ルが前提となります。


6. 基本操作

6.1 スクリプト式評価

JavaScript を例にとり、式を評価してその結果を表示する手順を示します。
以下に例を示します。

> a = 10
10
> b = 20
20
> a * b
200

先頭に '>' のついている行がユーザの入力を表しています。
入力した行を逐次解釈し、その結果を表示します。
逐次解釈は内部で eval() を呼んでいるだけなので、可能なことは各言語の eval の機
能に依存します。


6.2 アプリケーションの操作

ユーザ入力欄に評価式、あるいは、システムコマンドを入力して、[実行]ボタンをク
リックします。
[実行]ボタンの代わりのキーボードのエンターキーも使用できます。
ユーザ入力欄にフォーカスがあるとき、[↑]または[↓]キーで過去の入力テキストの履
歴を呼び出すことができます。


5.3 複数行入力

「複数行」ボタンのクリック、または、[Ctrl]+'M' で複数行入力モードに切り替わります。
複数行入力モードではスクリプトの実行は [Ctrl]+[Enter] になります。
複数行入力モードで入力履歴を呼び出すには [Ctrl]+[↑]/[Ctrl]+[↓] または
[PageUp]/[PageDown] キーを使用します。


6.3 主なキー割り当て

  ----------------------------------------------------------------------------
                               1行編集モード         複数行編集モード
  ----------------------------------------------------------------------------
   編集モード切替え              Ctrl+'M'                Ctrl+'M'
   (単一行<=>複数行)
  ----------------------------------------------------------------------------
   入力履歴呼出し                [↑]/[↓]          Ctrl+[↑]/Ctrl+[↓]
                             [PageUp]/[PageDown]     [PageUp]/[PageDown]
  ----------------------------------------------------------------------------
   コマンド実行                   [Enter]               Ctrl+[Enter]
  ----------------------------------------------------------------------------


5.4 システムコマンド

入力文字列の先頭に # を付加すると、システムコマンドとして扱われます。
システムコマンドには以下のものがあります。

  #exit

    アプリケーションを終了します。

  #version

   バージョンを表示します。

  #clear

   コンソールを初期化します。

  #cs <scriptname>

   スクリプトエンジンを切り替えます。
   scriptname は ScriptControl 作成時に指定したスクリプト名、またはそのエイリ
   アスを指定します。
   以下の名前が指定可能です。ただし、事前に該当のスクリプトエンジンを有効化し
   ておく必要があります。

     JavaScript, javascript, js
     VBScript, vbscript, vbs
     RubyScript, rubyscript, ruby
     PerlScript, perlscript, perl
     Python, python, py

   なお、アプリケーション初回起動時には JavaScript が有効になっています。

  #!<command>

   Windows アプリケーションを起動します。
   #!notepad の指定でメモ帳が起動します。

  ##<n>

   入力テキストの履歴の一覧表示、または、特定の入力履歴の再実行を行います。

  #alias <keyword> <text>

   エイリアスを定義します。
   入力した文字列がエイリアス定義のキーワードに一致した場合、最初にエイリアス
   置換を行い、次にスクリプトの式または文の評価を行います。

   例えば、JavaScript と VBScript を頻繁に切り替えたい場合、

      #alias js #changeScript JavaScript
      #alias vbs #changeScript VBScript

   として定義しておくと、"js" または "vbs" の入力のみでスクリプトエンジンの切
   り替えを行うことができます。
   エイリアス定義を行う場合、キーワードおよびテキストのどちらの文字列も "(ダブ
   ルクォーテーション)などで囲まないように注意してください。ダブルクォーテー
   ションは正しく認識できません。

  #edit

   コンソールの内容をそのまま外部エディタで開きます。
   あらかじめ、#editor コマンドで、エディタ起動コマンドを設定しておく必要があり
   ます。

  #editor [<editor>]

   コンソールの内容を開くための外部エディタ起動コマンドの参照/設定を行います。
   クリップボード経由でテキストの受け渡しを行うため、外部エディタがクリップボー
   ドのテキストを開くことができる必要があります。

   以下にサクラエディタを登録する例を示します。
   起動直後に「貼付け」マクロを実行するよう、起動オプションを追加しています。

     #editor "C:\Program Files\sakura\sakura.exe" -M=S_Paste() -MTYPE=mac

   秀丸の場合、あらかじめ貼り付けのみを行うマクロ(ex. paste.mac)を作成しておき、
   以下のように指定します。

     #editor "C:\Program Files (x86)\Hidemaru\Hidemaru.exe" /xpaste.mac

  #enterexec [true|false]

   複数行編集モードにおいて、Enter キーの動作を設定します。

  ----------------------------------------------------------------------------
    enterexec = false (デフォルト)   [Enter] で改行挿入
                                     Ctrl+[Enter] で実行
  ----------------------------------------------------------------------------
    enterexec = true                 [Enter] で実行
                                     Ctrl+[Enter] で改行挿入
  ----------------------------------------------------------------------------


7. 組込み機能

便利機能としていくつかの組込み関数などを提供しています。
すべてのスクリプトから使えるものと、一部のスクリプトでのみ使える関数があります。

7.1 スクリプト間のオブジェクトの受け渡し

スクリプト間でオブジェクトの受け渡しを行うための組込み関数です。
直接他のスクリプトに渡すのではなく、共有領域を経由した間接的な受け渡しとなります。

  - オブジェクトを共通領域への退避
    exportVariable(<名前>, <オブジェクト>)

  - 共通領域からのオブジェクト取得
    importVariable(<名前>)


7.2 COM 情報取得

COM 情報取得のため、以下の組込み関数を提供します。
この機能を使用するためには ActiveScriptRuby が必要です。

  comInfo(obj [,name])   # インタフェース名の情報
                         # 名前が指定されている場合はメソッドまたはプロパティの情報
                         # (VBScript ではパラメタの省略ができないため、obj のみを
                         # パラメタとする場合に、name として Null を指定のこと)
  comMethods(obj)        # そのオブジェクトが持つメソッド一覧
  comProperties(obj)     # そのオブジェクトが持つプロパティ一覧

JavaScript を例にとり、実際の使い方を示します。
IE の制御そのものは通常の JavaScript の機能です。

> ie = new ActiveXObject('InternetExplorer.Application')
Microsoft Internet Explorer
> ie.visible = true
true
> ie.gohome()
undefined
> comInfo(ie)
[guid({D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}),
helpstring("Web Browser Interface for IE4.")]
Interface IWebBrowser2

> comInfo(ie.document)
[guid({3050F55F-98B5-11CF-BB82-00AA00BDCE0B}),
helpstring("")]
Dispatch DispHTMLDocument

> comMethods(ie)
QueryInterface(), AddRef(), Release(), GetTypeInfoCount(), GetTypeInfo(),
GetIDsOfNames(), Invoke(), GoBack(), GoForward(), GoHome(), GoSearch(),
Navigate(), Refresh(), Refresh2(), Stop(), Quit(), ClientToWindow(),
PutProperty(), GetProperty(), Navigate2(), QueryStatusWB(), ExecWB(),
ShowBrowserBar(), GetTypeInfoCount(), GetTypeInfo(), GetIDsOfNames(),
Invoke()

> comProperties(ie)
Application, Parent, Container, Document, TopLevelContainer, Type, Left,
Left=, Top, Top=, Width, Width=, Height, Height=, LocationName, LocationURL,
Busy, Name, HWND, FullName, Path, Visible, Visible=, StatusBar, StatusBar=,
StatusText, StatusText=, ToolBar, ToolBar=, MenuBar, MenuBar=, FullScreen,
FullScreen=, ReadyState, Offline, Offline=, Silent, Silent=,
RegisterAsBrowser, RegisterAsBrowser=, RegisterAsDropTarget,
RegisterAsDropTarget=, TheaterMode, TheaterMode=, AddressBar,
AddressBar=, Resizable, Resizable=

> comInfo(ie, "navigate")
[id(0x00000068), method, helpstring("Navigates to a URL or file.")]
VOID Navigate(
    [in] BSTR URL,
    [in, optional] VARIANT Flags,
    [in, optional] VARIANT TargetFrameName,
    [in, optional] VARIANT PostData,
    [in, optional] VARIANT Headers)


7.3 n進数変換

電卓用途としての便利機能です。スクリプトの機能とはあまり関係ありません。
以下の関数は JavaScript でのみ提供されます

  bin([n, ...]])   // 2進数変換   0b[01]+
  oct([n, ...]])   // 8進数変換   0[0-7]*
  dec([n, ...]])   // 10進数変換  0|0([1-9][0-9]*)
  hex([n, ...]])   // 16進数変換  0x[0-9A-F]+

使用例:
> hex(10, 20, 15*2)
0xA, 0x14, 0x1E
> dec()
10, 20, 30
> dec()
10, 20, 30
> oct()
012, 024, 036
> bin()
0b1010, 0b10100, 0b11110

パラメタなしで関数を呼び出すと直前の hex() or dec() or oct() or bin() へのパラ
メタを使用して再変換を行います。
また、hex() の代わりに単に hex としても動作します。
2進数の 0bXXX の表記は独自の出力用表現形式です。入力値として記述することはでき
ません。


7.4 ブラウザアタッチ

Web クライアントのスクリプト言語として JavaScript がよく使用されますが、最近で
は JavaScript をそのまま使用することは少なく、jQuery などのライブラリを前提と
することが多いようです。

これらのライブラリを利用可能とするため、IE で表示中の web ページにアタッチした
り、内部的にダミーのブラウザオブジェクトを生成したりするためのヘルパ関数を提供
します。

ブラウザのアタッチを行うと入力した式をローカル環境の eval() で評価する代わりに、
ブラウザの window.eval() を用いて式を評価するようになります。

以下に jQuery を例とした簡単な使い方の例を説明します。


7.4.1 ローカル環境でライブラリのみを使用する

- ライブラリファイルの準備

  ライブラリファイル(jquery-1.12.4.js など)を、配布サイト(http://jquery.com/) 
  から取得し、ASConsole フォルダに格納します。

- アタッチ

  > __attach(null, utils.load("jquery-1.12.4.js"))

  __attach() の第一パラメタに null を指定すると、実際の web ページの代わりに、
  システム内部に不可視のブラウザオブジェクトを1つ作成し、その window オブジェ
  クトにアタッチします。
  なお、不可視のブラウザオブジェクトは new ActiveXObject('htmlfile') として作
  成するものであり、使用される JavaScript エンジンは古いものになります。
  そのため、残念ながら jquery 2.x や 3.x は使用できません。

- 動作確認

  > $.fn.jquery
  1.12.4
  > $.ajax({url:"https://www.google.co.jp/",success:function(data){alert(data);}})


7.4.2 jQuery を使用した Web ページをアタッチする

- ASConsole 内から IE を開く

  > ie = __createIE()

- 表示された IE で、該当ページを開く

- 該当ページをアタッチ

  > __attach(ie.Document)

- 動作確認

  適当な要素を書き換えてみる

  > $("body").html("<h1>title</h1>body")


7.4.3 jQuery を使用していない Web ページをアタッチし、jQuery を強制的に挿入する

- ASConsole 内から IE を開く

  > ie = __createIE()

- 表示された IE で、該当ページを開く

  > ie.navigate("https://www.example.com/")

- 該当ページをアタッチし、同時に jQuery を挿入

  > __attach(ie.Document, utils.load("jquery-1.12.4.js"))

- 動作確認

  適当な要素を書き換えてみる

  > $("body").html("<h1>title</h1>body")


7.4.4 デタッチ

アタッチは該当ページ毎に行います。
アタッチをするということは、現在アタッチしているブラウザのページ内で作業してい
るイメージであり、ブラウザに強く依存している状態です。
アタッチした状態で該当のブラウザを終了させると、すべての式評価ができなくなって
しまいます。
該当ページでの作業が終わったら、必ずデタッチを行ってください。

  > __detach()

デタッチを行うことで、アタッチ前の状態に環境を戻します。
なお、ブラウザを終了してしまい、すべての式評価ができない状態となった場合、

  [ERROR] 462: リモート サーバ マシンが存在しないか、利用できません。

のようなエラーメッセージが表示されます。
このエラーが表示された場合、自動でデタッチが行われますが、万が一、うまくデタッ
チされない場合は明示的にデタッチコマンドを入力してください。
すべての式評価ができない状態であっても __detach() コマンドだけは受け付けるよう
になっています。

ASConsole 自体を終了すると、すべてのアタッチ状態は完全に初期化されます。
デタッチをせずに、ASConsole を終了させても問題はありません。


8. ユーザライブラリ

ASConsole ではユーザが関数などを自由に定義できます。(JavaScript のみが対象です。)
ASConsole を展開したフォルダに lib フォルダがありますが、lib フォルダ内に .js
ファイルを作成すると ASConsole 起動時に自動的に取り込まれ、スクリプト内に定義
した自作の関数などが使用できるようになります。
lib/ipcalc.js はライブラリのサンプルです。


9. VBScript/Python における動作

VBScript では式と文が区別されます。
ASConsole では入力をすべて文として扱いますので、単に変数 a の内容を見る場合でも、
print a のように文になるように構成してください。

> a = 123         # 変数 a に 123 を代入

> println a       # 変数 a の内容を表示して改行
123
> a               # 単に a と記述すると、エラーになる
[ERROR] 13: 型が一致しません。
> :a              # a を式として評価
=>123

※ print/println は ASConsole が提供するコンソール出力用ヘルパ関数です。

Python も VBScript と同様、文と式が区別されるため、入力は文となるよう構成して
ください。

なお、VBScript/Python のいずれも変数の内容を簡単に確認する用途のために、特殊構
文をサポートします。
  ':'(コロン) + 式
のように先頭に ':' を付加して実行すると、':' を取り除いた残りのテキストを
Execute/exec ではなく、Eval/eval で評価し、その結果を表示します。


10. スクリプトの呼び出し方法について

入力式の評価は ScriptControl#Run でスクリプト毎のヘルパ関数(execEval)を呼び出
し、ヘルパ関数内で各言語毎の eval を呼び出します。VBScript/Python ではヘルパ関
数内で Execute/exec を呼び出し、文の実行を行います。
ScriptControl#Eval ではなく、あえてヘルパ関数を使用するのは、スクリプトの例外
を呼び出し元で正しく拾えないスクリプトエンジンがあること、拾えても呼び出し元で
はエラー情報を適切に抽出できない可能性があること、直前の結果の記憶などの追加処
理をヘルパ関数で処理するなど、複数の理由によります。
現状では例外を言語間の呼び出しで伝播させず、必ず、呼び出し先のスクリプト内で
catch して完結させる方式としています。そもそも言語間にまたがる例外の伝播が可能
かどうかが分かっていないこともあります。呼び出し元には文字列レベルの情報だけを
返します。

ヘルパ関数を経由して eval するのが基本的な使用方法となりますが、それでも
ScriptControl#ExecuteStatement が必要となる局面があります。また、
ScriptControl#Eval の挙動を確認したい場合もあると思います。(スクリプトの言語仕
様によりますが、モジュールのインポートや関数定義など eval ではエラーとなる
場合でも ScriptControl#ExecuteStatement で処理できる可能性があります。)
そのような場合、評価式の先頭に特定のプレフィクスを付加して実行することで、強制
的に特定の呼び出し方法を選択できます。

  プレフィクス 動作
 ----------------------------------------------------------------------------
  (なし)    ScriptControl#Run (ヘルパ関数経由で eval、VBScript/PythonはExecute/exec)
  :         ScriptControl#Run (ヘルパ関数経由で Eval/eval、VBScript/Pythonのみ)
  ;         ScriptControl#ExecuteStatement (直接呼出し)
  ;;        ScriptControl#Eval (直接呼出し)
  <         ScriptControl#AddCode (直接呼出し) スクリプトのテキストを入力とする
  <<        ScriptControl#AddCode (直接呼出し) スクリプトのファイル名を入力とする
  #         システムコマンド
 ----------------------------------------------------------------------------


11. 使用上の注意点など

ヘルパ関数内で eval するため、そのままでは入力式の実行コンテキストがヘルパ関数
の関数内になることに注意してください。
そのことにとどまらず、スコープや名前空間の考え方が各言語毎に異なりますが、各言
語毎の詳細は作者の理解を超えているため、ソースを見て判断してください。

次の言語は小細工によって、実行コンテキストをグローバル風に修正しています。
Ruby では、eval で TOPLEVEL_BINDING を指定しています。
Python では exec で globals() を指定しています。

JavaScript/VBScript 以外のスクリプト言語では、標準出力に対する操作が可能です。
このうち、標準出力、標準エラー出力については出力先を切り替えることで当アプリケー
ションの GUI コンソールへの表示をサポートします。
当アプリケーションでは標準入力はサポートしません。標準入力は正常に動作すること
はなく、スクリプトエンジンによってはアプリケーションエラーなどの致命的なエラー
を引き起こす可能性もあるため、該当機能を使用しないように注意してください。
標準出力の対応は暫定です。特に日本語の扱いについてはまだ不具合の残っている可能
性があります。

ActiveScriptRuby 1.8.7、2.4.x は同時にシステムに登録できますが、
ASConsole 起動時には最後にインストールしたバージョンのみが有効になります。
両方の Ruby を同時に認識させるためには以下の説明に従い、スクリプトを修正して
ください。

    (ProgID)
   スクリプト名   説明
 -----------------------------------------------------------------------
  RubyScript     ActiveScriptRuby の代表名(あとからインストールしたバージョンで上書き)
  RubyScript.1   ActiveScriptRuby 1.8.7 の個別名
  RubyScript.2.4 ActiveScriptRuby 2.4.x の個別名

ActiveScriptRuby をインストールすると、上記名前でスクリプトエンジンが登録され
ます。当アプリケーションの初期状態では上記の代表名のみが有効になっており、その
名前でスクリプトエンジンが選択されます。そのため、ASConsole を起動すると、1.8.7、
2.4 のうち、後からインストールしたバージョンが自動的に選択されます。
どのバージョンが選択されているかの確認するには定数 RUBY_VERSION を参照してくだ
さい。以下に例を示します。

> RUBY_VERSION
2.4.0

両方のバージョンの ActiveScriptRuby をインストールしたシステム上で、その両方の
スクリプトエンジンを有効にするためには、scriptmanager.js 内を以下のように修正し
ます。(コメントアウトされている箇所のコメントをはずすだけ)

...
	[RubyScript,    ["eval.rb"], "rb"],
	[RubyScript18,  ["eval.rb"], "rb18"],
	[RubyScript24,  ["eval.rb"], "rb24"],
...

サフィックスなしの RubyScript は内部的に COM 操作関数が使用します。
各バージョンを認識させる場合でも、デフォルトの RubyScript の行は有効なままにし
ておいてください。

ruby で JavaScript のメソッドに nil を渡しても null と解釈されないようです。
(JavaScript 側の typeof(value) が "unknown" になる)
なお、JavaScript側から RubyScriptのメソッドを呼んだ場合、null は自動的に nil
に変換されます。

ちなみに、他の言語の COM インタフェースにおける null 表現は以下となります。
(たぶん)

  ----------------------------------------------------------------------
   JavaScript  null
   VBScript    Null
   RubyScript  ?
   PerlScript  Win32::OLE::Variant->new(VT_NULL) または Variant(VT_NULL)
   Python      None
  ----------------------------------------------------------------------
   何が null かという問題がありますが(VT_NULL or VT_EMPTY など)、ここでは
   ASConsole のホストであり、言語間連携のハブとなる JavaScript のメソッドを呼
   び出したときの見え方を基準としています。(インタフェースオブジェクトを
   JavaScript で実装している都合のため、JavaScript から null と見えるも
   のを基準に考えています。)

文字コード、特に日本語の扱いは全体的に作者の理解が不十分です。以下は現状の覚書
になります。
HTA、直接取り込む.jsファイル、スクリプトファイルは UTF-8として扱っています。
たとえ、コメントであっても、Shift_JISで日本語などを記述するとスクリプトの解析
に失敗する可能性があります。

実行時の文字コードについて Ruby 1.8.7 では $KCODE='SJIS' を動的に指定していま
す。1.9以降は eval されるパラメタ文字列を強制的に .force_encoding('Windows-31J')
して日本語として認識させています。この対処は完全ではなく、例えば
importVariable/exportVariable で渡される情報は正しく認識できない可能性がありま
す。

Python では日本語文字列には u を必ず付加してください。(ex. u"日本語")
u をつけておく限り、簡単な日本語処理では問題は発生しないようです。

Perl における文字コードの扱いは正直よく分かっておらず、現状ではシステム側の考
慮はありません。
しかし、単純な日本語文字列を表示するだけであれば特に問題はないようです。

まれに表示が固まってしまうことがあります。(アプリ自体がハングするわけではない)
文字化け等でコンソール領域に特殊な文字を表示させた場合に問題が発生する可能性が
高いようです。そのような場合は #clear コマンドでコンソールを初期化してみてくだ
さい。それでも解決しない場合、アプリケーションを終了させてください。この問題は
調査中です。

comInfo(obj, name) でプロパティ情報を取得するとき、PROPERTYPUT の型が表示さ
れていません。これは現状の仕様です。(ruby の win32ole を使って PROPERTYPUT
の型情報を取得する方法が分からないため)

JavaScript でインスタンスメソッドを作成し、Ruby からオブジェクト経由でそのメソッ
ドを呼び出すとき、パラメタがあれば正しくメソッド呼び出しを行えますが、パラメタ
がないと関数自体の定義情報(テキスト?)を取得してしまうようです(JavaScript では
foo() と foo は別物だが、ruby ではメソッドとプロパティの区別がつかず、パラメタ
なしの場合、常に foo すなわちプロパティとして解釈してしまう?)。
正しい使い方が存在するのかも知れませんが、たとえパラメタが存在しなくてもダミー
のパラメタを付加することでとりあえず回避はできます。(getText() に対して
getText('dummy') などとする)これは Perl でも同じようです。
逆に Python ではパラメタなしのメソッドを正しく処理できるうえ、余分なパラメタを
つけるとエラーとなってメソッド呼び出しに失敗します。

スクリプトエンジンの追加はヘルパ関数(execEvalなどいくつかの必須関数)を含むスク
リプトテキストを用意して scriptmanager.js のテーブルを修正することで行います。
他のスクリプトの実装を参考にしてください。
最低限は execEval があればよいはずです。

ASConsole は複数起動できません。これは、コマンド履歴などのユーザ設定情報を複数
のインスタンスで共有することが難しいためです。
それでもあえて、複数起動したい場合には、ascon.hta を以下のように書き換えてくだ
さい。複数起動モードでアプリケーションが起動します。

   hta:applicationタグ内
      SINGLEINSTANCE="yes"   =>   SINGLEINSTANCE="no"

複数起動モードでは、コマンド履歴、エイリアスおよび選択中のスクリプトエンジンの
保存を行いません。そのため、あえて複数起動させたい場合は、普段使用する標準設定
のアプリはそのまま残し、フォルダ全体を別フォルダにコピーしてから、コピーしたも
のを編集することをお勧めします。


12. 作成者

kanegon


13. ライセンス

MIT ライセンスとして公開します。
詳細は同梱の LICENSE を参照してください。


14. 修正履歴

ver0.01  2003-04-26
    新規作成 (jsevalとして)

ver1.00  2009-03-24
    ASConsole に名称変更
    スクリプトの呼び出しを ScriptControl 経由に変更
    JavaScript、VBScript、Ruby、Perl、Python に対応

ver1.07  2011-06-01
    ブラウザアタッチ機能追加(jQuery 対応)

ver2.00  2018-05-28
    複数行編集機能のインライン化
    設定/履歴ファイル(ascon.dat)のJSON化、保存先変更
    MIT ランセンスに変更
