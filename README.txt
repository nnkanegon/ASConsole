
ASConsole


1. �T�v

ASConsole (Active script console) ��ActiveScript �p�C���^���N�e�B�u�V�F���ł��B
ActiveScript �ɑΉ����������̃X�N���v�g�����؂�ւ��ė��p���邱�Ƃ��ł��܂��B


2. �@�\

ScriptControl ���g�p���� eval ���邽�߂̊ȒP�� GUI �t�����g�G���h�ł��B
���̃c�[���� Windows ��œ��삷�� GUI �A�v���P�[�V�����ł����āA�������ȒP�Ȃ�
�ƁA���삪�y�����Ƃ������ł��B
Windows ���[�U���\�[�X�R�[�h���L�q����ۂ̂�����Ƃ������@�m�F�A����m�F����y
�ɍs�����Ƃ��ł��܂��B�p�X���̉��H�̂悤�ȕ����񑀍�̓���m�F�╡�G�Ȑ��K�\��
���L�q����ۂ̓��쌟�؂̍ۂɓ��ɗL�p�ƍl���܂��B
�܂��A�X�N���v�g����̑Θb�I���s���Ƃ��Ăł͂Ȃ��A����t���̊֐��d��Ƃ��Ă�
�֗��Ɏg�����Ƃ��ł��܂��B

�ȉ��̋@�\������܂��B

  - ���̕]��
  - ���̎��s
  - �����̃X�N���v�g�G���W���̐؂�ւ�
  - �����s����
  - �g���݊֐�(COM���\���A16�i���ϊ��Ƃ�)
  - Web ���C�u�����T�|�[�g(jQuery�Ȃ�)


3. �����

����m�F�ς݃v���b�g�t�H�[���͈ȉ��̂��̂ł��B

   Windows 7 Pro + IE11
   Windows 10 Pro + IE11


4. �C���X�g�[��/�A���C���X�g�[��

�z�z�̃A�[�J�C�u�t�@�C����C�ӂ̃f�B���N�g���ɓW�J���Ă��������B
���ʂȐݒ�͕K�v����܂���B

�A�v���P�[�V�����̋N���́A�W�J�����t�H���_���G�N�X�v���[���ŎQ�Ƃ��A�t�@�C��
ascon.hta ���_�u���N���b�N���܂��B

�A���C���X�g�[���͓W�J���ꂽ�t�@�C�������ׂč폜���Ă��������B


5. �X�N���v�g�G���W��

�W���Ŏg�p�\�ȃX�N���v�g�G���W���� JavaScript�AVBScript ��2�ł��B
���̃X�N���v�g�G���W�����g�p���邽�߂ɂ͑Ή����錾��� ActiveScript �ɑΉ�����
�p�b�P�[�W��ʓr�������Ă��������B
Ruby�APerl�APython �œ�����m�F���Ă��܂��B

�ȉ��A�Ή����Ă���(�ߋ��ɓ���m�F�Ɏg�p����)�p�b�P�[�W�������܂��B
��������ł��J�����A�p�b�P�[�W���قȂ�Ǝg�p�ł��Ȃ��\���������ł��B

�Ȃ��A��҂� JavaScript�ARuby �ȊO�͂قƂ�ǎg���܂���B�ȉ��ɋL�ڂ���Ă����
�̂ɂ��Ă��A�����ԑO��1�x�����������̂��̂��قƂ�ǂŁA��񂪌Â��Ȃ��Ă���
���B
���܂����삵�Ȃ��\��������܂��̂ŁA���ӂ��Ă��������B

  - Ruby
    [1.8.7]
    ActiveScriptRuby(1.8.7-p330 --with-winsock2 --enable-tcltk-stub) Microsoft Installer Package (1.8.7.36)
    ActiveRuby.msi
    http://www.artonx.org/data/asr/

    [2.4.0]
    Ruby-2.4.0(i386-mswin32 100) Microsoft Installer Package(2016-12-24)
    Ruby-2.4.msi
    http://www.artonx.org/data/asr/

    ��ActiveScriptRuby 2.4 (2016-12-24��) �̓p�b�P�[�W�̃C���X�g�[���ɖ�肪��
      ��悤�ŁA�C���X�g�[�����������ł� ActiveScript ���L���ɂȂ�܂���B

      > regsvr32 RScript22.dll

      �Ƃ��Ė����I�� COM �T�[�o��o�^����K�v������悤�ł��B
      �܂� ActiveX ��"���������S"�t���O�𖾎��I�ɐݒ肷��K�v������܂��B
      ���������S�̐ݒ�͍�҂̊��ł́A�ȉ��̃��W�X�g���ݒ�̒ǉ��ł��܂�����
      �܂����B
      [HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID\{39D7243A-AF85-
46BB-B70C-200EE1021A9B}\Implemented Categories\{7DD95802-9882-11CF-9FA9-
00AA006C42C4}]
      �� ����� Windows 10 64bit �� 32bit ActiveScriptRuby ���C���X�g�[������
         �ꍇ�̗�ł��B
         �ݒ肷�郌�W�X�g���� Windows �� ActiveScriptRuby �̃o�[�W�����A
         32bit/64bit �̈Ⴂ�ɂ���Ă��قȂ�\��������܂��B

    ��ActiveScriptRuby �ł́A1.8.x�A2.4.x �̕����̃o�[�W�����𓯎��ɃC���X�g�[
      ���ł��܂����A�C���X�g�[�����鏇�Ԃɒ��ӂ��Ă��������B
      ���A�v���P�[�V�����ł̓f�t�H���g�ł͌ォ��C���X�g�[�����ꂽ���݂̂̂�F
      �����܂��B
      ��ɃC���X�g�[�������o�[�W�����A�܂��͗����̃o�[�W�����𓯎��ɗL���ɂ���
      �ɂ̓X�N���v�g�̏C�����K�v�ɂȂ�܂��B
      �ڍׂ́u8. �g�p��̒��ӓ_�Ȃǁv�̊Y���������Q�Ƃ��Ă��������B

  - Perl
    ActivePerl 5.16 Windows(x86)
    ActivePerl-5.16.3.1603-MSWin32-x86-296746.msi
    http://www.activestate.com/activeperl/

  - Python
    ActivePython 2.7 Windows(x86)
    ActivePython-2.7.2.5-win32-x86.msi
    http://www.activestate.com/activepython/
    �̃y�[�W����uDownload ActivePython Community Edition�v�̃����N���N���b�N
    �C���X�g�[����A�ȉ������킹�ăC���X�g�[��
    Python for Windows extensions
    pywin32-218.win32-py2.7.exe
    http://sourceforge.net/projects/pywin32/?source=dlp

  - JavaScript
    ASConsole �� JavaScript �� IE7.0 ������ JavaScript �G���W���œ��삵�܂��B

    JavaScript �� Windows �W���̂��̂��g�p���܂����AWindows �ɂ͕�����
    JavaScript �G���W�������ڂ���Ă��܂��B�܂������̌݊����[�h�����݂��܂��B

    ASConsole �ł́AScriptControl ���g�p���� JavaScript �G���W����L�������܂��B
    ScriptControl ���g�p����ꍇ�AJScript.dll ���g�p����܂��BJScript.dll �ɂ�
    5.8 �� JScript �G���W�������ڂ���܂����AScriptControl �o�R�ł� 5.7 �̌݊�
    ���[�h(IE7.0����)�ł������삵�܂���B
    ����̓A�v���P�[�V������ HTA �ŋL�q���Ă��邽�߂̐����ɂȂ�܂��B

��q���� COM ���擾�̑g���݊֐����g�p����ꍇ�AActiveScriptRuby �̃C���X�g�[
�����O��ƂȂ�܂��B


6. ��{����

6.1 �X�N���v�g���]��

JavaScript ���ɂƂ�A����]�����Ă��̌��ʂ�\������菇�������܂��B
�ȉ��ɗ�������܂��B

> a = 10
10
> b = 20
20
> a * b
200

�擪�� '>' �̂��Ă���s�����[�U�̓��͂�\���Ă��܂��B
���͂����s�𒀎����߂��A���̌��ʂ�\�����܂��B
�������߂͓����� eval() ���Ă�ł��邾���Ȃ̂ŁA�\�Ȃ��Ƃ͊e����� eval �̋@
�\�Ɉˑ����܂��B


6.2 �A�v���P�[�V�����̑���

���[�U���͗��ɕ]�����A���邢�́A�V�X�e���R�}���h����͂��āA[���s]�{�^�����N
���b�N���܂��B
[���s]�{�^���̑���̃L�[�{�[�h�̃G���^�[�L�[���g�p�ł��܂��B
���[�U���͗��Ƀt�H�[�J�X������Ƃ��A[��]�܂���[��]�L�[�ŉߋ��̓��̓e�L�X�g�̗�
�����Ăяo�����Ƃ��ł��܂��B


5.3 �����s����

�u�����s�v�{�^���̃N���b�N�A�܂��́A[Ctrl]+'M' �ŕ����s���̓��[�h�ɐ؂�ւ��܂��B
�����s���̓��[�h�ł̓X�N���v�g�̎��s�� [Ctrl]+[Enter] �ɂȂ�܂��B
�����s���̓��[�h�œ��͗������Ăяo���ɂ� [Ctrl]+[��]/[Ctrl]+[��] �܂���
[PageUp]/[PageDown] �L�[���g�p���܂��B


6.3 ��ȃL�[���蓖��

  ----------------------------------------------------------------------------
                               1�s�ҏW���[�h         �����s�ҏW���[�h
  ----------------------------------------------------------------------------
   �ҏW���[�h�ؑւ�              Ctrl+'M'                Ctrl+'M'
   (�P��s<=>�����s)
  ----------------------------------------------------------------------------
   ���͗����ďo��                [��]/[��]          Ctrl+[��]/Ctrl+[��]
                             [PageUp]/[PageDown]     [PageUp]/[PageDown]
  ----------------------------------------------------------------------------
   �R�}���h���s                   [Enter]               Ctrl+[Enter]
  ----------------------------------------------------------------------------


5.4 �V�X�e���R�}���h

���͕�����̐擪�� # ��t������ƁA�V�X�e���R�}���h�Ƃ��Ĉ����܂��B
�V�X�e���R�}���h�ɂ͈ȉ��̂��̂�����܂��B

  #exit

    �A�v���P�[�V�������I�����܂��B

  #version

   �o�[�W������\�����܂��B

  #clear

   �R���\�[�������������܂��B

  #cs <scriptname>

   �X�N���v�g�G���W����؂�ւ��܂��B
   scriptname �� ScriptControl �쐬���Ɏw�肵���X�N���v�g���A�܂��͂��̃G�C��
   �A�X���w�肵�܂��B
   �ȉ��̖��O���w��\�ł��B�������A���O�ɊY���̃X�N���v�g�G���W����L������
   �Ă����K�v������܂��B

     JavaScript, javascript, js
     VBScript, vbscript, vbs
     RubyScript, rubyscript, ruby
     PerlScript, perlscript, perl
     Python, python, py

   �Ȃ��A�A�v���P�[�V��������N�����ɂ� JavaScript ���L���ɂȂ��Ă��܂��B

  #!<command>

   Windows �A�v���P�[�V�������N�����܂��B
   #!notepad �̎w��Ń��������N�����܂��B

  ##<n>

   ���̓e�L�X�g�̗����̈ꗗ�\���A�܂��́A����̓��͗����̍Ď��s���s���܂��B

  #alias <keyword> <text>

   �G�C���A�X���`���܂��B
   ���͂��������񂪃G�C���A�X��`�̃L�[���[�h�Ɉ�v�����ꍇ�A�ŏ��ɃG�C���A�X
   �u�����s���A���ɃX�N���v�g�̎��܂��͕��̕]�����s���܂��B

   �Ⴆ�΁AJavaScript �� VBScript ��p�ɂɐ؂�ւ������ꍇ�A

      #alias js #changeScript JavaScript
      #alias vbs #changeScript VBScript

   �Ƃ��Ē�`���Ă����ƁA"js" �܂��� "vbs" �̓��݂͂̂ŃX�N���v�g�G���W���̐�
   ��ւ����s�����Ƃ��ł��܂��B
   �G�C���A�X��`���s���ꍇ�A�L�[���[�h����уe�L�X�g�̂ǂ���̕������ "(�_�u
   ���N�H�[�e�[�V����)�Ȃǂň͂܂Ȃ��悤�ɒ��ӂ��Ă��������B�_�u���N�H�[�e�[
   �V�����͐������F���ł��܂���B

  #edit

   �R���\�[���̓��e�����̂܂܊O���G�f�B�^�ŊJ���܂��B
   ���炩���߁A#editor �R�}���h�ŁA�G�f�B�^�N���R�}���h��ݒ肵�Ă����K�v������
   �܂��B

  #editor [<editor>]

   �R���\�[���̓��e���J�����߂̊O���G�f�B�^�N���R�}���h�̎Q��/�ݒ���s���܂��B
   �N���b�v�{�[�h�o�R�Ńe�L�X�g�̎󂯓n�����s�����߁A�O���G�f�B�^���N���b�v�{�[
   �h�̃e�L�X�g���J�����Ƃ��ł���K�v������܂��B

   �ȉ��ɃT�N���G�f�B�^��o�^�����������܂��B
   �N������Ɂu�\�t���v�}�N�������s����悤�A�N���I�v�V������ǉ����Ă��܂��B

     #editor "C:\Program Files\sakura\sakura.exe" -M=S_Paste() -MTYPE=mac

   �G�ۂ̏ꍇ�A���炩���ߓ\��t���݂̂��s���}�N��(ex. paste.mac)���쐬���Ă����A
   �ȉ��̂悤�Ɏw�肵�܂��B

     #editor "C:\Program Files (x86)\Hidemaru\Hidemaru.exe" /xpaste.mac

  #enterexec [true|false]

   �����s�ҏW���[�h�ɂ����āAEnter �L�[�̓����ݒ肵�܂��B

  ----------------------------------------------------------------------------
    enterexec = false (�f�t�H���g)   [Enter] �ŉ��s�}��
                                     Ctrl+[Enter] �Ŏ��s
  ----------------------------------------------------------------------------
    enterexec = true                 [Enter] �Ŏ��s
                                     Ctrl+[Enter] �ŉ��s�}��
  ----------------------------------------------------------------------------


7. �g���݋@�\

�֗��@�\�Ƃ��Ă������̑g���݊֐��Ȃǂ�񋟂��Ă��܂��B
���ׂẴX�N���v�g����g������̂ƁA�ꕔ�̃X�N���v�g�ł̂ݎg����֐�������܂��B

7.1 �X�N���v�g�Ԃ̃I�u�W�F�N�g�̎󂯓n��

�X�N���v�g�ԂŃI�u�W�F�N�g�̎󂯓n�����s�����߂̑g���݊֐��ł��B
���ڑ��̃X�N���v�g�ɓn���̂ł͂Ȃ��A���L�̈���o�R�����ԐړI�Ȏ󂯓n���ƂȂ�܂��B

  - �I�u�W�F�N�g�����ʗ̈�ւ̑ޔ�
    exportVariable(<���O>, <�I�u�W�F�N�g>)

  - ���ʗ̈悩��̃I�u�W�F�N�g�擾
    importVariable(<���O>)


7.2 COM ���擾

COM ���擾�̂��߁A�ȉ��̑g���݊֐���񋟂��܂��B
���̋@�\���g�p���邽�߂ɂ� ActiveScriptRuby ���K�v�ł��B

  comInfo(obj [,name])   # �C���^�t�F�[�X���̏��
                         # ���O���w�肳��Ă���ꍇ�̓��\�b�h�܂��̓v���p�e�B�̏��
                         # (VBScript �ł̓p�����^�̏ȗ����ł��Ȃ����߁Aobj �݂̂�
                         # �p�����^�Ƃ���ꍇ�ɁAname �Ƃ��� Null ���w��̂���)
  comMethods(obj)        # ���̃I�u�W�F�N�g�������\�b�h�ꗗ
  comProperties(obj)     # ���̃I�u�W�F�N�g�����v���p�e�B�ꗗ

JavaScript ���ɂƂ�A���ۂ̎g�����������܂��B
IE �̐��䂻�̂��̂͒ʏ�� JavaScript �̋@�\�ł��B

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


7.3 n�i���ϊ�

�d��p�r�Ƃ��Ă֗̕��@�\�ł��B�X�N���v�g�̋@�\�Ƃ͂��܂�֌W����܂���B
�ȉ��̊֐��� JavaScript �ł̂ݒ񋟂���܂�

  bin([n, ...]])   // 2�i���ϊ�   0b[01]+
  oct([n, ...]])   // 8�i���ϊ�   0[0-7]*
  dec([n, ...]])   // 10�i���ϊ�  0|0([1-9][0-9]*)
  hex([n, ...]])   // 16�i���ϊ�  0x[0-9A-F]+

�g�p��:
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

�p�����^�Ȃ��Ŋ֐����Ăяo���ƒ��O�� hex() or dec() or oct() or bin() �ւ̃p��
���^���g�p���čĕϊ����s���܂��B
�܂��Ahex() �̑���ɒP�� hex �Ƃ��Ă����삵�܂��B
2�i���� 0bXXX �̕\�L�͓Ǝ��̏o�͗p�\���`���ł��B���͒l�Ƃ��ċL�q���邱�Ƃ͂ł�
�܂���B


7.4 �u���E�U�A�^�b�`

Web �N���C�A���g�̃X�N���v�g����Ƃ��� JavaScript ���悭�g�p����܂����A�ŋ߂�
�� JavaScript �����̂܂܎g�p���邱�Ƃ͏��Ȃ��AjQuery �Ȃǂ̃��C�u������O���
���邱�Ƃ������悤�ł��B

�����̃��C�u�����𗘗p�\�Ƃ��邽�߁AIE �ŕ\������ web �y�[�W�ɃA�^�b�`����
��A�����I�Ƀ_�~�[�̃u���E�U�I�u�W�F�N�g�𐶐������肷�邽�߂̃w���p�֐����
���܂��B

�u���E�U�̃A�^�b�`���s���Ɠ��͂����������[�J������ eval() �ŕ]���������ɁA
�u���E�U�� window.eval() ��p���Ď���]������悤�ɂȂ�܂��B

�ȉ��� jQuery ���Ƃ����ȒP�Ȏg�����̗��������܂��B


7.4.1 ���[�J�����Ń��C�u�����݂̂��g�p����

- ���C�u�����t�@�C���̏���

  ���C�u�����t�@�C��(jquery-1.12.4.js �Ȃ�)���A�z�z�T�C�g(http://jquery.com/) 
  ����擾���AASConsole �t�H���_�Ɋi�[���܂��B

- �A�^�b�`

  > __attach(null, utils.load("jquery-1.12.4.js"))

  __attach() �̑��p�����^�� null ���w�肷��ƁA���ۂ� web �y�[�W�̑���ɁA
  �V�X�e�������ɕs���̃u���E�U�I�u�W�F�N�g��1�쐬���A���� window �I�u�W�F
  �N�g�ɃA�^�b�`���܂��B
  �Ȃ��A�s���̃u���E�U�I�u�W�F�N�g�� new ActiveXObject('htmlfile') �Ƃ��č�
  ��������̂ł���A�g�p����� JavaScript �G���W���͌Â����̂ɂȂ�܂��B
  ���̂��߁A�c�O�Ȃ��� jquery 2.x �� 3.x �͎g�p�ł��܂���B

- ����m�F

  > $.fn.jquery
  1.12.4
  > $.ajax({url:"https://www.google.co.jp/",success:function(data){alert(data);}})


7.4.2 jQuery ���g�p���� Web �y�[�W���A�^�b�`����

- ASConsole ������ IE ���J��

  > ie = __createIE()

- �\�����ꂽ IE �ŁA�Y���y�[�W���J��

- �Y���y�[�W���A�^�b�`

  > __attach(ie.Document)

- ����m�F

  �K���ȗv�f�����������Ă݂�

  > $("body").html("<h1>title</h1>body")


7.4.3 jQuery ���g�p���Ă��Ȃ� Web �y�[�W���A�^�b�`���AjQuery �������I�ɑ}������

- ASConsole ������ IE ���J��

  > ie = __createIE()

- �\�����ꂽ IE �ŁA�Y���y�[�W���J��

  > ie.navigate("https://www.example.com/")

- �Y���y�[�W���A�^�b�`���A������ jQuery ��}��

  > __attach(ie.Document, utils.load("jquery-1.12.4.js"))

- ����m�F

  �K���ȗv�f�����������Ă݂�

  > $("body").html("<h1>title</h1>body")


7.4.4 �f�^�b�`

�A�^�b�`�͊Y���y�[�W���ɍs���܂��B
�A�^�b�`������Ƃ������Ƃ́A���݃A�^�b�`���Ă���u���E�U�̃y�[�W���ō�Ƃ��Ă�
��C���[�W�ł���A�u���E�U�ɋ����ˑ����Ă����Ԃł��B
�A�^�b�`������ԂŊY���̃u���E�U���I��������ƁA���ׂĂ̎��]�����ł��Ȃ��Ȃ���
���܂��܂��B
�Y���y�[�W�ł̍�Ƃ��I�������A�K���f�^�b�`���s���Ă��������B

  > __detach()

�f�^�b�`���s�����ƂŁA�A�^�b�`�O�̏�ԂɊ���߂��܂��B
�Ȃ��A�u���E�U���I�����Ă��܂��A���ׂĂ̎��]�����ł��Ȃ���ԂƂȂ����ꍇ�A

  [ERROR] 462: �����[�g �T�[�o �}�V�������݂��Ȃ����A���p�ł��܂���B

�̂悤�ȃG���[���b�Z�[�W���\������܂��B
���̃G���[���\�����ꂽ�ꍇ�A�����Ńf�^�b�`���s���܂����A������A���܂��f�^�b
�`����Ȃ��ꍇ�͖����I�Ƀf�^�b�`�R�}���h����͂��Ă��������B
���ׂĂ̎��]�����ł��Ȃ���Ԃł����Ă� __detach() �R�}���h�����͎󂯕t����悤
�ɂȂ��Ă��܂��B

ASConsole ���̂��I������ƁA���ׂẴA�^�b�`��Ԃ͊��S�ɏ���������܂��B
�f�^�b�`�������ɁAASConsole ���I�������Ă����͂���܂���B


8. ���[�U���C�u����

ASConsole �ł̓��[�U���֐��Ȃǂ����R�ɒ�`�ł��܂��B(JavaScript �݂̂��Ώۂł��B)
ASConsole ��W�J�����t�H���_�� lib �t�H���_������܂����Alib �t�H���_���� .js
�t�@�C�����쐬����� ASConsole �N�����Ɏ����I�Ɏ�荞�܂�A�X�N���v�g���ɒ�`
��������̊֐��Ȃǂ��g�p�ł���悤�ɂȂ�܂��B
lib/ipcalc.js �̓��C�u�����̃T���v���ł��B


9. VBScript/Python �ɂ����铮��

VBScript �ł͎��ƕ�����ʂ���܂��B
ASConsole �ł͓��͂����ׂĕ��Ƃ��Ĉ����܂��̂ŁA�P�ɕϐ� a �̓��e������ꍇ�ł��A
print a �̂悤�ɕ��ɂȂ�悤�ɍ\�����Ă��������B

> a = 123         # �ϐ� a �� 123 ����

> println a       # �ϐ� a �̓��e��\�����ĉ��s
123
> a               # �P�� a �ƋL�q����ƁA�G���[�ɂȂ�
[ERROR] 13: �^����v���܂���B
> :a              # a �����Ƃ��ĕ]��
=>123

�� print/println �� ASConsole ���񋟂���R���\�[���o�͗p�w���p�֐��ł��B

Python �� VBScript �Ɠ��l�A���Ǝ�����ʂ���邽�߁A���͕͂��ƂȂ�悤�\������
���������B

�Ȃ��AVBScript/Python �̂�������ϐ��̓��e���ȒP�Ɋm�F����p�r�̂��߂ɁA����\
�����T�|�[�g���܂��B
  ':'(�R����) + ��
�̂悤�ɐ擪�� ':' ��t�����Ď��s����ƁA':' ����菜�����c��̃e�L�X�g��
Execute/exec �ł͂Ȃ��AEval/eval �ŕ]�����A���̌��ʂ�\�����܂��B


10. �X�N���v�g�̌Ăяo�����@�ɂ���

���͎��̕]���� ScriptControl#Run �ŃX�N���v�g���̃w���p�֐�(execEval)���Ăяo
���A�w���p�֐����Ŋe���ꖈ�� eval ���Ăяo���܂��BVBScript/Python �ł̓w���p��
������ Execute/exec ���Ăяo���A���̎��s���s���܂��B
ScriptControl#Eval �ł͂Ȃ��A�����ăw���p�֐����g�p����̂́A�X�N���v�g�̗�O
���Ăяo�����Ő������E���Ȃ��X�N���v�g�G���W�������邱�ƁA�E���Ă��Ăяo������
�̓G���[����K�؂ɒ��o�ł��Ȃ��\�������邱�ƁA���O�̌��ʂ̋L���Ȃǂ̒ǉ���
�����w���p�֐��ŏ�������ȂǁA�����̗��R�ɂ��܂��B
����ł͗�O������Ԃ̌Ăяo���œ`�d�������A�K���A�Ăяo����̃X�N���v�g����
catch ���Ċ�������������Ƃ��Ă��܂��B������������Ԃɂ܂������O�̓`�d���\
���ǂ������������Ă��Ȃ����Ƃ�����܂��B�Ăяo�����ɂ͕����񃌃x���̏�񂾂���
�Ԃ��܂��B

�w���p�֐����o�R���� eval ����̂���{�I�Ȏg�p���@�ƂȂ�܂����A����ł�
ScriptControl#ExecuteStatement ���K�v�ƂȂ�ǖʂ�����܂��B�܂��A
ScriptControl#Eval �̋������m�F�������ꍇ������Ǝv���܂��B(�X�N���v�g�̌���d
�l�ɂ��܂����A���W���[���̃C���|�[�g��֐���`�Ȃ� eval �ł̓G���[�ƂȂ�
�ꍇ�ł� ScriptControl#ExecuteStatement �ŏ����ł���\��������܂��B)
���̂悤�ȏꍇ�A�]�����̐擪�ɓ���̃v���t�B�N�X��t�����Ď��s���邱�ƂŁA����
�I�ɓ���̌Ăяo�����@��I���ł��܂��B

  �v���t�B�N�X ����
 ----------------------------------------------------------------------------
  (�Ȃ�)    ScriptControl#Run (�w���p�֐��o�R�� eval�AVBScript/Python��Execute/exec)
  :         ScriptControl#Run (�w���p�֐��o�R�� Eval/eval�AVBScript/Python�̂�)
  ;         ScriptControl#ExecuteStatement (���ڌďo��)
  ;;        ScriptControl#Eval (���ڌďo��)
  <         ScriptControl#AddCode (���ڌďo��) �X�N���v�g�̃e�L�X�g����͂Ƃ���
  <<        ScriptControl#AddCode (���ڌďo��) �X�N���v�g�̃t�@�C��������͂Ƃ���
  #         �V�X�e���R�}���h
 ----------------------------------------------------------------------------


11. �g�p��̒��ӓ_�Ȃ�

�w���p�֐����� eval ���邽�߁A���̂܂܂ł͓��͎��̎��s�R���e�L�X�g���w���p�֐�
�̊֐����ɂȂ邱�Ƃɒ��ӂ��Ă��������B
���̂��ƂɂƂǂ܂炸�A�X�R�[�v�▼�O��Ԃ̍l�������e���ꖈ�ɈقȂ�܂����A�e��
�ꖈ�̏ڍׂ͍�҂̗����𒴂��Ă��邽�߁A�\�[�X�����Ĕ��f���Ă��������B

���̌���͏��׍H�ɂ���āA���s�R���e�L�X�g���O���[�o�����ɏC�����Ă��܂��B
Ruby �ł́Aeval �� TOPLEVEL_BINDING ���w�肵�Ă��܂��B
Python �ł� exec �� globals() ���w�肵�Ă��܂��B

JavaScript/VBScript �ȊO�̃X�N���v�g����ł́A�W���o�͂ɑ΂��鑀�삪�\�ł��B
���̂����A�W���o�́A�W���G���[�o�͂ɂ��Ă͏o�͐��؂�ւ��邱�Ƃœ��A�v���P�[
�V������ GUI �R���\�[���ւ̕\�����T�|�[�g���܂��B
���A�v���P�[�V�����ł͕W�����͂̓T�|�[�g���܂���B�W�����͂͐���ɓ��삷�邱��
�͂Ȃ��A�X�N���v�g�G���W���ɂ���Ă̓A�v���P�[�V�����G���[�Ȃǂ̒v���I�ȃG���[
�������N�����\�������邽�߁A�Y���@�\���g�p���Ȃ��悤�ɒ��ӂ��Ă��������B
�W���o�͂̑Ή��͎b��ł��B���ɓ��{��̈����ɂ��Ă͂܂��s��̎c���Ă���\
��������܂��B

ActiveScriptRuby 1.8.7�A2.4.x �͓����ɃV�X�e���ɓo�^�ł��܂����A
ASConsole �N�����ɂ͍Ō�ɃC���X�g�[�������o�[�W�����݂̂��L���ɂȂ�܂��B
������ Ruby �𓯎��ɔF�������邽�߂ɂ͈ȉ��̐����ɏ]���A�X�N���v�g���C������
���������B

    (ProgID)
   �X�N���v�g��   ����
 -----------------------------------------------------------------------
  RubyScript     ActiveScriptRuby �̑�\��(���Ƃ���C���X�g�[�������o�[�W�����ŏ㏑��)
  RubyScript.1   ActiveScriptRuby 1.8.7 �̌ʖ�
  RubyScript.2.4 ActiveScriptRuby 2.4.x �̌ʖ�

ActiveScriptRuby ���C���X�g�[������ƁA��L���O�ŃX�N���v�g�G���W�����o�^����
�܂��B���A�v���P�[�V�����̏�����Ԃł͏�L�̑�\���݂̂��L���ɂȂ��Ă���A����
���O�ŃX�N���v�g�G���W�����I������܂��B���̂��߁AASConsole ���N������ƁA1.8.7�A
2.4 �̂����A�ォ��C���X�g�[�������o�[�W�����������I�ɑI������܂��B
�ǂ̃o�[�W�������I������Ă��邩�̊m�F����ɂ͒萔 RUBY_VERSION ���Q�Ƃ��Ă���
�����B�ȉ��ɗ�������܂��B

> RUBY_VERSION
2.4.0

�����̃o�[�W������ ActiveScriptRuby ���C���X�g�[�������V�X�e����ŁA���̗�����
�X�N���v�g�G���W����L���ɂ��邽�߂ɂ́Ascriptmanager.js �����ȉ��̂悤�ɏC����
�܂��B(�R�����g�A�E�g����Ă���ӏ��̃R�����g���͂�������)

...
	[RubyScript,    ["eval.rb"], "rb"],
	[RubyScript18,  ["eval.rb"], "rb18"],
	[RubyScript24,  ["eval.rb"], "rb24"],
...

�T�t�B�b�N�X�Ȃ��� RubyScript �͓����I�� COM ����֐����g�p���܂��B
�e�o�[�W������F��������ꍇ�ł��A�f�t�H���g�� RubyScript �̍s�͗L���Ȃ܂܂ɂ�
�Ă����Ă��������B

ruby �� JavaScript �̃��\�b�h�� nil ��n���Ă� null �Ɖ��߂���Ȃ��悤�ł��B
(JavaScript ���� typeof(value) �� "unknown" �ɂȂ�)
�Ȃ��AJavaScript������ RubyScript�̃��\�b�h���Ă񂾏ꍇ�Anull �͎����I�� nil
�ɕϊ�����܂��B

���Ȃ݂ɁA���̌���� COM �C���^�t�F�[�X�ɂ����� null �\���͈ȉ��ƂȂ�܂��B
(���Ԃ�)

  ----------------------------------------------------------------------
   JavaScript  null
   VBScript    Null
   RubyScript  ?
   PerlScript  Win32::OLE::Variant->new(VT_NULL) �܂��� Variant(VT_NULL)
   Python      None
  ----------------------------------------------------------------------
   ���� null ���Ƃ�����肪����܂���(VT_NULL or VT_EMPTY �Ȃ�)�A�����ł�
   ASConsole �̃z�X�g�ł���A����ԘA�g�̃n�u�ƂȂ� JavaScript �̃��\�b�h����
   �яo�����Ƃ��̌���������Ƃ��Ă��܂��B(�C���^�t�F�[�X�I�u�W�F�N�g��
   JavaScript �Ŏ������Ă���s���̂��߁AJavaScript ���� null �ƌ������
   �̂���ɍl���Ă��܂��B)

�����R�[�h�A���ɓ��{��̈����͑S�̓I�ɍ�҂̗������s�\���ł��B�ȉ��͌���̊o��
�ɂȂ�܂��B
HTA�A���ڎ�荞��.js�t�@�C���A�X�N���v�g�t�@�C���� UTF-8�Ƃ��Ĉ����Ă��܂��B
���Ƃ��A�R�����g�ł����Ă��AShift_JIS�œ��{��Ȃǂ��L�q����ƃX�N���v�g�̉��
�Ɏ��s����\��������܂��B

���s���̕����R�[�h�ɂ��� Ruby 1.8.7 �ł� $KCODE='SJIS' �𓮓I�Ɏw�肵�Ă���
���B1.9�ȍ~�� eval �����p�����^������������I�� .force_encoding('Windows-31J')
���ē��{��Ƃ��ĔF�������Ă��܂��B���̑Ώ��͊��S�ł͂Ȃ��A�Ⴆ��
importVariable/exportVariable �œn�������͐������F���ł��Ȃ��\���������
���B

Python �ł͓��{�ꕶ����ɂ� u ��K���t�����Ă��������B(ex. u"���{��")
u �����Ă�������A�ȒP�ȓ��{�ꏈ���ł͖��͔������Ȃ��悤�ł��B

Perl �ɂ����镶���R�[�h�̈����͐����悭�������Ă��炸�A����ł̓V�X�e�����̍l
���͂���܂���B
�������A�P���ȓ��{�ꕶ�����\�����邾���ł���Γ��ɖ��͂Ȃ��悤�ł��B

�܂�ɕ\�����ł܂��Ă��܂����Ƃ�����܂��B(�A�v�����̂��n���O����킯�ł͂Ȃ�)
�����������ŃR���\�[���̈�ɓ���ȕ�����\���������ꍇ�ɖ�肪��������\����
�����悤�ł��B���̂悤�ȏꍇ�� #clear �R�}���h�ŃR���\�[�������������Ă݂Ă���
�����B����ł��������Ȃ��ꍇ�A�A�v���P�[�V�������I�������Ă��������B���̖���
�������ł��B

comInfo(obj, name) �Ńv���p�e�B�����擾����Ƃ��APROPERTYPUT �̌^���\����
��Ă��܂���B����͌���̎d�l�ł��B(ruby �� win32ole ���g���� PROPERTYPUT
�̌^�����擾������@��������Ȃ�����)

JavaScript �ŃC���X�^���X���\�b�h���쐬���ARuby ����I�u�W�F�N�g�o�R�ł��̃��\�b
�h���Ăяo���Ƃ��A�p�����^������ΐ��������\�b�h�Ăяo�����s���܂����A�p�����^
���Ȃ��Ɗ֐����̂̒�`���(�e�L�X�g?)���擾���Ă��܂��悤�ł�(JavaScript �ł�
foo() �� foo �͕ʕ������Aruby �ł̓��\�b�h�ƃv���p�e�B�̋�ʂ������A�p�����^
�Ȃ��̏ꍇ�A��� foo ���Ȃ킿�v���p�e�B�Ƃ��ĉ��߂��Ă��܂�?)�B
�������g���������݂���̂����m��܂��񂪁A���Ƃ��p�����^�����݂��Ȃ��Ă��_�~�[
�̃p�����^��t�����邱�ƂłƂ肠��������͂ł��܂��B(getText() �ɑ΂���
getText('dummy') �ȂǂƂ���)����� Perl �ł������悤�ł��B
�t�� Python �ł̓p�����^�Ȃ��̃��\�b�h�𐳂��������ł��邤���A�]���ȃp�����^��
����ƃG���[�ƂȂ��ă��\�b�h�Ăяo���Ɏ��s���܂��B

�X�N���v�g�G���W���̒ǉ��̓w���p�֐�(execEval�Ȃǂ������̕K�{�֐�)���܂ރX�N
���v�g�e�L�X�g��p�ӂ��� scriptmanager.js �̃e�[�u�����C�����邱�Ƃōs���܂��B
���̃X�N���v�g�̎������Q�l�ɂ��Ă��������B
�Œ���� execEval ������΂悢�͂��ł��B

ASConsole �͕����N���ł��܂���B����́A�R�}���h�����Ȃǂ̃��[�U�ݒ���𕡐�
�̃C���X�^���X�ŋ��L���邱�Ƃ�������߂ł��B
����ł������āA�����N���������ꍇ�ɂ́Aascon.hta ���ȉ��̂悤�ɏ��������Ă���
�����B�����N�����[�h�ŃA�v���P�[�V�������N�����܂��B

   hta:application�^�O��
      SINGLEINSTANCE="yes"   =>   SINGLEINSTANCE="no"

�����N�����[�h�ł́A�R�}���h�����A�G�C���A�X����ёI�𒆂̃X�N���v�g�G���W����
�ۑ����s���܂���B���̂��߁A�����ĕ����N�����������ꍇ�́A���i�g�p����W���ݒ�
�̃A�v���͂��̂܂܎c���A�t�H���_�S�̂�ʃt�H���_�ɃR�s�[���Ă���A�R�s�[������
�̂�ҏW���邱�Ƃ������߂��܂��B


12. �쐬��

kanegon


13. ���C�Z���X

MIT ���C�Z���X�Ƃ��Č��J���܂��B
�ڍׂ͓����� LICENSE ���Q�Ƃ��Ă��������B


14. �C������

ver0.01  2003-04-26
    �V�K�쐬 (jseval�Ƃ���)

ver1.00  2009-03-24
    ASConsole �ɖ��̕ύX
    �X�N���v�g�̌Ăяo���� ScriptControl �o�R�ɕύX
    JavaScript�AVBScript�ARuby�APerl�APython �ɑΉ�

ver1.07  2011-06-01
    �u���E�U�A�^�b�`�@�\�ǉ�(jQuery �Ή�)

ver2.00  2018-05-28
    �����s�ҏW�@�\�̃C�����C����
    �ݒ�/�����t�@�C��(ascon.dat)��JSON���A�ۑ���ύX
    MIT �����Z���X�ɕύX