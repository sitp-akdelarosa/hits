<%
'========================================================================
' ExcelCreator Ver3.0 �萔�ꗗ
'========================================================================

' Excel�o�[�W����
Public Const ver95 = 95                        ' Excel95
Public Const ver97 = 97                        ' Excel97
Public Const ver2000 = 2000                    ' Excel2000
Public Const ver2002 = 2002                    ' Excel2002

' �J���[�p���b�g
Public Const xcDefault = 0                     ' �J���[�p���b�g ����
Public Const xcBlack = 8                       '                ��
Public Const xcWhite = 9                       '                ��
Public Const xcRed = 10                        '                ��
Public Const xcGreen = 11                      '                ��
Public Const xcBlue = 12                       '                ��
Public Const xcYellow = 13                     '                ��
Public Const xcPink = 14                       '                �s���N
Public Const xcCyan = 15                       '                ��
' �t�H���g�X�^�C��
Public Const fsDefault = &H0                   ' �t�H���g�X�^�C�� �m�[�}��
Public Const fsNormal = &H0                    '                  �m�[�}��
Public Const fsBold = &H1                      '                  �{�[���h
Public Const fsItalic = &H2                    '                  �C�^���b�N
Public Const fsStrike = &H4                    '                  �������
Public Const fsUp = &H8                        '                  ��t��
Public Const fsDown = &H10                     '                  ���t��
' �t�H���g����
Public Const fuDefault = 0                     ' �t�H���g���� �Ȃ�
Public Const fuNone = 0                        '              �Ȃ�
Public Const fuNormal = 1                      '              ����
Public Const fuDouble = 2                      '              ��d��
' �r���̐���
Public Const lsNone = &H0                      ' �r�� �Ȃ�
Public Const lsDefault = &H0                   '      �Ȃ�
Public Const lsNormal = &H1000000              '      ����
Public Const lsThick = &H2000000               '      ����
Public Const lsBroken = &H3000000              '      �j��
Public Const lsDot = &H4000000                 '      �_��
Public Const lsThick2 = &H5000000              '      �ɑ���
Public Const lsDouble = &H6000000              '      ��d��
Public Const lsSlender = &H7000000             '      �׎���
Public Const lsMidBroken = &H8000000           '      ���j��
Public Const lsSlnChain1 = &H9000000           '      �׈�_����
Public Const lsMidChain1 = &HA000000           '      ����_����
Public Const lsSlnChain2 = &HB000000           '      �ד�_����
Public Const lsMidChain2 = &HC000000           '      ����_����
Public Const lsMidChains = &HD000000           '      ����_�΍���
' �p�^�[���ԍ�
Public Const pn01 = &H1000000                  ' �p�^�[���ԍ� �P
Public Const pn02 = &H2000000                  ' �p�^�[���ԍ� �Q
Public Const pn03 = &H3000000                  ' �p�^�[���ԍ� �R
Public Const pn04 = &H4000000                  ' �p�^�[���ԍ� �S
Public Const pn05 = &H5000000                  ' �p�^�[���ԍ� �T
Public Const pn06 = &H6000000                  ' �p�^�[���ԍ� �U
Public Const pn07 = &H7000000                  ' �p�^�[���ԍ� �V
Public Const pn08 = &H8000000                  ' �p�^�[���ԍ� �W
Public Const pn09 = &H9000000                  ' �p�^�[���ԍ� �X
Public Const pn10 = &HA000000                  ' �p�^�[���ԍ� 10
Public Const pn11 = &HB000000                  ' �p�^�[���ԍ� 11
Public Const pn12 = &HC000000                  ' �p�^�[���ԍ� 12
Public Const pn13 = &HD000000                  ' �p�^�[���ԍ� 13
Public Const pn14 = &HE000000                  ' �p�^�[���ԍ� 14
Public Const pn15 = &HF000000                  ' �p�^�[���ԍ� 15
Public Const pn16 = &H10000000                 ' �p�^�[���ԍ� 16
Public Const pn17 = &H11000000                 ' �p�^�[���ԍ� 17
Public Const pn18 = &H12000000                 ' �p�^�[���ԍ� 18
' �{�b�N�X�̌`��
Public Const btBox = 0                         ' �{�b�N�X�̌`�� ����
Public Const btLtc = 1                         '                �i�q��
Public Const btOver = 2                        '                �㉡��
Public Const btUnder = 3                       '                ������
Public Const btLeft = 4                        '                ���c��
Public Const btRight = 5                       '                �E�c��
' �����ʒu
Public Const phNormal = 20                     ' �����ʒu �W��
Public Const phLeft = 21                       '          ���l
Public Const phCenter = 22                     '          ��������
Public Const phRight = 23                      '          �E�l
Public Const phLoop = 24                       '          �J��Ԃ�
Public Const phBothEven = 25                   '          ���[����
Public Const phSelCenter = 26                  '          �I��͈͓��Œ���
Public Const phEven = 27                       '          �ϓ�
' �����ʒu
Public Const pvUp = 0                          ' �����ʒu ��l��
Public Const pvCenter = 1                      '          ��������
Public Const pvDown = 2                        '          ���l��
Public Const pvBothEven = 3                    '          ���[����
Public Const pvEven = 4                        '          �ϓ�
' ��������
Public Const ptHorz = 100                      ' �������� ��
Public Const ptVert = 101                      '          �c
Public Const ptRvTurn90 = 102                  '          �����v���ɂX�O��
Public Const ptTurn90 = 103                    '          ���v���ɂX�O��
' �p���T�C�Y
Public Const psLetter = 1                      ' �p���T�C�Y ���^�[
Public Const psRejoer = 3                      '            ���W���[
Public Const psRegal = 5                       '            ���[�K��
Public Const psExecutive = 7                   '            �G�O�[�N�e�B�u
Public Const psA3 = 8                          '            �`�R
Public Const psA4 = 9                          '            �`�S
Public Const psA5 = 11                         '            �`�T
Public Const psB4 = 12                         '            �a�S
Public Const psB5 = 13                         '            �a�T
' �p������
Public Const orLandscape = 0                   ' �p������ ��
Public Const orPortrait = 2                    '          �c
' �I�[�v�����[�h
Public Const opOpen = 0                        ' �I�[�v�����[�h �I�[�v��
Public Const opCreate = 1                      '                �N���G�C�g
' �y�[�W����
Public Const pcNone = 0                        ' �w��Ȃ�
Public Const pcHorz = 1                        ' ��������
Public Const pcVert = 2                        ' ��������
' �y�[�W����
Public Const poRightFromLeft = 0               ' ������E��
Public Const poBottomFromTop = 1               ' �ォ�牺��
' �ϐ����̍������݃��[�h
Public Const vmValidAll = 0                    ' �S�Ă̕ϐ�����Ώ�
Public Const vmInvalidAnswer = 1               ' �֐����ʂ̕ϐ����͑ΏۊO
Public Const vmInvalidTextBox = 2              ' �e�L�X�g�{�b�N�X�̕ϐ����͑ΏۊO
Public Const vmInvalidCellText = 4             ' �Z���ɒ��ڐݒ肵���ϐ����͑ΏۊO

Public Const vtStr = 0                         ' ������
Public Const vtDouble = 1                      ' �����l
Public Const vtInt = 2                         ' ����
Public Const vtFunc = 3                        ' �v�Z��
Public Const vtNull = 4                        ' �f�[�^�Ȃ�

Public Const errDontOpen = 1                   ' �t�@�C�����I�[�v���ł��Ȃ�
Public Const errFileHandle = 2                 ' �����ȃt�@�C���n���h��
Public Const errNotOpen = 3                    ' ���I�[�v���G���[
Public Const errSheetNo = 4                    ' �V�[�g�ԍ��G���[
Public Const errWrite = 5                      ' �������ݎ��̃G���[
Public Const errAction = 6                     ' ���샂�[�h�G���[
Public Const errDataCat = 7                    ' �f�[�^��ʃG���[
Public Const errClear = 20                     ' �N���A�G���[
Public Const errCopy = 21                      ' �R�s�[�G���[
Public Const errPaste = 22                     ' �\��t���G���[
Public Const errInsert = 23                    ' �}���G���[
Public Const errDelete = 24                    ' �폜�G���[
Public Const errLength = 31                    ' �����G���[
Public Const errLocate = 32                    ' ���W�G���[
Public Const errAttr = 33                      ' �����ԍ��G���[
Public Const errParame = 34                    ' �p�����[�^�G���[
Public Const errNoData = 35                    ' �f�[�^������
Public Const errEndOfData = 36                 ' �f�[�^�̏I���
Public Const errVerPoint = 37                  ' �w�肵���ϐ������݂��Ȃ�
Public Const errBreakCount = 38                ' ���y�[�W���������𒴂�����
Public Const errMemory1 = 40                   ' �������[�s���G���[�P
Public Const errMemory2 = 41                   ' �������[�s���G���[�Q
Public Const errOther = 50                     ' �����G���[
Public Const errFunction = 100                 ' �v�Z���`���G���[
%>