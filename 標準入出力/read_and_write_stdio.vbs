'http://www.atmarkit.co.jp/ait/articles/0707/26/news128.html

'cscript.exe //NoLogo �t�@�C�����ŋN�����邱��

'http://homepage2.nifty.com/nihon-nouen/programming-stdinout.htm
'StdIn	Read	�w�肳�ꂽ����������̓X�g���[������ǂݍ��݁A���ʂ̕������Ԃ�
'       ReadAll	���̓X�g���[���S�̂�ǂݍ��݁A���ʂ̕������Ԃ�
'       ReadLine	�s�S�� (���s�����̒��O�܂�) ����̓X�g���[������ǂݍ��݁A���ʂ̕������Ԃ�
'       Skip	���̓X�g���[���̓ǂݍ��ݒ��ɁA�w�肳�ꂽ���������X�L�b�v����
'       SkipLine	���̓X�g���[���̓ǂݍ��ݒ��ɁA���� 1 �s���X�L�b�v����
'StdOut	Write	�w�肳�ꂽ��������o�̓X�g���[���ɏ�������
'       WriteBlankLines	�w�肳�ꂽ���̉��s�������o�̓X�g���[���ɏ�������
'       WriteLine	�w�肳�ꂽ������Ɖ��s�������o�̓X�g���[���ɏ�������


Option Explicit
Dim objStdIn, objStdOut
Set objStdIn  = WScript.StdIn  '�W�����̓X�g���[����Ԃ�
Set objStdOut = WScript.StdOut '�W���o�̓X�g���[����Ԃ�

Dim strFromStdIn
'�W�����͂��當�����1�s�ǂݍ���
strFromStdIn = objStdIn.ReadLine()

'�W���o�͂ɕ������1�s��������
objStdOut.WriteLine strFromStdIn

objStdIn.Close()  '�W�����̓X�g���[�������
objStdOut.Close() '�W���o�̓X�g���[�������

'�I�u�W�F�N�g�̔j��
Set objStdIn  = Nothing
Set objStdOut = Nothing

