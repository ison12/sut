#ifndef SEZARWIN_COMMON_H_
#define SEZARWIN_COMMON_H_

#include <string>
#include <vector>
#include <iostream>
#include <comutil.h>
#include <windows.h>

// �f�o�b�O���[�h�̗L���ŏo�͂�ON�EOFF��؂�ւ���
#ifdef _DEBUG
// �f�o�b�O���[�h�̏ꍇ

    // UNICODE�����Z�b�g�̏ꍇ
#ifdef UNICODE
#define C_OUT(message)    std::wcout << message
#define C_OUT_NL(message) std::wcout << message << std::endl

// �}���`�o�C�g�����Z�b�g�̏ꍇ
#else
#define C_OUT(message)     std::cout << message
#define C_OUT_NL(message)  std::cout << message << std::endl

#endif

#else
// �f�o�b�O���[�h�ł͂Ȃ��ꍇ

    // �����o�͂��Ȃ�
#define C_OUT(message)
#define C_OUT_NL(message)
#endif

/**
 * [�T�v]     : ���ʏ������`�B
 *
 *
 * [���l]     :
 *
 * [�쐬��]   : Sandora
 * [����]     : 2007/08/13   Sandora  �V�K�쐬
 *            :
 *
 * Copyright(c)2007 Sandora All rights reserved.
 *
 */
namespace common {

    /**
        * ������̒u���B
        *
        * @param str �u���Ώە�����
        * @param from ����������
        * @param to   �u��������
        */
    template<class T> void replaceStr(std::basic_string<T>& str
        , const std::basic_string<T>& from
        , const std::basic_string<T>& to)
    {

        std::basic_string<T>::size_type pos = 0;

        while (pos = str.find(from, pos), pos != std::string::npos) {

            str.replace(pos, from.length(), to);

            pos += to.length();
        }

    }

    /**
        * �^����ꂽ�����̍ŏ��l�����߂�B
        *
        * @param n1
        * @param n2
        * @param n3
        * @param n4
        *
        * @return �ŏ��l
        */
    int choiceMinNum(int n1, int n2, int n3, int n4);

    /**
        * �^����ꂽ�����̍ő�l�����߂�B
        *
        * @param n1
        * @param n2
        * @param n3
        * @param n4
        *
        * @return �ő�l
        */
    int choiceMaxNum(int n1, int n2, int n3, int n4);

    /**
        * �^����ꂽ�����̍ő�l�����߂�B
        *
        * @param int�z��
        *
        * @return �ő�l
        */
    int choiceMaxNum(int*);

    /**
        * exe�t�@�C�����u����Ă���p�X���擾����B
        *
        * @return exe�t�@�C���̃p�X
        */
    std::basic_string<TCHAR> getApplicationPath();

    /**
        * ���W���[���t�@�C���iexe��dll�j���u����Ă���p�X���擾����B
        *
        * @param hModule ���W���[���n���h��
        * @return �t�@�C���p�X
        */
    std::basic_string<TCHAR> getModulePath(HMODULE hModule);

    /**
        * �Ō�ɔ��������G���[���b�Z�[�W��\������
        */
    std::basic_string<TCHAR> getErrorMessage(DWORD errorNo);

    /**
        * �Ō�ɔ��������G���[���b�Z�[�W��\������
        */
    std::basic_string<TCHAR> getLastErrorMessage();

    /**
        * �Ō�ɔ��������G���[���b�Z�[�W��\������
        */
    void showErrorMessage(DWORD errorNo);

    /**
        * �Ō�ɔ��������G���[���b�Z�[�W��\������
        */
    void showLastErrorMessage();

    /**
        * �S�Ẵf�B�X�v���C�f�o�C�X�̃O���t�B�b�N�X���[�h�Ɋւ�������擾���܂��B
        *
        */
    std::vector<DEVMODE> getDisplaySettingsInfo();

    /**
        * �S�Ẵf�B�X�v���C�f�o�C�X�̃O���t�B�b�N�X���[�h�Ɋւ�������o�͂��܂��B
        *
        */
    void outDisplaySettingsInfo();

    /**
        * �C�ӂ̐F�ɑ΂��āA�P�x�����������F���v�Z����B
        *
        * @param r ���F�i�Ԑ����j
        * @param g ���F�i�ΐ����j
        * @param b ���F�i�����j
        * @param bgLight �P�x �y0�i�^���Áj <= 1�i�ʏ�j <= ���i���邢�j�z
        * @return �v�Z��̐F
        */
    COLORREF calcBgLight(int r, int g, int b, double bgLight);

    /**
        * �C�ӂ̐F�ɑ΂��āA�P�x�����������F���v�Z����B
        *
        * @param r ���F�i�Ԑ����j
        * @param g ���F�i�ΐ����j
        * @param b ���F�i�����j
        * @param bgLight �P�x �y0�i�^���Áj <= 1�i�ʏ�j <= ���i���邢�j�z
        * @param calcR �v�Z��̐F�i�Ԑ����j
        * @param calcG �v�Z��̐F�i�ΐ����j
        * @param calcB �v�Z��̐F�i�����j
        * @return �v�Z��̐F
        */
    void calcBgLight(int r, int g, int b, double bgLight, int& calcR, int& calcG, int& calcB);

    /**
        * ��̐F����������B
        *
        * @param sr �O�i�i�Ԑ����j
        * @param sg �O�i�i�ΐ����j
        * @param sb �O�i�i�����j
        * @param dr �w�i�i�Ԑ����j
        * @param dg �w�i�i�ΐ����j
        * @param db �w�i�i�����j
        * @param alpha ���ߗ� 0������ 1���s����
        * @return �v�Z��̐F
        */
    COLORREF calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha);

    /**
        * ��̐F����������B
        *
        * @param sr �O�i�i�Ԑ����j
        * @param sg �O�i�i�ΐ����j
        * @param sb �O�i�i�����j
        * @param dr �w�i�i�Ԑ����j
        * @param dg �w�i�i�ΐ����j
        * @param db �w�i�i�����j
        * @param alpha ���ߗ� 0������ 1���s����
        * @param calcR �v�Z��̐F�i�Ԑ����j
        * @param calcG �v�Z��̐F�i�ΐ����j
        * @param calcB �v�Z��̐F�i�����j
        * @return �v�Z��̐F
        */
    void calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha, int& calcR, int& calcG, int& calcB);

    SAFEARRAY* createSafeArrayOneDim(VARENUM type, int size);

    SAFEARRAY* createSafeArrayTwoDim(VARENUM type, int size1, int size2);

    void initSafeArrayOneDim(SAFEARRAY* var);

    void initSafeArrayTwoDim(SAFEARRAY* var);

};

#endif /*COMMON_H_*/
