#include "pch.h"
#include "Common.h"

int common::choiceMinNum(int n1, int n2, int n3, int n4) {

    int ret = 0;

    if (n1 <= n2 && n1 <= n3 && n1 <= n4) {

        ret = n1;

    }
    else if (n2 <= n1 && n2 <= n3 && n2 <= n4) {

        ret = n2;

    }
    else if (n3 <= n1 && n3 <= n2 && n3 <= n4) {

        ret = n3;

    }
    else {

        ret = n4;
    }

    return ret;
}

int common::choiceMaxNum(int n1, int n2, int n3, int n4) {

    int ret = 0;

    if (n1 >= n2 && n1 >= n3 && n1 >= n4) {

        ret = n1;

    }
    else if (n2 >= n1 && n2 >= n3 && n2 >= n4) {

        ret = n2;

    }
    else if (n3 >= n1 && n3 >= n2 && n3 >= n4) {

        ret = n3;

    }
    else {

        ret = n4;
    }

    return ret;
}

int common::choiceMaxNum(int* var) {

    int ret = 0;

    int length = sizeof(var) / sizeof(var[0]);

    for (int i = 0; i < length; i++) {

        if (ret < var[i]) {

            ret = var[i];
        }
    }

    return ret;
}

std::basic_string<TCHAR> common::getApplicationPath() {

    TCHAR path[_MAX_PATH];

    // ���s�t�@�C���̃t�@�C���p�X�i�t�@�C�����܂ށj���擾����
    GetModuleFileName(NULL, path, sizeof(path));

    // string�^�ɓ���ւ�
    std::basic_string<TCHAR> strPath(path);

    // �t�@�C���p�X�̂ݏ��O���āA�߂�l�Ƃ���
    return strPath.substr(0, strPath.find_last_of(_T("\\")));

}

std::basic_string<TCHAR> common::getModulePath(HMODULE hModule) {

    TCHAR path[_MAX_PATH];

    // ���s�t�@�C���̃t�@�C���p�X�i�t�@�C�����܂ށj���擾����
    GetModuleFileName(hModule, path, sizeof(path));

    // string�^�ɓ���ւ�
    std::basic_string<TCHAR> strPath(path);

    // �t�@�C���p�X�̂ݏ��O���āA�߂�l�Ƃ���
    return strPath.substr(0, strPath.find_last_of(_T("\\")));

}

std::basic_string<TCHAR> common::getErrorMessage(DWORD errorNo) {

    // ���b�Z�[�W�o�b�t�@
    LPTSTR lpMsgBuf;

    // �G���[No���烁�b�Z�[�W���擾����
    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER
        | FORMAT_MESSAGE_FROM_SYSTEM
        | FORMAT_MESSAGE_IGNORE_INSERTS
        , NULL
        , errorNo
        , MAKELANGID(LANG_NEUTRAL
            , SUBLANG_DEFAULT)
        , (LPTSTR)&lpMsgBuf
        , 0
        , NULL);

    // �߂�l
    std::basic_string<TCHAR> message(lpMsgBuf);

    // ���b�Z�[�W�o�b�t�@���������
    LocalFree(lpMsgBuf);

    return message;
}

std::basic_string<TCHAR> common::getLastErrorMessage() {

    // �Ō�ɔ��������G���[�̃G���[�R�[�h���烁�b�Z�[�W���擾����
    return getErrorMessage(GetLastError());
}


void common::showErrorMessage(DWORD errorNo) {

    // ���b�Z�[�W���擾����
    std::basic_string<TCHAR> message = getErrorMessage(errorNo);

    // ���b�Z�[�W�\��
    MessageBox(NULL
        , message.c_str()
        , NULL
        , MB_OK
        | MB_ICONERROR);

}

void common::showLastErrorMessage() {

    // �Ō�ɔ��������G���[�̃��b�Z�[�W��\������
    showErrorMessage(GetLastError());
}

std::vector<DEVMODE> common::getDisplaySettingsInfo() {

    std::vector<DEVMODE> list;

    DEVMODE tmp;

    // EnumDisplaySettings�̖߂�l
    int ret = 1;

    // �񋓂���f�B�X�v���C���̃C���f�b�N�X
    int i = 0;

    while (ret) {

        ret = EnumDisplaySettings(
            NULL,
            i,
            &tmp
        );

        list.push_back(tmp);

        i++;
    }

    return list;
}

void common::outDisplaySettingsInfo() {

    std::vector<DEVMODE> list = common::getDisplaySettingsInfo();

    std::cout << "���f�B�X�v���C�f�o�C�X���̗� " << std::endl;

    // �S�Ẵf�B�X�v���C�����擾����
    for (std::vector<DEVMODE>::iterator i = list.begin(); i != list.end(); i++) {

        // �f�B�X�v���C���[�h�\���̂��擾
        DEVMODE tmp = (*i);

        std::cout << "���f�B�X�v���C " << tmp.dmDeviceName << std::endl;

        std::cout << "�@���@�@�F" << tmp.dmPelsWidth << std::endl;
        std::cout << "�@����  �F" << tmp.dmPelsHeight << std::endl;
        std::cout << "�@�F�[�x�F" << tmp.dmBitsPerPel << std::endl;

    }

    std::cout << "���� " << std::endl;

}

COLORREF common::calcBgLight(int r, int g, int b, double bgLight) {

    // 0�̏ꍇ�A�v�Z���ł��Ȃ����߂P�ɂ���
    if (r == 0) r = 1;
    if (g == 0) g = 1;
    if (b == 0) b = 1;

    int calcR = (int)((double)r * bgLight);
    int calcG = (int)((double)g * bgLight);
    int calcB = (int)((double)b * bgLight);

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

    return RGB(calcR, calcG, calcB);
}

void common::calcBgLight(int r, int g, int b, double bgLight, int& calcR, int& calcG, int& calcB) {

    // 0�̏ꍇ�A�v�Z���ł��Ȃ����߂P�ɂ���
    if (r == 0) r = 1;
    if (g == 0) g = 1;
    if (b == 0) b = 1;

    calcR = (int)((double)r * bgLight);
    calcG = (int)((double)g * bgLight);
    calcB = (int)((double)b * bgLight);

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

}

COLORREF common::calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha) {

    int calcR = (int)(((double)dr * (1.0 - alpha)) + ((double)sr * alpha));
    int calcG = (int)(((double)dg * (1.0 - alpha)) + ((double)sg * alpha));
    int calcB = (int)(((double)db * (1.0 - alpha)) + ((double)sb * alpha));

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

    return RGB(calcR, calcG, calcB);
}

void common::calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha, int& calcR, int& calcG, int& calcB) {

    calcR = (int)(((double)dr * (1.0 - alpha)) + ((double)sr * alpha));
    calcG = (int)(((double)dg * (1.0 - alpha)) + ((double)sg * alpha));
    calcB = (int)(((double)db * (1.0 - alpha)) + ((double)sb * alpha));

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

}

SAFEARRAY* common::createSafeArrayOneDim(VARENUM type, int size)
{
    // �z�����ݒ肷��(�ꎟ��)
    // �z����\����
    SAFEARRAYBOUND rgb[1];
    // �T�C�Y��ݒ肷��
    rgb[0].cElements = size;
    // �����l��ݒ肷��
    rgb[0].lLbound = 0;

    // �z��𐶐�����
    SAFEARRAY* psa = SafeArrayCreate(type, 1, rgb);

    // �z��̐����ɐ����������𔻒肷��
    if (!psa) {

        return NULL;
    }

    return psa;
}

SAFEARRAY* common::createSafeArrayTwoDim(VARENUM type, int size1, int size2)
{
    // �z�����ݒ肷��(�񎟌�)
    // �z����\����
    SAFEARRAYBOUND rgb[2];
    // �T�C�Y��ݒ肷��
    rgb[0].cElements = size1;
    // �����l��ݒ肷��
    rgb[0].lLbound = 0;
    // �T�C�Y��ݒ肷��
    rgb[1].cElements = size2;
    // �����l��ݒ肷��
    rgb[1].lLbound = 0;

    // �z��𐶐�����
    SAFEARRAY* psa = SafeArrayCreate(type, 2, rgb);

    // �z��̐����ɐ����������𔻒肷��
    if (!psa) {

        return NULL;
    }

    return psa;
}

void common::initSafeArrayOneDim(SAFEARRAY* var)
{
    // ������ł��邱�Ƃ�O��Ƃ��ăe�X�g�f�[�^��ݒ肷��
    BSTR* sData;

    // �f�[�^�ɃA�N�Z�X����
    HRESULT hr = SafeArrayAccessData(var, (void**)&sData);

    if (S_OK != hr) {

        return;
    }

    // �z����\����
    SAFEARRAYBOUND bound = var->rgsabound[0];

    for (ULONG i = 0; i < bound.cElements; i++) {

        // ���C�h������X�g���[��
        std::wstringstream wsstream;
        // �������ݒ肷��
        wsstream << L"�f�[�^" << i << std::endl;
        // �z��ɕ������ݒ肷��
        sData[i] = SysAllocString(wsstream.str().c_str());
    }

    // �f�[�^�A�N�Z�X���������
    hr = SafeArrayUnaccessData(var);

    if (S_OK != hr) {

        return;
    }


}

void common::initSafeArrayTwoDim(SAFEARRAY* var)
{
    // �z������b�N����
    HRESULT hr = SafeArrayLock(var);

    if (S_OK != hr) {

        return;
    }

    // �z����\����
    SAFEARRAYBOUND bound1 = var->rgsabound[0];
    SAFEARRAYBOUND bound2 = var->rgsabound[1];

    // ������ł��邱�Ƃ�O��Ƃ��ăe�X�g�f�[�^��ݒ肷��
    BSTR* sData;
    // �v�f�ʒu
    long indices[2];

    for (ULONG i = 0; i < bound1.cElements; i++) {

        for (ULONG j = 0; j < bound2.cElements; j++) {

            // ���C�h������X�g���[��
            std::wstringstream wsstream;

            // �������ݒ肷��
            wsstream << L"�f�[�^" << i << L"-" << j << std::endl;

            //�v�f�ʒu[i][j]�ւ̒l�Z�b�g
            indices[0] = j;
            indices[1] = i;

            // �z�񂩂�f�[�^�|�C���^���擾����
            SafeArrayPtrOfIndex(var, indices, (void HUGEP * FAR*) & sData);

            // �z��ɕ������ݒ肷��
            *sData = SysAllocString(wsstream.str().c_str());
        }

    }

    // �z��̃��b�N���������
    hr = SafeArrayUnlock(var);

    if (S_OK != hr) {

        return;
    }


}
