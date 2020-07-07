#include "msgConvert.h"
#include <windows.h>
#include <vector>
#include <fstream>
#include <xlnt/xlnt.hpp>

MsgConvert::MsgConvert(const char *ptrStudioCN, const char *ptrStudioRus, const char *ptrExcel) :
    m_strNcStudioCN(ptrStudioCN), m_strNcStudioRus(ptrStudioRus), m_strExcel(ptrExcel)
{
}

MsgConvert::~MsgConvert() = default;


std::string MsgConvert::utf8ToString(const std::string& strSource)
{
    std::vector<wchar_t> vectorWChar;
    std::vector<char>    vectorChar;
    int nLengthWChar = MultiByteToWideChar(CP_UTF8, 0, strSource.c_str(), -1, nullptr, 0);
    vectorWChar.resize(nLengthWChar + 1);
    MultiByteToWideChar(CP_UTF8, 0, strSource.c_str(), strSource.length(), vectorWChar.data(), nLengthWChar);
    int nLengthChar = WideCharToMultiByte(CP_ACP, 0, vectorWChar.data(), -1, nullptr, 0, nullptr, nullptr);
    vectorChar.resize(nLengthChar + 1);
    WideCharToMultiByte(CP_ACP, 0, vectorWChar.data(), nLengthWChar, vectorChar.data(), nLengthChar, nullptr, nullptr);
    return vectorChar.data();
}

bool MsgConvert::convert()
{
    xlnt::workbook  objWork;
    objWork.load(m_strExcel.c_str());

    xlnt::worksheet objWorksheet = objWork.active_sheet();
    auto sheetCols = objWorksheet.columns();
    auto sheetRows = objWorksheet.rows();
    int ColLength = sheetCols.length();
    int RowLength = sheetRows.length();

    unsigned int nIndex = 0;
    unsigned int nRowCN = 0;
    unsigned int nRowRus = 0;
    std::string strRead;
    for (; nIndex < ColLength; nIndex++) {
        // 第0行第nIndex列（反的）
        strRead = sheetCols[nIndex][0].value<std::string>();
        // strRead = utf8ToString(strRead);
        if (strRead == "中文") {  // 在第一行找到中文列的列号
            nRowCN = nIndex;
        }
        else if (strRead == std::string("Русский язык")) {  // 在第一行找到俄语列的列号
            nRowRus = nIndex;
        }
        if (nRowCN != 0 && nRowRus != 0) {
            break;
        }
    }
    if (nRowCN == 0 || nRowRus == 0) {
        std::cout << "do not read Chinese or Russian data column, please check." << std::endl;
        return false;
    }

    std::cout << "start reading excel." << std::endl;

    std::string strMsgCN;
    std::string strMsgRus;
    for (nIndex = 1; nIndex < RowLength; nIndex++) {
        strRead = sheetCols[nRowCN][nIndex].value<std::string>();
        strMsgCN = utf8ToString(strRead);
        strRead = sheetCols[nRowRus][nIndex].value<std::string>();
        strMsgRus = utf8ToString(strRead);
        m_mapMsg.emplace(std::make_pair(strMsgCN, strMsgRus));
    }
//    打印数据
//    for (auto const msg : m_mapMsg) {
//        std::cout << msg.first << "\t" << msg.second << std::endl;
//    }
    std::cout << "finish reading excel file, start to translate, please wait..." << std::endl;
    return msgReplace();
}

bool MsgConvert::msgReplace()
{
    std::fstream fStreamRus;
    fStreamRus.open(m_strNcStudioCN.c_str(), std::fstream::in);
    if (!fStreamRus.is_open()) {
        std::cout << "can not open the file of " << m_strNcStudioCN << "." << std::endl;
        return false;
    }

    std::vector<char> vecRead;
    std::size_t nFound = 0;
    std::string strLine;
    while (!fStreamRus.eof()) {
        vecRead.clear();
        vecRead.resize(256);
        nFound = 0;
        fStreamRus.getline(vecRead.data(), static_cast<unsigned int>(vecRead.size()));
        strLine = vecRead.data();
        if (!strLine.empty()) {
        //  strLine = utf8ToString(strLine);
            for (auto const msg : m_mapMsg) {
                // 一行数据中可能有多个要翻译的文本
                while (!msg.second.empty()) {  // 翻译不能为空
                    nFound = strLine.find(msg.first);
                    if (nFound == std::string::npos) {
                        break;
                    }
                    else {
                        strLine = strLine.replace(nFound, msg.first.length(), msg.second);
                    }
                }
            }
        }
        strLine += "\n";
        m_strMsg += strLine;
    }
    fStreamRus.close();
    return writeMsgFile();
}

bool MsgConvert::writeMsgFile()
{
    std::fstream fStreamRus;
    fStreamRus.open(m_strNcStudioRus.c_str(), std::fstream::out | std::fstream::trunc);
    if (!fStreamRus.is_open()) {
        std::cout << "can not open the file of " << m_strNcStudioRus << "." << std::endl;
        return false;
    }
    fStreamRus << m_strMsg;
    fStreamRus.close();
    return true;
}
