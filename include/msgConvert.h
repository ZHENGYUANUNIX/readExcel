#pragma once
#include <string>
#include <map>

struct orderByStringLength {
    bool operator()(const std::string& keyFist, const std::string& keySecond) const {
        if (keyFist.length() > keySecond.length()) {
            return true;
        }
        else if (keyFist.length() < keySecond.length()) {
            return false;
        }
        else {
            return (keyFist > keySecond);
        }
    }
};

class MsgConvert {
public:
    MsgConvert(const char*, const char*, const char*);
    ~MsgConvert();

public:
    bool convert();

private:
    std::string utf8ToString(const std::string&);
    bool msgReplace();
    bool writeMsgFile();

private:
    std::string     m_strNcStudioCN;
    std::string     m_strMsg;
    std::string     m_strNcStudioRus;
    std::string     m_strExcel;
    std::map<std::string, std::string, orderByStringLength>  m_mapMsg;
};

