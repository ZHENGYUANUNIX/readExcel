#include <iostream>
#include "msgConvert.h"

enum FileName {
    NcStudioCN = 1,
    NcStudioRus,
    ExcelFile,
};

int main(int argc, char **argv)
{
    if (argc < 4) {
        std::cout << "please check whether your input is correct." << std::endl;
        return -1;
    }

    MsgConvert objMsg(argv[NcStudioCN], argv[NcStudioRus], argv[ExcelFile]);
    if (objMsg.convert()) {
        std::cout << "convert success." << std::endl;
    }
    else {
        std::cout << "convert failed." << std::endl;
    }
    return 0;
}
