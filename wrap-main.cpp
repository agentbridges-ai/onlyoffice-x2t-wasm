#include "../../OOXML/Base/Unit.h"

extern "C" {
    int main1(char* xmlPath) {
        // Generated document-local identifiers must be reproducible. The
        // converter runs in an isolated module, so resetting this sequence for
        // each request preserves uniqueness inside a document without leaking
        // wall-clock time into Editor.bin.
        XmlUtils::ResetRand(0x4F4F5832U);
        char *argv[2] = {
            "",
            xmlPath
        };

        return main(2, argv);
    }
}
