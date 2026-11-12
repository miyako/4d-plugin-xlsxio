#include "4DPlugin-XLSXIO.h"
#include "PrivateTypes.h"
#include "EntryPoints.h"
#include "ARRAY_TEXT.h"
#include <cassert>
#include <cstring>
#include <iostream>
#include <map>
#include <set>

static std::set<PA_Unichar*> strings;
static std::map<PA_Handle, size_t> handles;
static unsigned created = 0, disposed = 0;
static int readerToken, listToken;
static bool failOpen = false, failList = false;
static unsigned readerOwners = 0, listOwners = 0;
static size_t sheetIndex = 0;
static std::vector<std::string> sheets;

[[noreturn]] static void unexpected() { std::abort(); }

static PA_Unistring makeString(const PA_Unichar* input) {
    size_t length = 0;
    while (input[length]) ++length;
    PA_Unistring result = {};
    result.fLength = length;
    result.fString = new PA_Unichar[length + 1];
    std::copy_n(input, length + 1, result.fString);
    PA_Unichar* pointer = result.fString;
    assert(strings.insert(pointer).second);
    ++created;
    return result;
}

static void releaseString(PA_Unistring& value) {
    if (value.fString) {
        PA_Unichar* pointer = value.fString;
        assert(strings.erase(pointer) == 1);
        delete[] value.fString;
        ++disposed;
    }
    value = {};
}

static void releaseVariable(PA_Variable& value) {
    if (value.fType == eVK_ArrayUnicode) {
        PA_Handle handle = value.uValue.fArray.fData;
        assert(handles.count(handle));
        auto* elements = reinterpret_cast<PA_Unistring*>(*handle);
        for (int i = 0; i <= value.uValue.fArray.fNbElements; ++i)
            releaseString(elements[i]);
        std::free(*handle);
        delete handle;
        handles.erase(handle);
    } else if (value.fType == eVK_Unistring) {
        releaseString(value.uValue.fString);
    } else {
        assert(value.fType == eVK_Undefined || value.fType == eVK_Longint);
    }
    value = {};
    value.fType = eVK_Undefined;
}

static void host(short selector, EngineBlock* eb) {
    eb->fError = eER_NoErr;
    switch (selector) {
        case EX_CREATE_UNISTRING:
            eb->fUniString1 = makeString(reinterpret_cast<PA_Unichar*>(eb->fHandle));
            break;
        case EX_DISPOSE_UNISTRING:
            releaseString(eb->fUniString1);
            break;
        case EX_SET_UNISTRING: {
            auto next = makeString(reinterpret_cast<PA_Unichar*>(eb->fHandle));
            releaseString(eb->fUniString1);
            eb->fUniString1 = next;
            break;
        }
        case EX_CLEAR_VARIABLE:
            releaseVariable(*static_cast<PA_Variable*>(eb->fPtr1));
            break;
        case EX_HANDLE_MANAGER: {
            auto handle = eb->fHandle;
            switch (eb->fParam1) {
                case 1:
                    handle = new char*;
                    *handle = static_cast<char*>(std::calloc(1, eb->fParam2));
                    assert(*handle);
                    handles[handle] = eb->fParam2;
                    eb->fHandle = handle;
                    break;
                case 2:
                    assert(handles.erase(handle) == 1);
                    std::free(*handle);
                    delete handle;
                    break;
                case 3: {
                    assert(handles.count(handle));
                    auto old = handles.at(handle);
                    auto* replacement = static_cast<char*>(std::calloc(1, eb->fParam2));
                    assert(replacement);
                    std::memcpy(replacement, *handle, std::min(old, size_t(eb->fParam2)));
                    std::free(*handle);
                    *handle = replacement;
                    handles[handle] = eb->fParam2;
                    break;
                }
                case 4:
                    eb->fParam2 = handles.at(handle);
                    break;
                case 5:
                    assert(handles.count(handle));
                    eb->fParam3 = reinterpret_cast<sLONG_PTR>(*handle);
                    break;
                case 6:
                    assert(handles.count(handle));
                    break;
                default: unexpected();
            }
            break;
        }
        default:
            std::cerr << "Unexpected 4D selector: " << selector << '\n';
            unexpected();
    }
}

extern "C" {
xlsxioreader xlsxioread_open(const char* path) {
    assert(std::strcmp(path, "synthetic-only") == 0);
    if (failOpen) return nullptr;
    ++readerOwners;
    return reinterpret_cast<xlsxioreader>(&readerToken);
}
void xlsxioread_close(xlsxioreader reader) {
    assert(reader == reinterpret_cast<xlsxioreader>(&readerToken) && readerOwners == 1);
    --readerOwners;
}
xlsxioreadersheetlist xlsxioread_sheetlist_open(xlsxioreader reader) {
    assert(reader == reinterpret_cast<xlsxioreader>(&readerToken));
    if (failList) return nullptr;
    ++listOwners;
    sheetIndex = 0;
    return reinterpret_cast<xlsxioreadersheetlist>(&listToken);
}
void xlsxioread_sheetlist_close(xlsxioreadersheetlist list) {
    assert(list == reinterpret_cast<xlsxioreadersheetlist>(&listToken) && listOwners == 1);
    --listOwners;
}
const char* xlsxioread_sheetlist_next(xlsxioreadersheetlist list) {
    assert(list == reinterpret_cast<xlsxioreadersheetlist>(&listToken));
    return sheetIndex == sheets.size() ? nullptr : sheets[sheetIndex++].c_str();
}
xlsxioreadersheet xlsxioread_sheet_open(xlsxioreader, const char*, unsigned) { unexpected(); }
void xlsxioread_sheet_close(xlsxioreadersheet) { unexpected(); }
int xlsxioread_sheet_next_row(xlsxioreadersheet) { unexpected(); }
char* xlsxioread_sheet_next_cell(xlsxioreadersheet) { unexpected(); }
xlsxiowriter xlsxiowrite_open(const char*, const char*) { unexpected(); }
int xlsxiowrite_close(xlsxiowriter) { unexpected(); }
void xlsxiowrite_set_row_height(xlsxiowriter, size_t) { unexpected(); }
void xlsxiowrite_set_detection_rows(xlsxiowriter, size_t) { unexpected(); }
void xlsxiowrite_add_cell_string(xlsxiowriter, const char*) { unexpected(); }
void xlsxiowrite_add_cell_int(xlsxiowriter, int64_t) { unexpected(); }
void xlsxiowrite_add_cell_float(xlsxiowriter, double) { unexpected(); }
void xlsxiowrite_next_row(xlsxiowriter) { unexpected(); }
}

static PA_Unistring text(const char16_t* input) {
    static_assert(sizeof(PA_Unichar) == sizeof(char16_t), "UTF-16 units");
    return PA_CreateUnistring(reinterpret_cast<PA_Unichar*>(const_cast<char16_t*>(input)));
}

static void equal(PA_Unistring value, const std::u16string& expected) {
    PA_Unichar* pointer = value.fString;
    assert(pointer && strings.count(pointer));
    assert(size_t(value.fLength) == expected.size());
    for (size_t i = 0; i < expected.size(); ++i) assert(value.fString[i] == expected[i]);
    assert(value.fString[expected.size()] == 0);
}

static void populate(PA_Variable& value) {
    PA_ResizeArray(&value, 2);
    auto zero = text(u"old-zero");
    auto first = text(u"old-\u65e5");
    auto second = text(u"old-\U0001f600");
    PA_SetStringInArray(value, 0, &zero);
    PA_SetStringInArray(value, 1, &first);
    PA_SetStringInArray(value, 2, &second);
}

static void invoke(PA_Variable& output) {
    auto path = text(u"synthetic-only");
    void* args[] = {&path, &output};
    FourDPackex(3, args, nullptr, nullptr);
    PA_DisposeUnistring(&path);
    assert(!readerOwners && !listOwners);
}

static bool destinationCase(const std::string& name) {
    const bool undefined = name == "undefined-destination";
    const bool rejected = name == "rejected-kind";
    const bool initiallyEmpty = name == "empty-caller";
    const bool emptyResult = name == "empty-result";
    failOpen = name == "failed-open";
    failList = name == "failed-list";
    sheets = (emptyResult || failOpen || failList) ? std::vector<std::string>{} :
        std::vector<std::string>{"ASCII", "\xe6\x97\xa5", "\xf0\x9f\x98\x80", ""};
    PA_Variable output = {};
    output.fType = eVK_Undefined;
    if (rejected) {
        output = PA_CreateVariable(eVK_Longint);
        PA_SetLongintVariable(&output, 123);
    } else if (!undefined) {
        output = PA_CreateVariable(eVK_ArrayUnicode);
        if (!initiallyEmpty) populate(output);
    }
    std::set<PA_Unichar*> oldStrings = strings;
    const auto oldHandle = handles.empty() ? nullptr : handles.begin()->first;
    const unsigned oldCount = oldStrings.size();
    const unsigned disposedBefore = disposed;
    const unsigned rounds = name == "repeated-reuse" ? 16 : 1;
    for (unsigned round = 0; round < rounds; ++round) {
        invoke(output);
        if (rejected) {
            assert(output.fType == eVK_Longint && PA_GetLongintVariable(output) == 123);
        } else {
            assert(output.fType == eVK_ArrayUnicode);
            assert(PA_GetArrayNbElements(output) == int(sheets.size()));
            equal(PA_GetStringInArray(output, 0), u"");
            if (!sheets.empty()) {
                equal(PA_GetStringInArray(output, 1), u"ASCII");
                equal(PA_GetStringInArray(output, 2), u"\u65e5");
                equal(PA_GetStringInArray(output, 3), u"\U0001f600");
                equal(PA_GetStringInArray(output, 4), u"");
            }
        }
    }
    PA_ClearVariable(&output);
    const bool balanced = strings.empty() && handles.empty();
    std::cout << name << ": old_owners=" << oldCount << " rounds=" << rounds
              << " disposals=" << disposed - disposedBefore
              << " outstanding_strings=" << strings.size()
              << " outstanding_handles=" << handles.size()
              << " old_handle_live=" << (oldHandle && handles.count(oldHandle))
              << (balanced ? " PASS\n" : " FAIL\n");
    return balanced;
}

static void sdkControls() {
    auto array = PA_CreateVariable(eVK_ArrayUnicode);
    auto zero = text(u"\U0001f600");
    PA_SetStringInArray(array, 0, &zero);
    equal(PA_GetStringInArray(array, 0), u"\U0001f600");
    auto replacement = text(u"\u65e5");
    PA_SetStringInArray(array, 0, &replacement);
    equal(PA_GetStringInArray(array, 0), u"\u65e5");
    auto rejected = text(u"not-transferred");
    PA_SetStringInArray(array, -1, &rejected);
    PA_SetStringInArray(array, 1, &rejected);
    PA_Variable wrong = PA_CreateVariable(eVK_Longint);
    PA_SetStringInArray(wrong, 0, &rejected);
    PA_Variable noData = {};
    noData.fType = eVK_ArrayUnicode;
    PA_SetStringInArray(noData, 0, &rejected);
    PA_SetStringInArray(array, 0, nullptr);
    equal(rejected, u"not-transferred");
    PA_DisposeUnistring(&rejected);
    populate(array);
    PA_ResizeArray(&array, 0);
    equal(PA_GetStringInArray(array, 0), u"old-zero");
    PA_ClearVariable(&array);
}

static void wrapperControls() {
    const char* utf8[] = {"ASCII", "\xe6\x97\xa5", "\xf0\x9f\x98\x80", ""};
    const std::u16string expected[] = {u"ASCII", u"\u65e5", u"\U0001f600", u""};
    ARRAY_TEXT array;
    array.setSize(1);
    for (int i = 0; i < 4; ++i) {
        C_TEXT input;
        input.setUTF8String(reinterpret_cast<const uint8_t*>(utf8[i]), std::strlen(utf8[i]));
        assert(input.getUTF16Length() == expected[i].size());
        CUTF8String roundtrip;
        input.copyUTF8String(&roundtrip);
        assert(std::string(roundtrip.begin(), roundtrip.end()) == utf8[i]);
        array.appendUTF16String(input.getUTF16StringPtr(), input.getUTF16Length());
        auto destination = text(u"previous");
        BytePtr args[] = {reinterpret_cast<BytePtr>(&destination)};
        input.toParamAtIndex(args, 1);
        equal(destination, expected[i]);
        PA_DisposeUnistring(&destination);
    }
    auto destination = PA_CreateVariable(eVK_ArrayUnicode);
    populate(destination);
    BytePtr args[] = {reinterpret_cast<BytePtr>(&destination)};
    array.toParamAtIndex(args, 1);
    for (int i = 0; i < 4; ++i) equal(PA_GetStringInArray(destination, i + 1), expected[i]);
    PA_ClearVariable(&destination);
}

static void scalarDispatcherControls() {
    auto path = text(u"synthetic-only");
    auto errors = text(u"");
    auto result = text(u"previous-result");
    auto* resultPointer = &result;
    PA_long32 options = 0;
    void* readArgs[] = {&path, &options, &errors};
    failOpen = true;
    FourDPackex(1, readArgs, nullptr, &resultPointer);
    equal(result, u"{}");
    equal(errors, u"{\"error\":\"failed to open xlsx\",\"path\":\"synthetic-only\"}");
    PA_DisposeUnistring(&errors);
    errors = text(u"previous-error");
    auto json = text(u"{invalid");
    PA_long32 height = 0, detection = 0;
    void* writeArgs[] = {&path, &json, &height, &detection, &errors};
    FourDPackex(2, writeArgs, nullptr, nullptr);
    equal(errors, u"{\"error\":\"failed to parse json\",\"path\":\"synthetic-only\"}");
    PA_DisposeUnistring(&json);
    PA_DisposeUnistring(&errors);
    PA_DisposeUnistring(&result);
    PA_DisposeUnistring(&path);
    assert(!readerOwners && !listOwners);
}

int main(int argc, char** argv) {
    assert(argc == 2);
    gCall4D = host;
    const std::string name = argv[1];
    bool passed = true;
    if (name == "sdk-transfer-controls") sdkControls();
    else if (name == "wrapper-controls") wrapperControls();
    else if (name == "scalar-dispatcher-controls") scalarDispatcherControls();
    else passed = destinationCase(name);
    if (name == "sdk-transfer-controls" || name == "wrapper-controls" ||
        name == "scalar-dispatcher-controls") {
        passed = strings.empty() && handles.empty();
        std::cout << name << (passed ? " PASS\n" : " FAIL\n");
    }
    // Baseline leak evidence is recorded before bounded fixture cleanup.
    for (auto* string : strings) delete[] string;
    for (auto entry : handles) {
        std::free(*entry.first);
        delete entry.first;
    }
    return passed ? 0 : 1;
}
