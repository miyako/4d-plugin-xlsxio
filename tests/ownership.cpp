#include "4DPlugin-XLSXIO.h"
#include "PrivateTypes.h"
#include "EntryPoints.h"
#include "ARRAY_TEXT.h"
#include "C_BLOB.h"
#include "support/4DPlugin-JSON.h"
#include <cassert>
#include <cstring>
#include <iostream>
#include <map>
#include <set>
#include <stdexcept>
#include <type_traits>

static std::set<PA_Unichar*> strings;
static std::map<PA_Handle, size_t> handles;
static unsigned created = 0, disposed = 0;
static int readerToken, listToken;
static bool failOpen = false, failList = false;
static unsigned readerOwners = 0, listOwners = 0;
static size_t sheetIndex = 0;
static std::vector<std::string> sheets;
struct SyntheticRef {
    std::map<std::u16string, PA_Variable> properties;
};
static std::map<SyntheticRef*, unsigned> references;
static bool failProperty = false, failHas = false, failStringify = false, failCopy = false;
static bool failSet = false, nullEmptyPicture = false;
static bool failStringGet = false;
static PA_Variable commandResult = {};

static void releaseVariable(PA_Variable& value);

static SyntheticRef* newReference() {
    auto* reference = new SyntheticRef;
    references[reference] = 1;
    return reference;
}

static void releaseReference(void* pointer) {
    if (!pointer) return;
    auto* reference = static_cast<SyntheticRef*>(pointer);
    assert(references.count(reference));
    if (--references.at(reference) == 0) {
        references.erase(reference);
        for (auto& entry : reference->properties) releaseVariable(entry.second);
        delete reference;
    }
}

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
    } else if (value.fType == eVK_Unistring || value.fType == 0) {
        releaseString(value.uValue.fString);
    } else if (value.fType == eVK_Object) {
        releaseReference(value.uValue.fObject);
    } else if (value.fType == eVK_Collection) {
        releaseReference(value.uValue.fCollection);
    } else if (value.fType == eVK_Picture) {
        releaseReference(value.uValue.fPicture);
    } else {
        assert(value.fType == eVK_Undefined || value.fType == eVK_Longint ||
               value.fType == eVK_Boolean || value.fType == eVK_Real);
    }
    value = {};
    value.fType = eVK_Undefined;
}

static PA_Variable cloneVariable(PA_Variable value) {
    if (value.fType == eVK_Unistring || value.fType == 0)
        value.uValue.fString = makeString(value.uValue.fString.fString);
    else if (value.fType == eVK_Object || value.fType == eVK_Collection ||
             value.fType == eVK_Picture) {
        auto* reference = static_cast<SyntheticRef*>(value.uValue.fObject);
        assert(references.count(reference));
        ++references.at(reference);
    }
    return value;
}

static std::u16string keyText(PA_Unistring key) {
    PA_Unichar* pointer = key.fString;
    assert(strings.count(pointer));
    return std::u16string(pointer, pointer + key.fLength);
}

static void poison(PA_Variable* output, EngineBlock* eb) {
    // Failed output has no ownership or interpretation contract.
    std::memset(output, 0xa5, sizeof(*output));
    output->fType = 0; // The SDK normalizes raw zero to text.
    eb->fError = eER_CallNotImplemented;
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
        case EX_VARIABLE_TO_STRING: {
            auto* value = reinterpret_cast<PA_Variable*>(eb->fHandle);
            assert(value->fType == eVK_Unistring || value->fType == 0);
            PA_Unichar* pointer = value->uValue.fString.fString;
            assert(strings.count(pointer));
            if (failStringGet) {
                static PA_Unichar uninterpreted[] = {'X', 0};
                eb->fUniString1 = {1, uninterpreted, 0, 0};
                eb->fError = eER_CallNotImplemented;
            } else eb->fUniString1 = value->uValue.fString;
            break;
        }
        case EX_CREATE_PICTURE:
            assert(!eb->fPtr1 && eb->fParam1 == 0);
            eb->fPicture = nullEmptyPicture ? nullptr : newReference();
            break;
        case EX_DISPOSE_PICTURE:
            releaseReference(eb->fPicture);
            break;
        case EX_GET_OBJ_VALUE: {
            assert(eb->fShort == eVK_Object);
            auto* object = static_cast<SyntheticRef*>(eb->fPtr1);
            assert(references.count(object));
            auto* output = static_cast<PA_Variable*>(eb->fPtr2);
            if (failProperty) poison(output, eb);
            else *output = cloneVariable(object->properties.at(keyText(eb->fUniString1)));
            break;
        }
        case EX_SET_OBJ_VALUE: {
            assert(eb->fShort == eVK_Object);
            auto* object = static_cast<SyntheticRef*>(eb->fPtr1);
            assert(references.count(object));
            if (failSet) {
                eb->fError = eER_CallNotImplemented;
                break;
            }
            auto key = keyText(eb->fUniString1);
            auto found = object->properties.find(key);
            if (found != object->properties.end()) releaseVariable(found->second);
            object->properties[key] = cloneVariable(*static_cast<PA_Variable*>(eb->fPtr2));
            break;
        }
        case EX_EXECUTE_COMMAND_BY_ID: {
            auto** args = static_cast<PA_Variable**>(eb->fPtr1);
            auto* output = static_cast<PA_Variable*>(eb->fPtr2);
            if (eb->fParam1 == 1231) { // OB Is defined
                assert(eb->fParam2 == 2);
                if (failHas) poison(output, eb);
                else {
                    auto* object = static_cast<SyntheticRef*>(args[0]->uValue.fObject);
                    assert(references.count(object));
                    *output = {};
                    PA_SetBooleanVariable(output,
                        object->properties.count(keyText(args[1]->uValue.fString)) != 0);
                }
            } else if (eb->fParam1 == 1225) { // OB Copy
                assert(eb->fParam2 == 1);
                if (failCopy) poison(output, eb);
                else {
                    auto* source = static_cast<SyntheticRef*>(args[0]->uValue.fObject);
                    assert(references.count(source));
                    auto* copy = newReference();
                    for (auto entry : source->properties)
                        copy->properties[entry.first] = cloneVariable(entry.second);
                    PA_SetObjectVariable(output, copy);
                }
            } else if (eb->fParam1 == 1217) { // JSON Stringify
                assert(eb->fParam2 == 1);
                auto* object = static_cast<SyntheticRef*>(args[0]->uValue.fObject);
                assert(references.count(object));
                if (failStringify) poison(output, eb);
                else *output = cloneVariable(commandResult);
            } else unexpected();
            break;
        }
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
    const uint8_t payload[] = {0, 1, 0xff};
    C_BLOB blob;
    blob.setBytes(payload, sizeof(payload));
    assert(blob.getBytesLength() == sizeof(payload));
    assert(std::memcmp(blob.getBytesPtr(), payload, sizeof(payload)) == 0);
    CBytes source;
    source.setBytes(payload, sizeof(payload));
    CBytes copy = source;
    source.addBytes(payload, sizeof(payload));
    assert(copy.getBytesLength() == sizeof(payload));
    assert(copy.getBytesPtr() != source.getBytesPtr());
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

static void jsonSetter(const std::string& name) {
    auto* object = newReference();
    const char* utf8[] = {"ASCII", "\xe6\x97\xa5", "\xf0\x9f\x98\x80", ""};
    const wchar_t* wide[] = {L"ASCII", L"\u65e5", L"\U0001f600", L""};
    const std::u16string expected[] = {u"ASCII", u"\u65e5", u"\U0001f600", u""};
    for (int round = 0; round < 3; ++round) {
        for (int i = 0; i < 4; ++i) {
            if (name == "json-set-narrow") ob_set_s(object, "key", utf8[i]);
            else if (name == "json-set-wide") ob_set_s(object, L"key", utf8[i]);
            else if (name == "json-set-wide-value") ob_set_a(object, L"key", wide[i]);
            else {
                CUTF16String value(expected[i].begin(), expected[i].end());
                ob_set_a(object, L"key", &value);
            }
            equal(object->properties.at(u"key").uValue.fString, expected[i]);
        }
    }
    PA_DisposeObject(object);
}

static void pictureSetter(bool emptyNull) {
    nullEmptyPicture = emptyNull;
    auto* object = newReference();
    for (int round = 0; round < 3; ++round) {
        auto* picture = newReference();
        ob_set_p(object, L"key", picture);
        assert(references.at(picture) == 1); // The property owns the consumed input.
        assert(object->properties.at(u"key").uValue.fPicture == picture);
    }
    auto* untouched = newReference();
    ob_set_p(nullptr, L"key", untouched);
    assert(references.at(untouched) == 1);
    releaseReference(untouched);
    PA_DisposeObject(object);
}

static PA_Variable fixtureValue(int kind) {
    PA_Variable value = {};
    value.fType = eVK_Undefined;
    if (kind == 0 || kind == 1) {
        auto string = text(u"ASCII-\u65e5-\U0001f600");
        PA_SetStringVariable(&value, &string);
        if (kind == 1) value.fType = 0;
    } else if (kind == 2) PA_SetBooleanVariable(&value, true);
    else if (kind == 3) PA_SetRealVariable(&value, 42.5);
    else if (kind == 4) PA_SetObjectVariable(&value, newReference());
    else if (kind == 5) PA_SetCollectionVariable(&value, newReference());
    else if (kind == 6) PA_SetPictureVariable(&value, newReference());
    else unexpected();
    return value;
}

static void callGetter(int getter, SyntheticRef* object, int expectedKind) {
    if (getter == 0) {
        CUTF8String output(reinterpret_cast<const uint8_t*>("unchanged"));
        assert(ob_get_a(object, L"key", &output));
        const char* expected = expectedKind <= 1 ?
            "ASCII-\xe6\x97\xa5-\xf0\x9f\x98\x80" : "unchanged";
        assert(std::string(output.begin(), output.end()) == expected);
    } else if (getter == 1) {
        assert(ob_get_b(object, L"key") == (expectedKind == 2));
    } else if (getter == 2) {
        assert(ob_get_n(object, L"key") == (expectedKind == 3 ? 42.5 : 0));
    } else {
        auto collection = ob_get_c(object, L"key");
        assert(bool(collection) == (expectedKind == 5));
        if (collection) PA_DisposeCollection(collection);
    }
}

static void jsonGetter(int getter) {
    for (int kind = 0; kind < 7; ++kind) {
        auto* object = newReference();
        object->properties[u"key"] = fixtureValue(kind);
        for (int round = 0; round < 3; ++round) callGetter(getter, object, kind);
        PA_DisposeObject(object);
    }
    auto* absent = newReference();
    CUTF8String output;
    assert(!ob_get_a(absent, L"missing", &output));
    assert(!ob_get_b(absent, L"missing"));
    assert(ob_get_n(absent, L"missing") == 0);
    assert(!ob_get_c(absent, L"missing"));
    PA_DisposeObject(absent);
}

static void getterFailure(bool hasFailure) {
    auto* object = newReference();
    object->properties[u"key"] = fixtureValue(0);
    failHas = hasFailure;
    failProperty = !hasFailure;
    for (int getter = 0; getter < 4; ++getter) {
        bool threw = false;
        try { callGetter(getter, object, 0); }
        catch (const std::runtime_error& error) {
            threw = std::string(error.what()).find("-1") != std::string::npos;
        }
        assert(threw);
    }
    if (hasFailure) {
        bool threw = false;
        try { ob_is_defined(object, L"key"); }
        catch (const std::runtime_error& error) {
            threw = std::string(error.what()).find("-1") != std::string::npos;
        }
        assert(threw);
    }
    PA_DisposeObject(object);
}

static void referenceControls() {
    auto* object = newReference();
    ob_set_o(object, L"object", newReference());
    ob_set_o(object, "object", newReference());
    ob_set_c(object, L"collection", newReference());
    ob_set_b(object, L"bool", true);
    ob_set_i(object, L"int", 42);
    ob_set_n(object, L"number", 42.5);
    assert(ob_is_defined(object, L"object"));
    PA_DisposeObject(object);
}

static void stringify(const std::string& name) {
    auto* object = newReference();
    object->properties[u"input"] = fixtureValue(0);
    failStringify = name == "json-stringify-failure";
    failCopy = name == "json-copy-failure";
    if (failStringify || failCopy) {
        bool threw = false;
        CUTF8String output;
        try { ob_stringify(object, &output); }
        catch (const std::runtime_error& error) {
            threw = std::string(error.what()).find("-1") != std::string::npos;
        }
        assert(threw);
    } else {
        for (int kind : {0, 1, 4, 5, 6}) {
            commandResult = fixtureValue(kind);
            CUTF8String output(reinterpret_cast<const uint8_t*>("unchanged"));
            ob_stringify(object, &output);
            const char* expected = kind <= 1 ?
                "ASCII-\xe6\x97\xa5-\xf0\x9f\x98\x80" : "unchanged";
            assert(std::string(output.begin(), output.end()) == expected);
            releaseVariable(commandResult);
        }
    }
    PA_DisposeObject(object);
}

static void setterFailure() {
    auto* object = newReference();
    object->properties[u"key"] = fixtureValue(0);
    failSet = true;
    for (int setter = 0; setter < 11; ++setter) {
        bool threw = false;
        try {
            switch (setter) {
                case 0: ob_set_s(object, "key", "text"); break;
                case 1: ob_set_s(object, L"key", "text"); break;
                case 2: ob_set_a(object, L"key", L"text"); break;
                case 3: {
                    CUTF16String value(reinterpret_cast<const PA_Unichar*>(u"text"));
                    ob_set_a(object, L"key", &value);
                    break;
                }
                case 4: ob_set_p(object, L"key", newReference()); break;
                case 5: ob_set_o(object, L"key", newReference()); break;
                case 6: ob_set_o(object, "key", newReference()); break;
                case 7: ob_set_c(object, L"key", newReference()); break;
                case 8: ob_set_b(object, L"key", true); break;
                case 9: ob_set_i(object, L"key", 42); break;
                case 10: ob_set_n(object, L"key", 42.5); break;
            }
        } catch (const std::runtime_error& error) {
            threw = std::string(error.what()).find("-1") != std::string::npos;
        }
        assert(threw);
        equal(object->properties.at(u"key").uValue.fString, u"ASCII-\u65e5-\U0001f600");
    }
    PA_DisposeObject(object);
}

static void stringGetFailure(bool stringifyResult) {
    auto* object = newReference();
    object->properties[u"key"] = fixtureValue(0);
    commandResult = fixtureValue(0);
    failStringGet = true;
    CUTF8String output;
    bool threw = false;
    try {
        if (stringifyResult) ob_stringify(object, &output);
        else ob_get_a(object, L"key", &output);
    } catch (const std::runtime_error& error) {
        threw = std::string(error.what()).find("-1") != std::string::npos;
    }
    assert(threw);
    releaseVariable(commandResult);
    PA_DisposeObject(object);
}

template<typename T> static void arrayCopy(std::true_type) {
    T original;
    original.appendUTF16String(reinterpret_cast<const PA_Unichar*>(u"owner"));
    { T alias(original); }
    assert(original.getSize() == 1); // Baseline alias already destroyed this owner.
}
template<typename T> static void arrayCopy(std::false_type) {
    static_assert(!std::is_copy_assignable<T>::value, "No shallow assignment");
}
template<typename T> static void blobCopy(std::true_type) {
    T original;
    const uint8_t input[] = {1, 2, 3};
    original.setBytes(input, sizeof(input));
    { T alias(original); }
    assert(original.getBytesLength() == sizeof(input));
}
template<typename T> static void blobCopy(std::false_type) {
    static_assert(!std::is_copy_assignable<T>::value, "No shallow assignment");
}

int main(int argc, char** argv) {
    assert(argc == 2);
    gCall4D = host;
    const std::string name = argv[1];
    bool passed = true;
    if (name == "sdk-transfer-controls") sdkControls();
    else if (name == "wrapper-controls") wrapperControls();
    else if (name == "scalar-dispatcher-controls") scalarDispatcherControls();
    else if (name == "json-set-failure") setterFailure();
    else if (name == "json-string-get-failure") stringGetFailure(false);
    else if (name == "json-stringify-get-failure") stringGetFailure(true);
    else if (name.find("json-set-") == 0) jsonSetter(name);
    else if (name == "json-picture-setter") pictureSetter(false);
    else if (name == "json-null-picture-control") pictureSetter(true);
    else if (name == "json-reference-controls") referenceControls();
    else if (name == "json-get-text") jsonGetter(0);
    else if (name == "json-get-bool") jsonGetter(1);
    else if (name == "json-get-number") jsonGetter(2);
    else if (name == "json-get-collection") jsonGetter(3);
    else if (name == "json-get-failure") getterFailure(false);
    else if (name == "json-has-failure") getterFailure(true);
    else if (name.find("json-stringify-") == 0 || name == "json-copy-failure") stringify(name);
    else if (name == "array-owning-copy")
        arrayCopy<ARRAY_TEXT>(std::is_copy_constructible<ARRAY_TEXT>{});
    else if (name == "blob-owning-copy")
        blobCopy<C_BLOB>(std::is_copy_constructible<C_BLOB>{});
    else passed = destinationCase(name);
    if (name == "sdk-transfer-controls" || name == "wrapper-controls" ||
        name == "scalar-dispatcher-controls" || name.find("json-") == 0 ||
        name == "array-owning-copy" || name == "blob-owning-copy") {
        passed = strings.empty() && handles.empty() && references.empty();
        unsigned owners = 0;
        for (auto reference : references) owners += reference.second;
        std::cout << name << " strings=" << strings.size() << " handles=" << handles.size()
                  << " references=" << references.size() << " reference_owners=" << owners
                  << (passed ? " PASS\n" : " FAIL\n");
    }
    // Baseline leak evidence is recorded before bounded fixture cleanup.
    for (auto* string : strings) delete[] string;
    for (auto entry : handles) {
        std::free(*entry.first);
        delete entry.first;
    }
    for (auto entry : references) delete entry.first;
    return passed ? 0 : 1;
}
