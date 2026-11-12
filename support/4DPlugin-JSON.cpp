#include "4DPlugin-JSON.h"
#include <stdexcept>

static void check_error(PA_ErrorCode error, const char* operation) {
    if(error != eER_NoErr)
        throw std::runtime_error(std::string(operation) + " failed: " +
                                 std::to_string(static_cast<int>(error)));
}

class VariableOwner {
    PA_Variable* value;
public:
    explicit VariableOwner(PA_Variable* variable) : value(variable) {}
    ~VariableOwner() { clear(); }
    void clear() {
        if(value) PA_ClearVariable(value);
        value = nullptr;
    }
    void release() { value = nullptr; }
    VariableOwner(const VariableOwner&) = delete;
    VariableOwner& operator=(const VariableOwner&) = delete;
};

static void set_property_and_clear(PA_ObjectRef obj, PA_Unistring* key, PA_Variable* value) {
    PA_SetObjectProperty(obj, key, *value);
    PA_ErrorCode error = PA_GetLastError();
    PA_DisposeUnistring(key);
    PA_ClearVariable(value);
    check_error(error, "PA_SetObjectProperty");
}

void json_wconv(const wchar_t *value, CUTF16String *u16) {
    
    size_t wlen = wcslen(value);
    
#if VERSIONWIN
    *u16 = CUTF16String((const PA_Unichar *)value, wlen);
#else
    CFStringRef str = CFStringCreateWithBytes(kCFAllocatorDefault, (const UInt8 *)value, wlen*sizeof(wchar_t), kCFStringEncodingUTF32LE, true);
    if(str)
    {
        CFIndex len = CFStringGetLength(str);
        std::vector<uint8_t> buf((len+1) * sizeof(PA_Unichar));
        CFStringGetCharacters(str, CFRangeMake(0, len), (UniChar *)&buf[0]);
        *u16 = CUTF16String((const PA_Unichar *)&buf[0], len);
        CFRelease(str);
    }
#endif
}

static bool get_defined_property(PA_ObjectRef obj, const wchar_t* name, PA_Variable* output) {
    CUTF16String ukey;
    json_wconv(name, &ukey);
    PA_Unistring key = PA_CreateUnistring((PA_Unichar*)ukey.c_str());
    check_error(PA_GetLastError(), "PA_CreateUnistring");
    bool defined = PA_HasObjectProperty(obj, &key);
    PA_ErrorCode error = PA_GetLastError();
    if(error != eER_NoErr || !defined)
    {
        PA_DisposeUnistring(&key);
        check_error(error, "PA_HasObjectProperty");
        return false;
    }
    PA_Variable result = PA_GetObjectProperty(obj, &key);
    error = PA_GetLastError();
    PA_DisposeUnistring(&key);
    check_error(error, "PA_GetObjectProperty");
    *output = result;
    return true;
}

void ob_set_p(PA_ObjectRef obj, const wchar_t *_key, PA_Picture value) {
    
    if(obj)
    {
        if(value)
        {
            PA_Variable v = {};
            CUTF16String ukey;
            json_wconv(_key, &ukey);
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
            
            PA_SetPictureVariable(&v, value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_s(PA_ObjectRef obj, const char *_key, const char *_value) {

    if(obj && _key && _value)
    {
        CUTF8String u8k = CUTF8String((const uint8_t *)_key);
        CUTF8String u8v = CUTF8String((const uint8_t *)_value);
        CUTF16String u16k, u16v;
        
#ifdef _WIN32
        int len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8k.c_str(), u8k.length(), NULL, 0);
        if(len){
            std::vector<uint8_t> buf((len + 1) * sizeof(PA_Unichar));
            if(MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8k.c_str(), u8k.length(), (LPWSTR)&buf[0], len)){
                u16k = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
            }
        }
        u8v = CUTF8String((const uint8_t *)_value);
        len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8v.c_str(), u8v.length(), NULL, 0);
        if(len){
            std::vector<uint8_t> buf((len + 1) * sizeof(PA_Unichar));
            if(MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8v.c_str(), u8v.length(), (LPWSTR)&buf[0], len)){
                u16v = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
            }
        }
#else
        
        CFStringRef str = CFStringCreateWithBytes(kCFAllocatorDefault, u8k.c_str(), u8k.length(), kCFStringEncodingUTF8, true);
        if(str){
            CFIndex len = CFStringGetLength(str);
            std::vector<uint8_t> buf((len+1) * sizeof(PA_Unichar));
            CFStringGetCharacters(str, CFRangeMake(0, len), (UniChar *)&buf[0]);
            u16k = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
            CFRelease(str);
        }
        str = CFStringCreateWithBytes(kCFAllocatorDefault, u8v.c_str(), u8v.length(), kCFStringEncodingUTF8, true);
        if(str){
            CFIndex len = CFStringGetLength(str);
            std::vector<uint8_t> buf((len+1) * sizeof(PA_Unichar));
            CFStringGetCharacters(str, CFRangeMake(0, len), (UniChar *)&buf[0]);
            u16v = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
            CFRelease(str);
        }
#endif
        
        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)u16k.c_str());
        PA_Unistring value = PA_CreateUnistring((PA_Unichar *)u16v.c_str());
        
        PA_Variable v = {};
        PA_SetStringVariable(&v, &value);
        set_property_and_clear(obj, &key, &v);
        
    }

}

void ob_set_s(PA_ObjectRef obj, const wchar_t *_key, const char *_value) {
    
    if(obj)
    {
        if(_value)
        {
            PA_Variable v = {};
            CUTF16String ukey;
            
            json_wconv(_key, &ukey);
            
            CUTF8String u8 = CUTF8String((const uint8_t *)_value);
            CUTF16String u16;
            
#ifdef _WIN32
            int len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8.c_str(), u8.length(), NULL, 0);
            
            if(len){
                std::vector<uint8_t> buf((len + 1) * sizeof(PA_Unichar));
                if(MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8.c_str(), u8.length(), (LPWSTR)&buf[0], len)){
                    u16 = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
                }
            }else{
                u16 = CUTF16String((const PA_Unichar *)L"");
            }
            
#else
            CFStringRef str = CFStringCreateWithBytes(kCFAllocatorDefault, u8.c_str(), u8.length(), kCFStringEncodingUTF8, true);
            if(str){
                CFIndex len = CFStringGetLength(str);
                std::vector<uint8_t> buf((len+1) * sizeof(PA_Unichar));
                CFStringGetCharacters(str, CFRangeMake(0, len), (UniChar *)&buf[0]);
                u16 = CUTF16String((const PA_Unichar *)&buf[0]);
                CFRelease(str);
            }
#endif
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
            PA_Unistring value = PA_CreateUnistring((PA_Unichar *)u16.c_str());
            
            PA_SetStringVariable(&v, &value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_a(PA_ObjectRef obj, const wchar_t *_key, const wchar_t *_value) {
    
    if(obj)
    {
        if(_value)
        {
            PA_Variable v = {};
            CUTF16String ukey;
            CUTF16String uvalue;
            json_wconv(_key, &ukey);
            json_wconv(_value, &uvalue);
            
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
            PA_Unistring value = PA_CreateUnistring((PA_Unichar *)uvalue.c_str());
            
            PA_SetStringVariable(&v, &value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_a(PA_ObjectRef obj, const wchar_t *_key, const CUTF16String *_value) {

    if(obj && _value)
    {
        // _value is already a native UTF-16 (PA_Unichar) buffer with a known
        // length -- do NOT route it through json_wconv/wcslen the way _key
        // is below. json_wconv assumes its input is a genuine platform
        // wchar_t buffer (UTF-32 on macOS), so reinterpreting a CUTF16String's
        // UTF-16 data as wchar_t there would misread the buffer width and
        // corrupt or over-read it. Only _key -- a real wchar_t string
        // literal from the caller -- goes through the normal conversion.
        PA_Variable v = {};
        CUTF16String ukey;
        json_wconv(_key, &ukey);

        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
        PA_Unistring value = PA_CreateUnistring((PA_Unichar *)_value->c_str());

        PA_SetStringVariable(&v, &value);
        set_property_and_clear(obj, &key, &v);
    }
}

void ob_set_o(PA_ObjectRef obj, const wchar_t *_key, PA_ObjectRef value) {
    
    if(obj)
    {
        if(value)
        {
            PA_Variable v = PA_CreateVariable(eVK_Object);
            CUTF16String ukey;
            json_wconv(_key, &ukey);
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
            
            PA_SetObjectVariable(&v, value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_o(PA_ObjectRef obj, const char *_key, PA_ObjectRef value) {

    if(obj && _key)
    {
        if(value)
        {
            CUTF8String u8k = CUTF8String((const uint8_t *)_key);
            CUTF16String u16k;

#ifdef _WIN32
            int len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8k.c_str(), u8k.length(), NULL, 0);
            if(len){
                std::vector<uint8_t> buf((len + 1) * sizeof(PA_Unichar));
                if(MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)u8k.c_str(), u8k.length(), (LPWSTR)&buf[0], len)){
                    u16k = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
                }
            }
#else
            CFStringRef str = CFStringCreateWithBytes(kCFAllocatorDefault, u8k.c_str(), u8k.length(), kCFStringEncodingUTF8, true);
            if(str){
                CFIndex len = CFStringGetLength(str);
                std::vector<uint8_t> buf((len+1) * sizeof(PA_Unichar));
                CFStringGetCharacters(str, CFRangeMake(0, len), (UniChar *)&buf[0]);
                u16k = CUTF16String((const PA_Unichar *)&buf[0], (size_t)len);
                CFRelease(str);
            }
#endif

            PA_Variable v = PA_CreateVariable(eVK_Object);
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)u16k.c_str());

            PA_SetObjectVariable(&v, value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_c(PA_ObjectRef obj, const wchar_t *_key, PA_CollectionRef value) {
    
    if(obj)
    {
        if(value)
        {
            PA_Variable v = PA_CreateVariable(eVK_Collection);
            CUTF16String ukey;
            json_wconv(_key, &ukey);
            PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
            
            PA_SetCollectionVariable(&v, value);
            set_property_and_clear(obj, &key, &v);
        }
    }
}

void ob_set_n(PA_ObjectRef obj, const wchar_t *_key, double value) {
    
    if(obj)
    {
        PA_Variable v = PA_CreateVariable(eVK_Real);
        CUTF16String ukey;
        json_wconv(_key, &ukey);
        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
        
        PA_SetRealVariable(&v, value);
        set_property_and_clear(obj, &key, &v);
    }
}

void ob_set_i(PA_ObjectRef obj, const wchar_t *_key, PA_long32 value) {
    
    if(obj)
    {
        PA_Variable v = PA_CreateVariable(eVK_Longint);
        CUTF16String ukey;
        json_wconv(_key, &ukey);
        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
        
        PA_SetLongintVariable(&v, value);
        set_property_and_clear(obj, &key, &v);
    }
}

void ob_set_b(PA_ObjectRef obj, const wchar_t *_key, bool value) {
    
    if(obj)
    {
        PA_Variable v = PA_CreateVariable(eVK_Boolean);
        CUTF16String ukey;
        json_wconv(_key, &ukey);
        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
        
        PA_SetBooleanVariable(&v, value);
        set_property_and_clear(obj, &key, &v);
    }
}

bool ob_is_defined(PA_ObjectRef obj, const wchar_t *_key) {
    
    bool is_defined = false;
    
    if(obj)
    {
        CUTF16String ukey;
        json_wconv(_key, &ukey);
        PA_Unistring key = PA_CreateUnistring((PA_Unichar *)ukey.c_str());
        
        is_defined = PA_HasObjectProperty(obj, &key);
        PA_ErrorCode error = PA_GetLastError();
        PA_DisposeUnistring(&key);
        check_error(error, "PA_HasObjectProperty");
    }
    return is_defined;
}

bool ob_get_a(PA_ObjectRef obj, const wchar_t *_key, CUTF8String *value) {
    
    bool is_defined = false;
    
    if(obj)
    {
        PA_Variable v;
        is_defined = get_defined_property(obj, _key, &v);
        
        if(is_defined)
        {
            VariableOwner owner(&v);
            if(PA_GetVariableKind(v) == eVK_Unistring)
            {
                PA_Unistring uvalue = PA_GetStringVariable(v);
                PA_ErrorCode error = PA_GetLastError();
                check_error(error, "PA_GetStringVariable");
                
                CUTF16String u = CUTF16String(uvalue.fString, uvalue.fLength);
#ifdef _WIN32
                int len = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)u.c_str(), u.length(), NULL, 0, NULL, NULL);
                
                if(len){
                    std::vector<uint8_t> buf(len + 1);
                    if(WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)u.c_str(), u.length(), (LPSTR)&buf[0], len, NULL, NULL)){
                        *value = CUTF8String((const uint8_t *)&buf[0], (size_t)len);
                    }
                }else{
                    *value = CUTF8String((const uint8_t *)"");
                }
                
#else
                CFStringRef str = CFStringCreateWithCharacters(kCFAllocatorDefault, (const UniChar *)u.c_str(), u.length());
                if(str){
                    
                    size_t size = CFStringGetMaximumSizeForEncoding(
                                                                    CFStringGetLength(str),
                                                                    kCFStringEncodingUTF8) + sizeof(uint8_t);
                    std::vector<uint8_t> buf(size);
                    CFIndex len = 0;
                    CFStringGetBytes(str,
                                     CFRangeMake(
                                                 0,
                                                 CFStringGetLength(str)),
                                     kCFStringEncodingUTF8,
                                     0, true, (UInt8 *)&buf[0], size, &len);
                    
                    *value = CUTF8String((const uint8_t *)&buf[0], len);
                    CFRelease(str);
                }
#endif
            }
        }
    }
    
    return is_defined;
}

bool ob_get_b(PA_ObjectRef obj, const wchar_t *_key) {
    
    bool value = false;
    
    if(obj)
    {
        PA_Variable v;
        if(get_defined_property(obj, _key, &v))
        {
            VariableOwner owner(&v);
            if(PA_GetVariableKind(v) == eVK_Boolean)
            {
                value = PA_GetBooleanVariable(v);
            }
        }
    }
    
    return value;
}

double ob_get_n(PA_ObjectRef obj, const wchar_t *_key) {
    
    double value = 0.0f;
    
    if(obj)
    {
        PA_Variable v;
        if(get_defined_property(obj, _key, &v))
        {
            VariableOwner owner(&v);
            if(PA_GetVariableKind(v) == eVK_Real)
            {
                value = PA_GetRealVariable(v);
            }
        }
    }
    
    return value;
}

PA_CollectionRef ob_get_c(PA_ObjectRef obj, const wchar_t *_key) {
    
    PA_CollectionRef value = NULL;
    
    if(obj)
    {
        PA_Variable v;
        if(get_defined_property(obj, _key, &v))
        {
            VariableOwner owner(&v);
            if(PA_GetVariableKind(v) == eVK_Collection)
            {
                value = PA_GetCollectionVariable(v);
                owner.release();
            }
        }
    }
    return value;
}

void ob_stringify(PA_ObjectRef obj, CUTF8String *value) {
    
    if(!obj || !value) return;
    
    PA_Variable _params[1] = {};
    PA_ObjectRef copy = PA_DuplicateObject(obj);
    check_error(PA_GetLastError(), "PA_DuplicateObject");
    PA_SetObjectVariable(&_params[0], copy);
    VariableOwner argument(&_params[0]);
    PA_Variable vjson = PA_ExecuteCommandByID( /*JSON Stringify */1217, _params, 1);
    PA_ErrorCode error = PA_GetLastError();
    argument.clear();
    check_error(error, "JSON Stringify");
    VariableOwner result(&vjson);
    
    if(PA_GetVariableKind(vjson) != eVK_Unistring)
    {
        // "JSON Stringify" didn't return a string (unexpected failure) --
        // bail out instead of reading a Unistring out of a variable that
        // may not hold one.
        return;
    }
    
    PA_Unistring ujson = PA_GetStringVariable(vjson);
    error = PA_GetLastError();
    check_error(error, "PA_GetStringVariable");
    
#ifdef _WIN32
    int len = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)ujson.fString, ujson.fLength, NULL, 0, NULL, NULL);
    
    if(len){
        std::vector<uint8_t> buf(len + 1);
        if(WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)ujson.fString, ujson.fLength, (LPSTR)&buf[0], len, NULL, NULL)){
            *value = CUTF8String((const uint8_t *)&buf[0], (size_t)len);
        }
    }else{
        *value = CUTF8String((const uint8_t *)"");
    }
    
#else
    CFStringRef str = CFStringCreateWithCharacters(kCFAllocatorDefault, (const UniChar *)ujson.fString, ujson.fLength);
    if(str){
        
        size_t size = CFStringGetMaximumSizeForEncoding(CFStringGetLength(str), kCFStringEncodingUTF8) + sizeof(uint8_t);
        std::vector<uint8_t> buf(size);
        CFIndex len = 0;
        CFStringGetBytes(str, CFRangeMake(0, CFStringGetLength(str)), kCFStringEncodingUTF8, 0, true, (UInt8 *)&buf[0], size, &len);
        
        *value = CUTF8String((const uint8_t *)&buf[0], len);
        CFRelease(str);
    }
    
#endif
    
}
