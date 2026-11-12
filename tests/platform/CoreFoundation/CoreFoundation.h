#pragma once

#include <algorithm>
#include <cassert>
#include <codecvt>
#include <cstring>
#include <locale>
#include <string>

using CFIndex = long;
using UInt8 = unsigned char;
using UniChar = unsigned short;
using BOOL = bool;
struct SyntheticString { std::u16string text; };
using CFStringRef = SyntheticString*;
constexpr void* kCFAllocatorDefault = nullptr;
constexpr int kCFStringEncodingUTF8 = 1;
constexpr int kCFStringEncodingUTF32LE = 2;
struct CFRange { CFIndex location, length; };
inline CFRange CFRangeMake(CFIndex start, CFIndex length) { return {start, length}; }
inline CFStringRef CFStringCreateWithBytes(void*, const UInt8* bytes, CFIndex length,
                                          int encoding, bool) {
    std::u16string text;
    if (encoding == kCFStringEncodingUTF8) {
        std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> converter;
        text = converter.from_bytes(reinterpret_cast<const char*>(bytes),
                                    reinterpret_cast<const char*>(bytes) + length);
    } else {
        assert(encoding == kCFStringEncodingUTF32LE && length % 4 == 0);
        for (CFIndex i = 0; i < length; i += 4) {
            unsigned code = bytes[i] | (unsigned(bytes[i + 1]) << 8) |
                            (unsigned(bytes[i + 2]) << 16) | (unsigned(bytes[i + 3]) << 24);
            if (code > 0xffff) {
                code -= 0x10000;
                text.push_back(char16_t(0xd800 + (code >> 10)));
                text.push_back(char16_t(0xdc00 + (code & 0x3ff)));
            } else {
                text.push_back(char16_t(code));
            }
        }
    }
    return new SyntheticString{text};
}
inline CFStringRef CFStringCreateWithCharacters(void*, const UniChar* chars, CFIndex length) {
    std::u16string text;
    for (CFIndex i = 0; i < length; ++i) text.push_back(char16_t(chars[i]));
    return new SyntheticString{text};
}
inline CFIndex CFStringGetLength(CFStringRef str) { return str->text.size(); }
inline void CFStringGetCharacters(CFStringRef str, CFRange range, UniChar* output) {
    assert(range.location >= 0 && range.length >= 0);
    assert(size_t(range.location + range.length) <= str->text.size());
    std::copy_n(str->text.data() + range.location, range.length, output);
}
inline CFIndex CFStringGetMaximumSizeForEncoding(CFIndex length, int encoding) {
    assert(encoding == kCFStringEncodingUTF8);
    return length * 3;
}
inline CFIndex CFStringGetBytes(CFStringRef str, CFRange range, int encoding, int, bool,
                               UInt8* output, CFIndex capacity, CFIndex* used) {
    assert(encoding == kCFStringEncodingUTF8);
    std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> converter;
    auto bytes = converter.to_bytes(str->text.substr(range.location, range.length));
    assert(size_t(capacity) >= bytes.size());
    std::memcpy(output, bytes.data(), bytes.size());
    *used = bytes.size();
    return range.length;
}
inline void CFRelease(CFStringRef str) { delete str; }
