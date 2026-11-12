#pragma once
#include <cstdlib>
using CGContextRef = void*;
inline void CGContextScaleCTM(CGContextRef, double, double) { std::abort(); }
inline void CGContextTranslateCTM(CGContextRef, double, double) { std::abort(); }
