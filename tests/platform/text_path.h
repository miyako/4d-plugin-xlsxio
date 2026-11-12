#include "4DPluginAPI.h"

// Keep native SDK types, but use the identity UTF-8 path branch in this fixture.
// No Objective-C runtime or filesystem path resolution is permitted.
#undef VERSIONMAC
#define VERSIONMAC 0
