#pragma once

#include <windows.h>
#include <processthreadsapi.h>

#include <string>
#include <vector>
#include <iostream>

// This import statement causes generation of OPOSScanner.tlh and OPOSScanner.tli
#import "progid:OPOS.Scanner"

// This import statement informs the IDE of the COM interface and everything in the
// namespace OposScanner_CCO, but is not required for compilation.
#import "libid:ccb90180-b81e-11d2-ab74-0040054c3719"

#define USBCOM_SCANNER_ANY           "RS232Imager"
#define USBOEM_SCANNER_HANDHELD      "USBHHScanner"
#define USBOEM_SCANNER_FIXED_RETAIL  "USBScanner"
#define RS232_SCANNER_ANY            "RS232Scanner"
#define RS232SC_SCANNER_FIXED_RETAIL "SCRS232Scanner"
