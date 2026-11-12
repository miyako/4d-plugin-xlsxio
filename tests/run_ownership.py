#!/usr/bin/env python3
"""Compile actual plugin/SDK sources; execute only in-memory ownership fixtures."""
import argparse
import hashlib
import pathlib
import subprocess

parser = argparse.ArgumentParser()
parser.add_argument("--source", type=pathlib.Path, default=pathlib.Path(__file__).resolve().parents[1])
parser.add_argument("--build", type=pathlib.Path, required=True)
args = parser.parse_args()
source = args.source.resolve()
build = args.build.resolve()
tests = pathlib.Path(__file__).resolve().parent
build.mkdir(parents=True, exist_ok=True)
for fixture in sorted(tests.rglob("*")):
    if fixture.is_file():
        print("FIXTURE", fixture.relative_to(tests),
              hashlib.sha256(fixture.read_bytes()).hexdigest(), flush=True)
sources = [
    "4D Plugin API/4DPluginAPI.c",
    "4D Plugin API/Classes/C_TEXT.cpp",
    "4D Plugin API/Classes/C_LONGINT.cpp",
    "4DPlugin-XLSXIO.cpp",
    "4D Plugin API/Classes/ARRAY_TEXT.cpp",
    "4D Plugin API/Classes/C_BLOB.cpp",
    "support/4DPlugin-JSON.cpp",
]
common = [
    "clang++", "-std=c++14", "-g", "-O1", "-fno-omit-frame-pointer",
    "-fsanitize=address,undefined", "-fno-sanitize-recover=all", "-Wno-multichar",
    "-Wno-deprecated-declarations",
    "-I" + str(tests / "platform"), "-I" + str(source),
    "-I" + str(source / "4D Plugin API"),
    "-I" + str(source / "4D Plugin API/Classes"), "-I" + str(source / "include"),
    "-include", "cstring", "-include", "ctime",
]
objects = []
for index, relative in enumerate(sources + [str(tests / "ownership.cpp")]):
    path = source / relative
    print("SOURCE", relative, hashlib.sha256(path.read_bytes()).hexdigest(), flush=True)
    target = build / f"{index}.o"
    command = common + ["-x", "c++"]
    if relative == "4D Plugin API/Classes/C_TEXT.cpp":
        command += ["-include", str(tests / "platform/text_path.h")]
    command += ["-c", str(path), "-o", str(target)]
    subprocess.run(command, check=True)
    objects.append(str(target))
binary = build / "ownership"
subprocess.run(common + objects + [str(source / "a/libjsoncpp.a"),
                                  "-Wl,-dead_strip", "-o", str(binary)], check=True)
imports = subprocess.check_output(["nm", "-u", str(binary)], text=True)
(build / "imports.txt").write_text(imports)
allowed = {
    "__Unwind_Resume", "___assert_rtn", "___gxx_personality_v0",
    "___stack_chk_fail", "___stack_chk_guard", "__tlv_bootstrap",
    "_abort", "_calloc", "_free", "_localeconv", "_malloc", "_memchr", "_memcmp",
    "_memcpy", "_memmove", "_memset", "_snprintf", "_sscanf",
    "_strchr", "_strcmp", "_strlen", "_time", "_wcslen",
}
runtime_prefixes = (
    "__Z", "___asan_", "___ubsan_", "___cxa_", "___sanitizer_",
)
for symbol in imports.splitlines():
    symbol = symbol.strip()
    if symbol not in allowed and not symbol.startswith(runtime_prefixes):
        raise RuntimeError("Unexpected platform import: " + symbol)
cases = [
    "prepopulated-caller", "empty-caller", "undefined-destination", "rejected-kind",
    "empty-result", "failed-open", "failed-list", "repeated-reuse",
    "sdk-transfer-controls", "wrapper-controls", "scalar-dispatcher-controls",
    "json-set-narrow", "json-set-wide", "json-set-wide-value", "json-set-u16",
    "json-picture-setter", "json-null-picture-control", "json-reference-controls",
    "json-get-text", "json-get-bool", "json-get-number", "json-get-collection",
    "json-get-failure", "json-has-failure", "json-stringify-controls",
    "json-stringify-failure", "json-copy-failure", "array-owning-copy", "blob-owning-copy",
    "json-set-failure", "json-string-get-failure", "json-stringify-get-failure",
]
failures = []
for case in cases:
    print("CASE", case, flush=True)
    result = subprocess.run([str(binary), case], check=False)
    if result.returncode:
        failures.append(case)
print(f"RESULT {len(cases) - len(failures)}/{len(cases)} PASS; failures={failures}", flush=True)
raise SystemExit(bool(failures))
