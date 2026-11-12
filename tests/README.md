# Synthetic ownership regression

Run on macOS with the existing Apple Clang and Python 3:

```sh
python3 tests/run_ownership.py --build /tmp/xlsxio-ownership
```

The runner compiles the complete unchanged bundled SDK, the complete plugin
translation unit, all four bundled wrapper translation units, and the dormant
JSON helpers. It links the repository's existing universal `a/libjsoncpp.a`;
no dependencies are downloaded. ASan and UBSan are fatal. Platform headers are
replaced by the in-memory implementations in `platform/`. The `C_TEXT` object
alone selects the identity UTF-8 path branch after native SDK types are defined;
no Objective-C runtime, HFS/POSIX path resolution, or actual workbook is used.
The executable's undefined imports are checked before execution against
standard-library, sanitizer, and bounded C runtime allowlists.

The test calls the real SDK `FourDPackex` entrypoint, real `PluginMain` dispatcher
selector 3, and real `XLSX_SHEET_NAMES` consumer. XLSXIO is replaced by a synthetic
reader/list; other reader/writer operations abort. The host callback allows only
string creation/copy/disposal, handle operations, and variable cleanup. Unexpected
selectors abort. JSON helpers also execute directly in bounded ownership
fixtures, without claiming exported reachability. `ARRAY_TEXT`'s normal output path and
`C_TEXT`'s input conversion/output setter are separately executed as controls.
Selectors 1 and 2 also run through the real dispatcher for failed synthetic
open and invalid JSON respectively, checking exact error text and scalar return
replacement without entering workbook parsing or writer operations.

## Exported-array regression (merged in #2)

Baseline: `58e8210850826325ca3330e26e3abe91dfc9f4ad`. Use `git archive` to
materialize that immutable source tree, then pass it to the **same final**
runner using `--source /path/to/baseline --build /path/to/baseline-build`.
The runner prints SHA-256 identities for the fixtures and compiled sources.

| Named case | Baseline | Fixed |
|---|---|---|
| `prepopulated-caller` | FAIL: 3 strings, 1 handle retained | PASS |
| `empty-caller` | FAIL: 1 string, 1 handle retained | PASS |
| `undefined-destination` | PASS | PASS |
| `rejected-kind` | PASS | PASS |
| `empty-result` | FAIL: 3 strings, 1 handle retained | PASS |
| `failed-open` | FAIL: 3 strings, 1 handle retained | PASS |
| `failed-list` | FAIL: 3 strings, 1 handle retained | PASS |
| `repeated-reuse` | FAIL: 78 strings, 16 handles retained after 16 calls | PASS |
| `sdk-transfer-controls` | PASS | PASS |
| `wrapper-controls` | PASS | PASS |
| `scalar-dispatcher-controls` | PASS | PASS |

This original 11-case subset remains in the expanded suite. Baseline: 5/11 pass,
6 named failures. Fixed: 11/11 pass with zero outstanding
tracked strings or handles. Leak balances are inspected **before** bounded
fixture cleanup; the original defect is a leak, not a sanitizer UAF failure.
Controls verify exact ASCII, BMP, supplementary (two UTF-16 units), and empty
content, borrowed array reads, index-zero acceptance, element replacement,
shrink/clear, and rejected type/data/index/null-string setters retaining
caller ownership. Both source variants use identical host contracts and
assertions. No allocation/resize failure semantics are invented.

## Dormant-helper follow-up

Follow-up baseline: `16ab49e17558e1712534acb7c674027055bba1e0`, which already
contains the exported-array fix. The **same final expanded harness** runs against
an immutable `git archive` of this baseline and the fixed actual sources.
Follow-up baseline: **14/32 pass, 18 named failures**. Fixed: **32/32 pass**.
The original 11 exported/control cases pass in both follow-up revisions.

| Additional named case | Follow-up baseline | Fixed |
|---|---|---|
| `json-set-narrow` | FAIL: 12 empty strings leaked | PASS |
| `json-set-wide` | FAIL: 12 empty strings leaked | PASS |
| `json-set-wide-value` | FAIL: 12 empty strings leaked | PASS |
| `json-set-u16` | FAIL: 12 empty strings leaked | PASS |
| `json-picture-setter` | FAIL: 3 allocated empty picture references leaked | PASS |
| `json-null-picture-control` | PASS: null empty-picture representation | PASS |
| `json-reference-controls` | PASS: existing consuming reference handoffs | PASS |
| `json-get-text` | FAIL: 6 strings and 9 reference owners leaked | PASS |
| `json-get-bool` | FAIL: 6 strings and 9 reference owners leaked | PASS |
| `json-get-number` | FAIL: 6 strings and 9 reference owners leaked | PASS |
| `json-get-collection` | FAIL: 6 strings and 6 reference owners leaked on mismatches | PASS |
| `json-get-failure` | FAIL: interprets unspecified failed property output | PASS |
| `json-has-failure` | FAIL: reports failed presence check as absence | PASS |
| `json-stringify-controls` | PASS: normal/raw-zero text and owning mismatches | PASS |
| `json-stringify-failure` | FAIL: interprets unspecified failed command output | PASS |
| `json-copy-failure` | FAIL: adopts failed object-duplicate result | PASS |
| `array-owning-copy` | FAIL: ASan heap-use-after-free | PASS: copying disabled |
| `blob-owning-copy` | FAIL: ASan heap-use-after-free | PASS: copying disabled |
| `json-set-failure` | FAIL: suppresses property-set error | PASS |
| `json-string-get-failure` | FAIL: interprets failed string-conversion output | PASS |
| `json-stringify-get-failure` | FAIL: interprets failed string-conversion output | PASS |

String setters exercise ASCII, BMP, supplementary, and empty values through
repeated replacement. Getter matrices include normal and raw-zero text, boolean,
real, object, collection, and picture results; successful scalar and mismatched
owning results are cleared. Matching collection results transfer one returned
owner to the fixture caller, which disposes it without invalidating the object's
property. This is the same handoff pattern used by the bundled
`PA_CreateCollection`: it returns the reference from its command-result variable
without clearing that transferred owner.

Property/set/stringify failures are diagnostic exceptions containing the captured
SDK error code. Error capture precedes cleanup that could reset the code. Failed
property/command/string outputs deliberately have no ownership contract and are
never inspected or cleared; independent successful arguments/results still are.
Successful owning results have noncopyable scope guards, including exceptional
conversion exits. The presence helper also distinguishes errors from absence.

The picture fixture tests both allocated and null empty representations; it does
not claim the native host always allocates an empty picture. Avoiding the unused
empty creation is correct in either representation. Reference-valued setters
preserve their existing **consuming** behavior on non-null input branches,
including local cleanup after a set error. No external caller borrowing contract
is invented. Property retention/refcounts are synthetic host contracts, not
native picture-decoder or engine measurements.

Both raw-copy fixtures construct an owning source and shallow copy using the
actual baseline wrapper. Destroying the copy then reading the source produces
ASan UAF. Disabling copying follows existing `C_TEXT` prior art; compile-time
assertions also reject copy assignment. No production consumers/copies exist
here, and all target/helper translation units still compile. The engine SDK
implementation/ABI is untouched; only the two wrapper class headers change,
without changing their object layouts.
Normal blob storage still executes in the wrapper control; value-owned `CBytes`
copying remains allowed and is checked for independent storage.

## Ownership and reachability inventory

Both command manifests (`manifest.json` and the checked-in bundle resource)
declare the same complete three-command surface: text/longint/text to text,
text/text/longint/longint/text, and text/text-array. The project files compile
only `4DPlugin-XLSXIO.cpp`, `4DPluginAPI.c`, `C_TEXT.cpp`, and `C_LONGINT.cpp`.
The only 4D method consumer is `test/Project/Sources/Methods/TEST.4dm`; it passes
an initialized text array to `XLSX SHEET NAMES`. That real method is inspected,
not executed: it accesses Desktop workbooks.

* Exported selector 3 allocates an independent string per sheet name, then
  transfers each descriptor to its array. There is no immediate disposal bug.
  The fresh array handoff must set `clearOldValue=1`: the exact SDK comment on
  `PA_SetVariableParameter` requires this when using `PA_CreateVariable`.
  Rejected output kinds already clear the new array and remain unchanged.
* Exported selectors 1 and 2 use noncopyable `C_TEXT` owners backed by independent
  C++ strings. Parameter/result setters call the engine-backed content setter,
  not the array descriptor-transfer setter. No scalar `PA_SetStringVariable`,
  method-result, property, collection, picture, or callback ownership path is
  reachable from these commands. XLSXIO sheet names are borrowed until the next
  iteration and copied synchronously; sheet cell allocations are freed after
  synchronous JSON insertion.
* `ARRAY_TEXT` and `C_BLOB` are dormant in both native targets. Their raw-pointer
  implicit-copy defects are demonstrated directly, not described as exported
  crashes; copying is now disabled. `ARRAY_TEXT` copies borrowed array text into independent
  C++ strings and transfers newly created output strings without disposal.
  This audit does not establish OOM/failed-resize correctness of the dormant
  helper or its native path/picture behavior.
* `support/4DPlugin-JSON.{h,cpp}` are absent from both targets and have no plugin
  callers. Their demonstrated empty-owner overwrite, successful result/mismatch
  cleanup, and failure-before-adoption defects are now fixed and directly
  exercised as above. Matching collection handoff and consuming reference inputs
  are preserved controls, not falsely classified as leaks or borrowing fixes.
  These are not evidence of an XLSXIO exported-command crash.

## Exact SDK provenance and limitations

All engine implementation/ABI files below, `.def`, and root SDK readme remain
unchanged. The follow-up changes only wrapper headers `Classes/ARRAY_TEXT.h`
and `Classes/C_BLOB.h` within this directory; their ownership restrictions are
not changes to the engine implementation or packed ABI types.

| Path | Git blob |
|---|---|
| `4DPluginAPI.c` | `62f745e91d81ef2c3c901c656ea8e175d3b173e8` |
| `4DPluginAPI.h` | `528bf1d2552df08d6a8b1bc95011a9b72934ad5c` |
| `PublicTypes.h` | `9ae20b001bff9128f690fff3e6c3f873b6238333` |
| `PrivateTypes.h` | `a3098cb662cdbc43ba574fdef72a059c4c6d1cd4` |
| `EntryPoints.h` | `da428d5f6ca23b0a30371c7288a01d29aa3c466a` |
| `Flags.h` | `7a5a35539f27430a3fd9cb711d6b789051316471` |

The exact implementation accepts array insertion only for Unicode-array kind,
non-null data, index `0 <= i <= fNbElements`, and non-null descriptor pointer;
it disposes the previous element and shallow-assigns the incoming descriptor.
The getter is borrowed. The scalar string setter is shallow and does not
release an existing value. Successful property/collection getters explicitly
require variable cleanup. SDK version labels are not used as ownership proof.

Guidance reviewed: `miyako/skills/4d-skills/skills/4dplugin/SKILL.md` at immutable
`e6b492fcb81e5e1a7f7db723594b420e63bf6eb9` and current guidance on 2026-11-12.

**Deferred, not confirmed safe or fixed:** actual engine allocation/resize
failures have unspecified failed outputs here. In particular, dormant
`ARRAY_TEXT` resize/create rejection cleanup and `C_BLOB` destination replacement
after a failed allocation require a documented engine failure contract before
fault injection can support a native claim. The callback has no allocator
failure model; string/picture/handle creation failures in other helpers likewise
are not qualified by these success/result-failure fixtures. The intended ownership expectations of future external callers
of consuming picture/object/collection setters cannot be inferred without
those callers; existing consuming behavior is preserved and documented.

The fixture does not run native 4D on either platform, Windows conversion APIs,
Objective-C paths, real XLSXIO parsing/writing, callbacks, native picture
decoders, or OOM.
It does not reproduce a Windows host crash. No project build, Deployment
configuration, release, signing, notarization, binary packaging/publication,
tag, or manual workflow dispatch is needed. Existing release automation is
tag/manual-only and version bump is manual-only; the new CI is unsigned,
read-only, and limited to these synthetic fixtures on pull requests and
`master` pushes.
