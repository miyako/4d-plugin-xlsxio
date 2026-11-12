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
selectors abort. JSON helper code and unused wrappers are compiled, not treated
as evidence of exported reachability. `ARRAY_TEXT`'s normal output path and
`C_TEXT`'s input conversion/output setter are separately executed as controls.
Selectors 1 and 2 also run through the real dispatcher for failed synthetic
open and invalid JSON respectively, checking exact error text and scalar return
replacement without entering workbook parsing or writer operations.

## Baseline and expected results

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

Baseline: 5/11 pass, 6 named failures. Fixed: 11/11 pass with zero outstanding
tracked strings or handles. Leak balances are inspected **before** bounded
fixture cleanup; the original defect is a leak, not a sanitizer UAF failure.
Controls verify exact ASCII, BMP, supplementary (two UTF-16 units), and empty
content, borrowed array reads, index-zero acceptance, element replacement,
shrink/clear, and rejected type/data/index/null-string setters retaining
caller ownership. Both source variants use identical host contracts and
assertions. No allocation/resize failure semantics are invented.

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
  implicit-copy capabilities are not exported defects: no production consumers
  or copies exist here. `ARRAY_TEXT` copies borrowed array text into independent
  C++ strings and transfers newly created output strings without disposal.
  This audit does not establish OOM/failed-resize correctness of the dormant
  helper or its native path/picture behavior.
* `support/4DPlugin-JSON.{h,cpp}` are absent from both targets and have no plugin
  callers. Static review identifies dormant hazards: initialized empty string
  owners overwritten by shallow scalar setters; successful property results
  not cleared by the scalar getters; collection-result lifetime/handoff lacking
  an actual consumer contract; and picture input handed to a cleared variable
  without an independently documented ownership contract. `ob_stringify`
  clears its argument and successful string result but does not capture command
  error before interpreting the result. These are **not** claimed fixed or
  runtime-validated, nor evidence of an XLSXIO command crash. Re-enabling these
  helpers requires a separate actual-consumer ownership contract and fixtures.

## Exact SDK provenance and limitations

All files below and every other tracked `4D Plugin API/` file remain unchanged:

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

The fixture does not run native 4D on either platform, Windows conversion APIs,
Objective-C paths, real XLSXIO parsing/writing, callbacks, pictures, or OOM.
It does not reproduce a Windows host crash. No project build, Deployment
configuration, release, signing, notarization, binary packaging/publication,
tag, or manual workflow dispatch is needed. Existing release automation is
tag/manual-only and version bump is manual-only; the new CI is unsigned,
read-only, and limited to these synthetic fixtures on pull requests and
`master` pushes.
