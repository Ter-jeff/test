# Error Code Rules

## Code Format — `ECCSSSNNN`

```
E H I 1 5 5 0 0 1
^                  E   — Level        (I/W/E/C, 1 letter)
  ^ ^              CC  — Category     (AA–ZZ, 2 letters from [Description] on EnumErrorCategory)
       ^ ^ ^       SSS — Sub-code     (100–791, (int)EnumErrorBehavior * 100 + (int)EnumErrorTarget, zero-padded)
             ^ ^ ^ NNN — Error number (001–999, 3 digits, sequential within category+behavior+target)
```

Total length: **9 characters**, no separators.

---

## Field Definitions

| Field | Width | Range | Source |
| :--- | :--- | :--- | :--- |
| `E` | 1 alpha | C, E, W, I | First character of `ErrorLevel.ToString()` |
| `CC` | 2 alpha | AA–ZZ | `[Description("XX")]` attribute on `EnumErrorCategory` member |
| `SSS` | 3 digit | 100–791 | `(int)EnumErrorBehavior * 100 + (int)EnumErrorTarget`, zero-padded to 3 digits |
| `NNN` | 3 digit | 001–999 | `int code` constructor parameter, zero-padded to 3 digits |

`EnumErrorBehavior` values run 1 (`Duplicate`) through 7 (`Info`) — see [Behavior Codes](#behavior-codes-enumerrorbehavior) below.
`EnumErrorTarget` is a large flat enum (see `EnumErrorTarget.cs`); its ordinal (0-based) is the last one or two digits of `SSS`.

The `FullCode` property on `ErrorCode` computes the 9-char string at runtime:

```csharp
$"{eStr}{ccStr}{sssStr}{nnnStr}"
// e.g. "E" + "PM" + "155" + "001" → "EPM155001"
// where 155 = (int)EnumErrorBehavior.Duplicate (1) * 100 + (int)EnumErrorTarget.Pin (55)
```

---

## Level Codes (E)

| Code | Level Name | Enum Value | Description |
| :--- | :--- | :--- | :--- |
| `I` | Info | `ErrorLevel.Info = 0` | Informational audit logging and tracking details. |
| `W` | Warning | `ErrorLevel.Warning = 1` | Non-blocking issues or potential file structural anomalies. |
| `E` | Error | `ErrorLevel.Error = 2` | Standard validation failures requiring user correction. |
| `C` | Critical | `ErrorLevel.Critical = 3` | Catastrophic failure that aborts the processing loop. |

---

## Category Codes (CC)

Defined in `EnumErrorCategory`. The `[Description]` attribute on each member provides the 2-letter CC used in `FullCode`.

| CC | `EnumErrorCategory` Member | Static Error Class | Description |
|----|----------------------------|--------------------|-------------|
| `BA` | `Basic` | `BasicErrorType` | Basic pre-check errors |
| `BC` | `BinCut` | `BinCutErrorType` | BinCut errors |
| `BL` | `EfuseCheckCmdLib` | `EfuseCheckCmdLibError` | EfuseCheckCmdLib errors |
| `DI` | `DuplicateInstance` | `DuplicateInstance` | Duplicate instance errors |
| `DT` | `DuplicateTestNumber` | `DuplicateTestNumber` | Duplicate test number errors |
| `EF` | `EFuse` | `EFuseErrorType` | EFuse definition errors |
| `EN` | `Enum` | `EnumErrorType` | General / legacy enum errors |
| `EV` | `Evs` | `EvsErrorType` | EVS errors |
| `FM` | `FlowMain` | `FlowMainErrorType` | Flow-main errors |
| `HA` | `Harvest` | `HarvestErrorType` | Harvest errors |
| `HI` | `HardIp` | `HardIpErrorType` | HardIP sheet errors |
| `HT` | `Htol` | `HtolErrorType` | HTOL errors |
| `MB` | `Mbist` | `MbistErrorType` | MBIST errors |
| `MF` | `MainFlow` | `MainFlowErrorType` | Main flow errors |
| `PA` | `PreAction` | `PreActionErrorType` | Pre-action errors |
| `PI` | `PatInfo` | `PatInfoType` | PatInfo errors |
| `PM` | `PatternMissing` | `PatternMissing` | Missing pattern errors |
| `RT` | `Rtos` | `RtosErrorType` | RTOS errors |
| `SC` | `Scan` | `ScanErrorType` | Scan errors |
| `SN` | `Ssn` | `SsnErrorType` | SSN errors |
| `UF` | `UfInstance` | `UfInstanceErrorType` | UF instance errors |

---

## IErrorCode Interface

```csharp
public interface IErrorCode
{
    string     FullCode        { get; }  // ECCSSSNNN, e.g. "EPM155001"
    string     MessageTemplate { get; }  // {0}/{1}/… placeholders — no runtime values
    EnumErrorLevel ErrorLevel  { get; }  // fixed severity — callers do NOT pass this separately
    string     Guidance        { get; }  // debugging help for the user
}
```

`ErrorCode` exposes the raw code plus the computed `FullCode`:

```csharp
public int    Code     { get; }  // NNN — the raw 3-digit sequential error number
public string FullCode { get; }  // ECCSSSNNN — computed 9-char string
```

The component properties that feed `FullCode` are also publicly accessible:

```csharp
public EnumErrorCategory  EnumErrorCategory { get; }
public EnumErrorBehavior? EnumErrorBehavior { get; }
public EnumErrorTarget?   EnumErrorTarget   { get; }
```

There is no `Name` or `Category` property — those were dropped when `EnumErrorSubGroup` was retired in favor of `EnumErrorBehavior` + `EnumErrorTarget`.

---

## ErrorCode Constructor

```csharp
public ErrorCode(
    EnumErrorCategory enumErrorCategory,  // required — from EnumErrorCategory enum
    EnumErrorBehavior enumErrorBehavior,  // required — from EnumErrorBehavior enum
    EnumErrorTarget   enumErrorTarget,    // required — from EnumErrorTarget enum
    int               code,               // required — NNN (001–999, sequential within category+behavior+target)
    string            template,           // required — use {0},{1},… placeholders
    EnumErrorLevel    enumErrorLevel,     // required — fixed severity for this error
    string            guidance = null)    // optional — debugging help
```

### Definition example

```csharp
// arg {0} = missingPattern
public static readonly ErrorCode E_MissingPatternFile = new ErrorCode(
    enumErrorCategory: EnumErrorCategory.PatternMissing,
    enumErrorBehavior: EnumErrorBehavior.Missing,
    enumErrorTarget:   EnumErrorTarget.File,
    code:              1,
    template:          "{0} doesn't exist in pattern folder.",
    enumErrorLevel:    EnumErrorLevel.Error,
    guidance:          "Verify the pattern file path and name against the test-plan entry. "
                     + "Run the pattern compilation step if the file is generated. "
                     + "Check source-control to confirm the file was committed and not accidentally excluded.");
```

---

## MessageTemplate Rules

> **The template is a format string, not an interpolated string.**
> Runtime values are inserted at the **call site**, never at class initialisation.

| Rule | Wrong ❌ | Right ✓ |
|------|---------|---------|
| Use `{0}` placeholders | `template: $"{missingPattern} doesn't exist"` | `template: "{0} doesn't exist in pattern folder."` |
| Use `{0}` placeholders | `template: patSetStatus.PatternComment` | `template: "{0}"` |
| Document which arg maps to which `{N}` | *(nothing)* | `// arg {0} = missingPattern` comment above the field |
| Keep template generic | `template: "abc.pat not found"` | `template: "{0} doesn't exist in pattern folder."` |

**Why?** The static field is initialised once at program start. `missingPattern` and
`patSetStatus` are runtime variables that don't exist at that point.

---

## FormatMessage — Producing the Final String

`ErrorCode.FormatMessage(params object[] args)` calls `string.Format(MessageTemplate, args)`.

```csharp
// Definition
// arg {0} = missingPattern
public static readonly ErrorCode E_MissingPatternFile = new ErrorCode(
    enumErrorCategory: EnumErrorCategory.PatternMissing,
    enumErrorBehavior: EnumErrorBehavior.Missing,
    enumErrorTarget:   EnumErrorTarget.File,
    code:     1,
    template: "{0} doesn't exist in pattern folder.", ...);

// Call site — pass the runtime value here
string message = PatternMissing.E_MissingPatternFile.FormatMessage(missingPattern);
// → "abc.pat doesn't exist in pattern folder."

// Multi-arg example
// arg {0} = bitName, arg {1} = blockName
public static readonly ErrorCode E_DupBitName = new ErrorCode(
    enumErrorCategory: EnumErrorCategory.HardIp,
    enumErrorBehavior: EnumErrorBehavior.Duplicate,
    enumErrorTarget:   EnumErrorTarget.Bit,
    code:     2,
    template: "Register bit name '{0}' is duplicated in block '{1}'.", ...);

string message = HardIpErrorType.E_DupBitName.FormatMessage(bitName, blockName);
// → "Register bit name 'CLK_EN' is duplicated in block 'PLL_BLOCK'."
```

---

## AddError — Two Usage Patterns

### Pattern A — standard path (no special fields on `Error`)

`ErrorLevel` is read from the `IErrorCode`; you only supply location + formatted message.

```csharp
ErrorReportManager.AddError(
    errorCode: HardIpErrorType.E_DupBitName,
    sheetName: "HardIp_Core",
    rowNum:    42,
    columnNum: 5,
    message:   HardIpErrorType.E_DupBitName.FormatMessage(bitName, blockName));
```

`ErrorReportManager.AddError` signature (no `errorLevel` parameter):

```csharp
public static void AddError(
    IErrorCode errorCode,
    string     sheetName,
    int        rowNum,
    int        columnNum,
    string     message,
    params string[] comments)
{
    // errorLevel comes from errorCode, not from the caller
    if (errorCode.ErrorLevel == EnumErrorLevel.Critical && StopByCritical)
        Environment.Exit(-1);

    AddError(new Error
    {
        ErrorType  = errorCode,
        ErrorLevel = errorCode.ErrorLevel,   // ← taken from IErrorCode
        SheetName  = sheetName,
        RowNum     = rowNum,
        ColNum     = columnNum < 1 ? 0 : columnNum,
        Message    = message,
        Comments   = comments.ToList()
    });
}
```

### Wrong Pattern ❌

Use this when you need to set fields that `AddError` doesn't expose (e.g. `Error.Pattern`,
used by `RemoveNonUsedPattern`).

```csharp
var error = new Error
{
    ErrorType  = PatternMissing.E_MissingPatternFile,
    ErrorLevel = PatternMissing.E_MissingPatternFile.ErrorLevel,   // from IErrorCode
    SheetName  = "",
    RowNum     = 1,
    ColNum     = 0,
    Message    = PatternMissing.E_MissingPatternFile.FormatMessage(missingPattern),
    Pattern    = missingPattern   // extra field not available via AddError
};
```

---

## Guidance Rules

The `Guidance` string is shown to the user when the error fires.

### Must include

| # | Rule | Bad ❌ | Good ✓ |
|---|------|--------|--------|
| 1 | **Root cause** | "Error occurred." | "The pattern file name in the test plan does not match the compiled file name." |
| 2 | **Where to look** | "Check the file." | "Open the HardIp sheet and go to the 'Pattern' column." |
| 3 | **How to fix** | "Fix the value." | "Rename the file to match the test plan entry, or update the test plan." |

### Style

- Start each sentence with a verb: **Check**, **Verify**, **Ensure**, **Open**, **Compare**, **Re-run**.
- 1–4 sentences maximum.
- Plain English — expand acronyms on first use.
- Do **not** duplicate the `MessageTemplate` text; `Guidance` adds context, not repetition.

---

## Naming Conventions

| Item | Convention | Example |
|------|-----------|---------|
| Category CC | 2 uppercase letters, from `[Description]` on `EnumErrorCategory` | `HI`, `EF`, `BC` |
| Sub-code SSS | `(int)EnumErrorBehavior * 100 + (int)EnumErrorTarget`, zero-padded to 3 digits | `141`, `155` |
| Error number NNN | 3-digit zero-padded, sequential within the category+behavior+target combo | `001`, `042` |
| C# field name | Level prefix + original name: `E_Foo`, `W_Foo` | `E_DupBitName`, `W_HTOLGb` |
| `Guidance` | Imperative sentences, plain English | "Check …" |

There is no `name` or `category` string field on `ErrorCode` anymore — the field name itself (`E_DupBitName`) and the reflected static class name serve that purpose.

---

## MessageTemplate Rules — Bare `{0}` Anti-pattern

A template must include **descriptive context text** around the placeholder. A bare `"{0}"` alone
tells the reader nothing about what the value represents.

| Rule | Wrong ❌ | Right ✓ |
|------|---------|---------|
| Template must describe context | `template: "{0}"` | `template: "{0} doesn't exist in pattern folder."` |
| Template must describe context | `template: "{0}"` | `template: "Pattern version mismatch: {0}"` |
| Multi-arg templates must name each slot | `template: "{0} {1}"` | `template: "MeasC pin:{0} is different from pat info cap pin:{1}"` |

---

## Anti-patterns

| Anti-pattern | Problem | Fix |
|---|---|---|
| `template: "{0}"` alone | No context; message is unreadable without the code name | Add descriptive text: `"{0} <what happened>."` |
| `template: $"{myVar} not found"` | `myVar` not in scope at class init | Use `"{0} not found"` and pass `myVar` to `FormatMessage` |
| `template: someObj.SomeProperty` | Object not in scope at class init | Use `"{0} <context>."` and pass the property value to `FormatMessage` |
| `enumErrorBehavior` not in `EnumErrorBehavior` or `enumErrorTarget` not in `EnumErrorTarget` | Compile error | Add a new member to `EnumErrorBehavior` or `EnumErrorTarget` |
| `enumErrorCategory` not in `EnumErrorCategory` | Compile error | Add a new member to `EnumErrorCategory` with `[Description]` |
| Same `NNN` reused within a `(category, behavior, target)` combo | Code collision | Assign next available number |
| Same `FullCode` produced by two different fields, even across different static classes | `ErrorCodesFullCodeTests.AllErrorCodeFields_AcrossAllClasses_HaveUniqueFullCodes` fails the build | Before adding/editing a code, search the codebase for the target `ECCSSSNNN` string (or run the test) and pick an unused `NNN` |
| `CC` not in the Category Codes table | Undiscoverable | Add a `[Description("XX")]` entry to `EnumErrorCategory` |
| Passing `errorLevel` to `AddError` | Duplicates info already in `IErrorCode` | Let `AddError` read `errorCode.ErrorLevel` |
| Field name missing level prefix | Not self-documenting | Prefix with `E_`, `W_`, `I_`, or `C_` |

---

## Behavior Codes (`EnumErrorBehavior`)

The first digit (or two) of `SSS` is `(int)EnumErrorBehavior`. Defined in `EnumErrorBehavior.cs`:

| Value | Member | Meaning |
|-------|--------|---------|
| 1 | `Duplicate` | The same entity is defined more than once. |
| 2 | `Invalid` | Data exists but does not meet expected rules or constraints. |
| 3 | `Mismatch` | Conflicting definitions across related data. |
| 4 | `Missing` | Required data is absent or undefined. |
| 5 | `Redundant` | Unnecessary or duplicated data with no practical value. |
| 6 | `RuleViolation` | Data breaks system rules or logical constraints. |
| 7 | `Info` | Informational, non-error condition. |

To add a new behavior, append the next sequential value (8, 9, …) to `EnumErrorBehavior`.

---

## Target Codes (`EnumErrorTarget`)

`SSS - (int)EnumErrorBehavior * 100` is the 0-based ordinal of `EnumErrorTarget` — a large, alphabetically-ordered flat enum (`AcCategory` = 0 through `XYCoordinate`) describing *what* the error is about (`Pin`, `Pattern`, `TimeSet`, `Bit`, …). See `EnumErrorTarget.cs` for the full, current member list and ordinals — don't hardcode ordinals from memory; a new member inserted alphabetically shifts every ordinal after it.

The file currently lists members roughly alphabetically, but `SSS` is derived from ordinal position, not name. **Always append a new target at the end of the enum** — inserting it mid-list to "keep things alphabetical" silently reassigns the ordinal (and therefore the `FullCode`) of every member after it.

To compute `SSS` for a given `(behavior, target)` pair: `(int)EnumErrorBehavior * 100 + (int)EnumErrorTarget`, zero-padded to 3 digits. E.g. `Duplicate` (1) + `Pin` (55, i.e. the 56th member, 0-based index 55) → `155`.
