# DroidXLS

**Lightweight Android Excel library** — read and write .xlsx files with a Kotlin-native API.

A low-cost alternative to Aspose.Cells ($1,175+/year) at **$99/year** for commercial use. Free for personal, OSS, NPO, and educational use.

## Features

### Phase 1 (Complete)
- Read/write .xlsx (Excel 2007+)
- Cell values: text, numbers, booleans, dates, formulas, errors
- Styles: fonts, fills, borders, alignment, number formats (Kotlin DSL)
- Sheet operations: add, remove, copy, move, hide, rename
- Row/column: insert, delete, hide, width/height, freeze panes
- Merged cells
- Formula engine with ~30 functions (SUM, VLOOKUP, IF, CONCATENATE, etc.)
- Password-protected files (AES encryption)
- Async API (Kotlin Coroutines `suspend fun`)
- SAX streaming for memory efficiency

### Phase 2 (Complete)
- Image embedding (PNG, JPEG, WebP)
- Chart model (bar, line, pie, area, scatter)
- Auto filters, data validation, conditional formatting
- Named ranges
- Sheet protection
- CSV export
- HTML export

## Quick Start

```kotlin
// Create a workbook
val workbook = Workbook()
val sheet = workbook.addSheet("Sales")

// Write data
sheet["A1"].value = "Product"
sheet["B1"].value = "Revenue"
sheet["A2"].value = "Widget"
sheet["B2"].value = 12345.67

// Apply styles
sheet["A1"].style {
    font { bold = true; size = 14.0 }
    fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.LIGHT_BLUE }
    border { all = BorderStyle.THIN }
}

// Add formula
sheet["B3"].formula = "=SUM(B2:B2)"

// Save
workbook.save(outputStream)

// Save with password
workbook.save(outputStream, "secret123")
```

### Reading files

```kotlin
val workbook = Workbook.open(inputStream)
val sheet = workbook.sheets[0]
val name = sheet["A1"].stringValue
val revenue = sheet["B2"].numericValue

// Password-protected files
val wb = Workbook.open(inputStream, "password")
```

### Async API

```kotlin
val workbook = Workbook.openAsync(inputStream)
workbook.saveAsync(outputStream)
```

### CSV / HTML Export

```kotlin
val csv = CsvConverter.convertToString(sheet)
val html = HtmlConverter.convert(sheet, title = "Report")
```

## Installation

### Gradle (JitPack)

```kotlin
// settings.gradle.kts
dependencyResolutionManagement {
    repositories {
        maven { url = uri("https://jitpack.io") }
    }
}

// build.gradle.kts
dependencies {
    implementation("com.github.user:droidxls:0.1.0")
}
```

## Requirements

- Android API 26+ (Android 8.0)
- Kotlin

## License

[BSL 1.1](LICENSE) — Free for personal, non-commercial, OSS, NPO, and educational use.
Commercial use requires a license key ($99/year).
Converts to MIT 3 years after each release.

### Commercial Licensing

| Plan | Price | For |
|---|---|---|
| Personal | Free | Individuals, OSS, NPO, education |
| Indie | $99/year | Solo developers with commercial apps |
| Startup | $399/year | Small teams |
| Enterprise | Contact us | Large organizations |

Purchase at: *(Gumroad link to be added)*

## Library Size

- Release AAR: ~235 KB
- External dependencies: nimbus-jose-jwt (license validation only)
