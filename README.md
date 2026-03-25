# DroidXLS — Android Excel Library for Kotlin

[![License: BSL 1.1](https://img.shields.io/badge/License-BSL_1.1-blue.svg)](LICENSE)
[![Android API 26+](https://img.shields.io/badge/API-26%2B-brightgreen.svg)](https://developer.android.com/about/versions/oreo)
[![Kotlin](https://img.shields.io/badge/Kotlin-2.0-purple.svg)](https://kotlinlang.org)
[![](https://jitpack.io/v/youichi-uda/droidxls.svg)](https://jitpack.io/#youichi-uda/droidxls)

**Read and write Excel .xlsx files natively on Android.** Kotlin-first API, SAX streaming for low memory usage, and a ~235KB AAR — built specifically for Android, not ported from Java desktop libraries.

> **$99/year** for commercial use. Free for personal, open source, NPO, and education.
> An affordable alternative to Aspose.Cells for Android ($1,175+/year).

## Why DroidXLS?

| | DroidXLS | Aspose.Cells Android | Apache POI (Android port) |
|---|---|---|---|
| **Price** | $99/year | $1,175+/year | Free |
| **Size** | ~235 KB | ~30 MB | ~50 MB |
| **API** | Kotlin DSL | Java | Java |
| **Memory** | SAX streaming | DOM | DOM |
| **Android support** | Native | Yes | Unofficial |
| **Maintained** | Yes | Yes | No official Android |

## Features

### Core
- Read/write **.xlsx** (Excel 2007+ / OOXML)
- Cell values: text, numbers, booleans, dates, formulas, errors
- **Kotlin DSL** for styles: `sheet["A1"].style { font { bold = true } }`
- Fonts, fills, borders, alignment, number formats
- Sheet operations: add, remove, copy, move, hide, rename
- Row/column: insert, delete, hide, width/height, freeze panes
- Merged cells

### Formula Engine
~30 built-in functions:
- **Aggregate:** SUM, AVERAGE, COUNT, COUNTA, COUNTBLANK, MAX, MIN
- **Logic:** IF, AND, OR, NOT, IFERROR, IFNA
- **Lookup:** VLOOKUP, HLOOKUP, INDEX, MATCH
- **String:** CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, SUBSTITUTE
- **Math:** ROUND, ABS, MOD, INT, CEILING, FLOOR
- **Date:** TODAY, NOW, DATE, YEAR, MONTH, DAY

### Advanced
- **Images:** embed PNG, JPEG, WebP
- **Charts:** bar, column, line, pie, area, scatter
- **Auto filters** and **data validation** (dropdown lists, value constraints)
- **Conditional formatting**
- **Named ranges** and **sheet protection**
- **Password protection** (AES-256 encryption)
- **CSV export** and **HTML export**
- **Coroutines:** `suspend fun` async API

## Quick Start

### Installation (Gradle + JitPack)

```kotlin
// settings.gradle.kts
dependencyResolutionManagement {
    repositories {
        maven { url = uri("https://jitpack.io") }
    }
}

// build.gradle.kts
dependencies {
    implementation("com.github.y1uda:droidxls:0.1.0")
}
```

### Create and Write

```kotlin
val workbook = Workbook()
val sheet = workbook.addSheet("Sales")

// Write data
sheet["A1"].value = "Product"
sheet["B1"].value = "Revenue"
sheet["A2"].value = "Widget"
sheet["B2"].value = 12345.67

// Style with Kotlin DSL
sheet["A1"].style {
    font { bold = true; size = 14.0; color = OfficeColor.WHITE }
    fill { patternType = PatternType.SOLID; foregroundColor = OfficeColor.Rgb(47, 85, 151) }
    border { all = BorderStyle.THIN }
}

// Formulas
sheet["B3"].formula = "=SUM(B2:B2)"

// Save
workbook.save(outputStream)
```

### Read Existing Files

```kotlin
val workbook = Workbook.open(inputStream)
val sheet = workbook.sheets[0]
val name = sheet["A1"].stringValue      // "Product"
val revenue = sheet["B2"].numericValue  // 12345.67
```

### Password-Protected Files

```kotlin
// Save encrypted
workbook.save(outputStream, "myPassword")

// Open encrypted
val wb = Workbook.open(inputStream, "myPassword")
```

### Async API (Coroutines)

```kotlin
val workbook = Workbook.openAsync(inputStream)
workbook.saveAsync(outputStream)
```

### Export to CSV / HTML

```kotlin
val csv = CsvConverter.convertToString(sheet)
val html = HtmlConverter.convert(sheet, title = "Report")
```

## Sample App

The [`sample-app/`](sample-app/) module demonstrates all features with 8 interactive demos. Build and install:

```bash
./gradlew :sample-app:installDebug
```

## Requirements

- **Android API 26+** (Android 8.0 Oreo)
- **Kotlin** (Kotlin-only API, no Java interop)

## License

**[BSL 1.1](LICENSE)** — Business Source License

- **Free:** personal projects, open source, NPO, education
- **Commercial ($99/year):** revenue-generating apps, enterprise internal tools
- **Converts to MIT** 3 years after each release

### Purchase a Commercial License

**[Buy on Gumroad — $99/year](https://y1uda.gumroad.com/l/droidxls)**

```kotlin
// Add to Application.onCreate()
DroidXLS.initialize(context, licenseKey = "YOUR-LICENSE-KEY")
```

## Part of DroidOffice

DroidXLS is part of the **DroidOffice** library family:

| Library | Format | Status |
|---|---|---|
| **[DroidXLS](https://github.com/youichi-uda/droidxls)** | .xlsx (Excel) | Available |
| **DroidDoc** | .docx (Word) | Planned |
| **DroidSlide** | .pptx (PowerPoint) | Planned |

All libraries share [droidoffice-core](https://github.com/youichi-uda/droidoffice-core) for OOXML parsing, licensing, and common styles.

## Contributing

Issues and pull requests are welcome on [GitHub](https://github.com/youichi-uda/droidxls/issues).
