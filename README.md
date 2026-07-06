# Excel Schedule Importer for Revit

Recreates an Excel equipment schedule inside Revit as a **drafting view made of
native detail lines, text notes and filled regions** — fully vector, crisp at any
zoom and print scale, no raster images bloating the model.

Compatible with **Revit 2024, 2025 and 2026**.

## Installation

For colleagues installing this on their own PC — no admin rights needed, nothing
to build, nothing to type.

1. Download **[Install-ExcelScheduleImporter.bat](https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest/download/Install-ExcelScheduleImporter.bat)**
   (also attached to every [release](https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest)).
2. Double-click it. A console window shows progress, then closes.

It detects which Revit versions (2024/2025/2026) you have installed,
downloads the matching add-in from the latest release, and registers it under
your user profile.

Then start Revit — the **Import Excel Schedule** button is on the **Add-Ins**
tab. The first time, Revit may warn about an unsigned add-in — click
**Always Load**.

<details>
<summary>Prefer PowerShell directly?</summary>

```powershell
irm https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest/download/install.ps1 -OutFile "$env:TEMP\esi-install.ps1"; & "$env:TEMP\esi-install.ps1"
```

This is exactly what the `.bat` file runs under the hood.
</details>

**That's it — you only run this once.** The add-in checks for updates in the
background each time Revit launches and installs new versions automatically
(the swap happens after Revit closes, so it never interrupts your session).
See [DEPLOYMENT.md](DEPLOYMENT.md) for how updates are published and how the
auto-update mechanism works.

## What it does

1. Pick an `.xlsx` / `.xlsm` file (Excel does not need to be installed).
2. Pick a worksheet — the used portion of the sheet (content + formatting) is
   **auto-detected** and shown as an editable A1-style range.
3. On *Import*, a new drafting view is created containing:
   - **Detail lines** for every Excel border (thin / medium / thick mapped to
     Revit's Thin / Medium / Wide line styles, collinear segments merged into
     single lines to keep the element count low)
   - **Text notes** matching font, size, bold, italic, color, horizontal and
     vertical alignment, wrapping, and Excel's *displayed* value (number
     formats, percentages, dates are respected)
   - **Filled regions** for cell shading (optional)
   - **Merged cells** are handled correctly (one text note, no interior borders)
   - Hidden rows/columns collapse to zero, exactly like Excel printing
4. Drag the drafting view onto any sheet.

The geometry is drawn at *printed size × view scale*, so the schedule appears
on the sheet at the exact size Excel would have printed it — regardless of the
view scale you choose.

## Keeping schedules up to date

Every import stamps the drafting view with its source link (file, worksheet,
range, options) via extensible storage — the link is saved in the model.

`Add-Ins → Excel Import → Update Schedules` opens the manager:

- lists every imported schedule, which **sheet(s)** it is placed on, its source
  file/worksheet/range and import date
- **red rows = out of date** (the Excel file changed since import); they come
  pre-checked
- *Update Checked* / *Update All* rebuild the views **in place** — views keep
  their identity, so viewports on sheets are untouched
- imports that used the auto-detected range are **re-detected** on update, so
  rows added in Excel are included automatically

## Using it

`Add-Ins tab → Excel Import panel → Import Excel Schedule`

Options in the dialog:

| Option | Meaning |
|---|---|
| Cell range | Auto-detected; edit to import a smaller portion (e.g. `A1:F42`) |
| View name | Name of the new drafting view |
| View scale | Cosmetic only — printed size is identical at any scale |
| Text size % | Global multiplier if you want text slightly smaller/larger than Excel |
| Match body text to X mm | OFF by default (1:1 with Excel is most faithful). When enabled, scales the WHOLE schedule uniformly so the body font prints at exactly X mm; titles keep their relative proportion. |
| Override font | OFF by default (keeps Excel fonts). When enabled, forces every text note to one font (e.g. Arial). |

After creation, a **fit pass** measures every text note's real rendered size and
repairs overflow: slightly-too-wide single-line text is shrunk a touch (like
Excel's shrink-to-fit, at most 30%); anything longer wraps at the spill
boundary; text taller than its row is shrunk to fit (floor 1.0 mm).
| Recreate cell fills | Creates `XLS Fill RRGGBB` filled region types |
| Draw thin gridlines | Adds gridlines where Excel had no explicit borders |

Text styles are created on demand as `XLS <font> <size>mm [B][I]` text note
types, and re-used on subsequent imports.

## Building

Requires the .NET 8 SDK (a per-user copy lives in `%USERPROFILE%\.dotnet` if it
was installed by the setup script).

```powershell
.\build.ps1        # builds R24 + R25 + R26 and deploys to the Revit Addins folders
```

or a single version:

```powershell
dotnet build src\ExcelScheduleImporter -c "Release R26"
```

Each build auto-copies the DLLs and the `.addin` manifest to
`%AppData%\Autodesk\Revit\Addins\<year>\`. Restart Revit to load it.

- Revit 2024 → .NET Framework 4.8
- Revit 2025 / 2026 → .NET 8

Revit API references come from the `Nice3point.Revit.Api.*` NuGet packages
(compile-time only); Excel parsing uses [ClosedXML](https://github.com/ClosedXML/ClosedXML) (MIT).

## Project layout

```
src/ExcelScheduleImporter/
├── App.cs                       IExternalApplication — ribbon button
├── ImportExcelCommand.cs        IExternalCommand — orchestrates the import
├── Excel/
│   ├── ExcelReader.cs           ClosedXML: range detection, formatting extraction
│   └── TableModel.cs            Intermediate model (paper-space mm)
├── Revit/
│   └── DraftingViewBuilder.cs   Creates view, lines, text notes, filled regions
└── UI/
    └── ImportForm.cs            WinForms dialog
addin/ExcelScheduleImporter.addin
build.ps1
```

## Known limitations

- Legacy `.xls` (BIFF) files are not supported — save as `.xlsx` first.
- Border **colors** are not reproduced (Revit line color comes from the line
  style); border weights are.
- Diagonal borders, rotated text, in-cell rich text (mixed formatting within
  one cell), and embedded images are not reproduced.
- Column widths are estimated from Excel's character-width model assuming the
  default Calibri 11 base font — widths may differ by a percent or two for
  workbooks with unusual default fonts.
- If another add-in ships a conflicting `DocumentFormat.OpenXml` version, Excel
  reading may fail; the fix is aligning package versions.
