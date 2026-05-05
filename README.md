# Deriva

Excel-DNA add-in for financial calculations and derivatives pricing utilities.

## Calendar Functions

| Function | Description |
|---|---|
| `IsDU(date)` | Returns TRUE if date is a business day. |
| `IsHoliday(date)` | Returns TRUE if date is an ANBIMA national holiday (weekends are not holidays). |
| `DU(date1, date2)` | Counts positive business days between two dates (exclusive start, inclusive end). |
| `ProxDU(initial_date, number_working_days)` | Adds or subtracts business days. N=0 adjusts to next business day if date is not already one. |
| `AdjustDU(date, convention)` | Adjusts date to the nearest business day per convention. |
| `ProxMonths(date, months, convention)` | Adds calendar months then adjusts to business day per convention. |
| `Holidays(start_date, end_date)` | Returns a vertical spill array of all ANBIMA national holidays in the period. |

## Interpolation Functions

| Function | Description |
|---|---|
| `InterpolateValues(x_values, y_values, x, [method])` | Interpolates the value of `y` for a target `x`. Uses natural cubic spline by default; pass `"linear"` to use piecewise linear interpolation. |

`x_values` can contain numbers or Excel dates. `y_values` must contain numeric values with the same number of points as `x_values`. The function sorts points by `x`, rejects duplicate `x` values, and does not extrapolate outside the known range.

Examples:

```excel
=InterpolateValues(A2:A20,B2:B20,F2)
=InterpolateValues(A2:A20,B2:B20,F2,"linear")
```

Accepted methods are `cubic`, `cubic_spline`, `spline`, and `linear`. Leaving the method blank is the same as `cubic_spline`.

## ETTJ Functions

| Function | Description |
|---|---|
| `GetETTJ(date, [curve], [cache])` | Fetches B3 TaxaSwap ETTJ data for one date and spills `refdate`, `curva`, `descricao`, `dias_corridos`, `dias_uteis`, `taxa`, and `vertice`. |
| `GetETTJHistorico(start_date, end_date, [curve], [cache], [ignore_errors])` | Fetches the same long table for an interval, skipping weekends and optionally skipping no-data/error dates. |
| `GetCurve(curve, reference_date, end_date)` | Fetches one ETTJ curve and returns the interpolated rate for `DU(reference_date, end_date)` using the default cubic spline interpolation. `end_date` may be a scalar or an Excel range/spill. |

`curve` defaults to `PRE` and accepts a single code, comma-separated codes such as `"PRE,DIC"`, an Excel range, or `"TODOS"`. Rates are returned in decimal form, so `0.1465` means 14.65% a.a.

Examples:

```excel
=GetETTJ("09/04/2026","PRE,DIC")
=GetETTJHistorico("01/04/2026","09/04/2026","PRE")
=GetCurve("PRE",ProxDU(TODAY(),-1),ProxDU(ProxDU(TODAY(),-1),252))
=GetCurve("PRE",$A$1,B2:B5000)
```

For large sheets, prefer one `GetCurve` call over an `end_date` range instead of thousands of scalar formulas. ETTJ data is fetched from B3 TaxaSwap files at `https://www.b3.com.br/pesquisapregao/download?filelist=TS{YYMMDD}.ex_,` and cached by date under `%LOCALAPPDATA%\Deriva\ETTJ\Cache\TS` by default.

## Ribbon

The add-in adds a `Deriva` ribbon tab with:

| Button | Description |
|---|---|
| `Dashboard` | Opens a dialog showing the last Holiday and ETTJ update time, status, and detail. |
| `Settings` | Opens a dialog for the holiday URL, ETTJ init curves, ETTJ cache directory, and B3 source metadata. |

## Business Day Conventions

| Code | Meaning |
|---|---|
| `F`  | Following — move to the next business day if date is not a business day. |
| `MF` | Modified Following — Following, but if that crosses into a new month, use Preceding instead. |
| `P`  | Preceding — move to the previous business day if date is not a business day. |
| `MP` | Modified Preceding — Preceding, but if that crosses into a previous month, use Following instead. |

Convention codes are case-insensitive. Dates that are already business days are returned unchanged under any convention.

## DU Counting Convention

`DU(date1, date2)` counts business days in the half-open interval `(min, max]` — the start date is **excluded**, the end date is **included**. Order does not matter and the result is always positive. `DU(D, D) = 0`.

This matches the standard Brazilian fixed-income convention (ANBIMA/Cetip basis 252).

## Holiday Data

Holiday data is fetched automatically from the [ANBIMA national holiday list](https://www.anbima.com.br/feriados/arqs/feriados_nacionais.xls) each time Excel starts, and is cached locally at `%APPDATA%\Deriva\holidays.json` for up to 24 hours.

### Degraded Mode

| Situation | Behaviour |
|---|---|
| ANBIMA unreachable, cache fresh (< 24 h) | Functions work normally using the cached list |
| ANBIMA unreachable, cache stale (≥ 24 h) | Functions work using the stale cache; a warning appears in the Excel status bar for ~8 seconds |
| ANBIMA unreachable, no cache | Functions return `#N/A`; an error is shown in the Deriva log window. Reconnect and restart Excel to restore full functionality |

## Installation

1. Build the project in Visual Studio (**Release** configuration). Excel-DNA produces a self-contained `bin\Release\Deriva.Excel-packed.xll`.
2. Copy `Deriva.Excel-packed.xll` to any permanent local folder.
3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse**, select the `.xll` file, and click **OK**.
4. The add-in loads automatically on every subsequent Excel startup.

## Requirements

- Microsoft Excel 2010 or later (32-bit or 64-bit); Excel 2019 / Microsoft 365 required for `Holidays()` spill arrays
- .NET Framework 4.8 (pre-installed on Windows 10 version 1903 and later, and on all Windows 11 editions)
- Internet access on first run (or an existing cache from a previous session)
