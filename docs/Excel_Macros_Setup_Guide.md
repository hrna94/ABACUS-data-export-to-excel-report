# Excel Report with VBA Analysis - User Guide

## Enabling Macros on First Open

When opening the `.xlsm` file, a **yellow security banner** will appear:

```
 SECURITY WARNING Macros have been disabled. [Enable Content]
```

**→ Click "Enable Content"** to activate analytical functions.

## Permanent Macro Settings (Optional)

For frequent use, you can adjust settings:

1. **File** → **Options** → **Trust Center**
2. **Trust Center Settings** → **Macro Settings**
3. Select: **"Disable all macros with notification"** 

## Running the Analysis Tool

After enabling macros:

1. Go to the **"Summary"** worksheet
2. Press **Alt + F11** → opens VBA editor
3. Run: **`SetupSummaryUI`** (if UI is not visible)

### Or simply:

1. In the "Summary" sheet, find the analysis buttons
2. Set filters (date, devices, intervals)
3. Click **"Run Analysis"**

## Analysis Tool Features

- ** Time Filters:** Set analysis period
- **️ Intervals:** Hourly, daily, weekly, monthly analysis
- ** Device Filters:** Analysis by specific entry/exit points
- ** Results:** Clear table with traffic counts
- ** Help:** Integrated documentation

## Troubleshooting

**Macros not working:**
- Check if you clicked "Enable Content"
- Restart Excel and try again

**UI not showing:**
- Press Alt + F11, run `SetupSummaryUI`
- Or use "Setup UI" button if available

**Analysis errors:**
- Check if data exists in the selected date range
- Verify date formats in settings
