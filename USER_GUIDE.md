# Process Flow Builder - User Guide

## For Non-Technical Users

This guide explains how to use the Process Flow Builder to create PowerPoint presentations from simple text files.

## What You Need

1. **ProcessFlowBuilder** (the executable file)
2. A text file describing your process (`.txt` file)
3. Optional: `my_power_logo.png` (your logo image)
4. Optional: `key.txt` (for a Key box on the flowchart)

## Quick Start

### Step 1: Prepare Your Files

1. Place the **ProcessFlowBuilder** executable in a folder
2. Place your process description file (e.g., `my_process.txt`) in the same folder
3. Optional: Place your logo file `my_power_logo.png` in the same folder
4. Optional: If you want a Key box, create `key.txt` with your key information

### Step 2: Run the Builder

**On Windows:**
- Double-click `ProcessFlowBuilder.exe`
- Or, open Command Prompt, navigate to the folder, and type: `ProcessFlowBuilder.exe`

**On Mac:**
- Open Terminal, navigate to the folder, and type: `./ProcessFlowBuilder`
- If you get a permission error, type: `chmod +x ProcessFlowBuilder` first, then try again

**On Linux:**
- Open Terminal, navigate to the folder, and type: `./ProcessFlowBuilder`
- If you get a permission error, type: `chmod +x ProcessFlowBuilder` first, then try again

### Step 3: Find Your Output

The program will create a PowerPoint file with the same name as your text file. For example:
- If your text file is `my_process.txt`
- Your PowerPoint file will be `my_process.pptx`

## Creating Your Process Description File

Your `.txt` file should describe your process using this format:

```
Start: [S1] START
 Leads to: [A1]

Action: [A1]
Title: Action Title
Details: First detail line
Details: • Second detail with bullet
Details: Third detail line
Leads to: [D1]

Decision: [D1] Should we proceed?
Path "Yes" -> [A2]
Path "No" -> [E1]

Action: [A2]
Title: Continue Process
Details: What happens when yes
Leads to: [E1]

End: [E1] END
```

### Key Elements:

- **Start:** `Start: [ID] Description` - Marks where the process begins
- **Action:** `Action: [ID]` with optional `Title:` and `Details:` lines
- **Information:** `Information: [ID]` or `Info: [ID]` - Information boxes
- **Decision:** `Decision: [ID] Question?` followed by `Path "Label" -> [NextID]` lines
- **End:** `End: [ID] Description` - Marks where the process ends

### Optional Features:

Add this line anywhere in your file to show a Key box:
```
ShowKey: True
```

This will display a Key box on the flowchart slide (if `key.txt` exists).

## Example: Complete Process File

Save this as `example_process.txt`:

```
Start: [S1] START
 Leads to: [A1]

Action: [A1]
Title: Validate Order
Details: • Check inventory
Details: • Verify payment
Details: • Confirm address
 Leads to: [D1]

Decision: [D1] Is order valid?
Path "Yes" -> [A2]
Path "No" -> [E1]

Action: [A2]
Title: Process Order
Details: • Create shipment
Details: • Send confirmation
 Leads to: [E1]

End: [E1] END

ShowKey: True
```

## Understanding the Output

The PowerPoint file will contain 4 slides:

1. **Cover Slide** - Title page with version table and distribution list
2. **Amendments Slide** - Change tracking table
3. **Flowchart Slide** - Visual diagram of your process
4. **Notes Slide** - Space for additional notes and guidance

## Troubleshooting

### "No .txt files found"
- Make sure you have a `.txt` file in the same folder as the executable
- The file should not be named `key.txt` (that's reserved for the Key box)

### "Logo not found"
- This is just a warning; the presentation will still be created
- Make sure `my_power_logo.png` is in the same folder if you want the logo

### Key box not showing
- Make sure your schema file contains `ShowKey: True`
- Make sure `key.txt` exists in the same folder
- Both conditions must be met

### Permission errors (Mac/Linux)
- In Terminal, type: `chmod +x ProcessFlowBuilder`
- Then try running it again

## Tips

1. **Multiple processes:** Place multiple `.txt` files in the folder. The program will use the newest one by default.

2. **Naming:** Use descriptive names for your text files (e.g., `inverter_order_process.txt`). The output PowerPoint will have the same name.

3. **Key box:** The Key box automatically sizes to fit your content. Keep `key.txt` concise for best results.

4. **Logo:** Supported formats include PNG, JPG, and other common image formats. The logo appears in the top-right of all slides.

## Need Help?

If you encounter issues:
1. Check that all files are in the same folder
2. Verify your text file follows the format described above
3. Make sure you're using the correct executable for your operating system (Windows/Mac/Linux)

