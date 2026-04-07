# Check Duct Orientation Tool

## Overview
This pyRevit tool detects fabrication duct items where the Width × Height dimension labels don't match the physical orientation in plan view. This is critical for fabrication because incorrect dimension labels can lead to manufacturing errors.

## The Problem
Rectangular duct should be labeled consistently:
- **In Plan View**: Looking down, the first dimension should be the depth (front-to-back), the second should be the width (left-to-right)
- **In Elevation**: Looking from the side, dimensions show width × height

Sometimes fabrication items get created with swapped dimensions. For example:
- A duct that's **20" tall × 40" wide** should display as:
  - **40×20** in elevation (width × height)
  - But the Width and Height **labels** should match the physical orientation

When labels are swapped, it creates confusion and potential fabrication errors.

## How It Works

### Detection Logic
1. **Collects Fabrication Parts**: From selection (if any) or entire model
2. **Filters Rectangular Ducts**: Only checks rectangular cross-sections
3. **Analyzes Orientation**: For each duct:
   - Gets labeled Width and Height dimensions
   - Determines duct's axis direction (which way it runs)
   - Gets the transform/rotation to determine physical orientation
   - Checks if Width aligns horizontally and Height aligns vertically
4. **Identifies Mismatches**: Flags ducts where labels are swapped

### Technical Approach
The tool uses the element's Transform property to get the local coordinate system:
- **BasisX, BasisY, BasisZ**: The three axes of the duct's local coordinate system
- **Vertical Alignment**: Determines which basis vector points up/down (Z direction)
- **Horizontal Alignment**: Confirms width is horizontal, not vertical

For horizontal ducts:
- Height vector should align with world Z-axis (vertical)
- Width vector should be horizontal (minimal Z component)
- If these are swapped, dimensions are incorrect

## Usage

### Running the Tool
1. **Optional**: Select specific fabrication duct to check
2. Click **"Check Duct Orientation"** button
3. Review results in output window
4. Mismatched elements will be automatically selected

### Output
The tool provides:
- **Summary Statistics**: 
  - Total parts analyzed
  - Number of rectangular ducts
  - Correct vs. mismatched orientations
- **Detailed Report**: Lists each mismatched element with:
  - Element ID
  - Current labeled dimensions (Width × Height)
  - What dimensions should be (corrected)
  - Alignment scores

### Results Interpretation
- **CORRECT**: Dimensions match physical orientation
- **MISMATCHED**: Width and Height labels are swapped
  - Example: Labeled as 20×40 but should be 40×20
- **SKIPPED**: Non-rectangular, vertical, or couldn't determine orientation

## Fixing Mismatches
If the tool finds mismatched ducts:

1. **Review Selected Elements**: The tool auto-selects all mismatched ducts
2. **Verify in Plan View**: Look at the duct orientation
3. **Correction Options**:
   - **Option A - Rotate Elements**: If possible, rotate the duct 90° to match labels
   - **Option B - Update Parameters**: Manually swap Width/Height parameter values
   - **Option C - Recreate**: Delete and recreate with correct orientation

## Limitations
- **Horizontal Ducts Only**: Currently focuses on horizontal runs (axis Z < 0.5)
- **Vertical/Angled Ducts**: Skipped for now (different orientation logic needed)
- **Requires Valid Geometry**: Elements must have:
  - CenterLine or ConnectorManager
  - Transform with valid basis vectors
  - NominalWidth and NominalHeight properties or parameters

## Technical Details

### Parameters Checked
The tool tries multiple parameter names in order:
- **Width**: NominalWidth → Width → Dim A
- **Height**: NominalHeight → Height → Dim B

### Geometry Sources
- **Axis Direction**: From CenterLine endpoints or Connector positions
- **Rotation**: From element Transform.BasisX/Y/Z vectors
- **Alignment**: Dot product with world Z-axis (0,0,1)

### Tolerance
- **Angle Tolerance**: 0.1 radians (~5.7°)
- **Distance Tolerance**: 0.01 feet

## File Structure
```
09Check Duct Orientation.pushbutton/
├── script.py          # Main tool logic
├── bundle.yaml        # pyRevit configuration
├── icon.png           # Large icon
├── icon-small.png     # Small icon
└── README.md          # This file
```

## Future Enhancements
Potential improvements:
1. **Vertical Duct Support**: Add logic for vertical runs
2. **Angled Duct Support**: Handle ducts at various angles
3. **Auto-Fix Option**: Automatically swap parameters where possible
4. **Export Report**: CSV export of findings
5. **Batch Correction**: Select all and apply consistent orientation rules

## Version History
- **v1.0.0** (2024): Initial release
  - Basic horizontal duct orientation checking
  - Selection of mismatched elements
  - Detailed reporting

## Author
Jeremiah Griffith - Southland Industries
Part of the DetailersTools extension

## Support
For issues or enhancement requests, contact the development team or submit feedback through pyRevit.
