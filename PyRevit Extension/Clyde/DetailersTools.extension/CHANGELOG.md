# Changelog

All notable changes to the Detailers Tools for pyRevit extension will be documented here.  
The format follows [Keep a Changelog](https://keepachangelog.com/) and adheres to [Semantic Versioning](https://semver.org/).

---

## [0.0.1] - 2025-08-10
### Added
- **Compare Elements** tool:
  - Compare key properties of two selected elements.
  - Pulls full MEP Fabrication ITM parameter data from the FabricationPart API.
  - Accurate connector reporting using Fabrication Connector Info.
  - Scrollable and resizable report dialog.
  - Option to export selected comparison parameters to CSV file.
  - Integrated icons for improved button recognition in pyRevit.

- **Compare Views** tool:
  - Compare categories, filters, worksets, links, and annotations between two Revit views.
  - Scrollable results window.

- **UI/UX Enhancements**:
  - Added icons to both Compare Elements and Compare Views buttons.
  - Compare tools grouped under a single `Compare` stack in the `Detailer QA` panel of the pyRevit tab.
  - Manual close for report dialogs to avoid unwanted auto-close behavior.

### Notes
- Requires **pyRevit 5.1+**.
- Tested in Revit versions 2022 and 2024.
- First packaged release with MSI installer.
