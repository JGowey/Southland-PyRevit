Detailers Tools for pyRevit
===========================

Version: 0.0.2
Author: Jeremiah Griffith
Date: 08/12/2025

Description
-----------
This pyRevit extension adds a suite of powerful comparison tools for Revit users:
1.  **Compare Elements**: Compare key properties of two or more selected elements, with special support for MEP Fabrication ITMs.
2.  **Compare Views**: Compare settings between any two views or view templates in the project.
3.  **Compare Schedules**: Compare the fields, filters, sorting, and appearance of two or more schedules.
4.  **Compare Families**: Compare the parameters and types of any two loadable families in the project.

Features
--------
-   Custom icons for all comparison tools.
-   Full integration into the pyRevit ribbon under the "Detailer QA" panel.
-   **Enhanced Comparison Workflow**:
    -   Select which specific aspects of an object you want to compare (e.g., just Filters and Worksets).
    -   Filter reports to show **everything**, only the **differences**, or only the **similarities**.
-   Select from all project views and families, not just active ones.
-   Scrollable, resizable comparison report dialogs.
-   Optional CSV export for all comparison reports.

Installation
------------
1.  Ensure pyRevit (version 5.1 or later) is installed and loaded in Revit.
2.  Place the `DetailerTools.extension` folder in your pyRevit `extensions` directory.
3.  Reload pyRevit or restart Revit.
4.  Access the tools from the "Detailer QA" panel on the pyRevit tab.

Uninstallation
--------------
-   Manually delete the `DetailersTools.extension` folder from your pyRevit `extensions` directory.

Support
-------
For bug reports, feedback, or feature requests, contact:
[Support Email or Internal Helpdesk URL]
