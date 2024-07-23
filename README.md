# EABaatScripts
A number of scripts for Sparx Enterprise Architect.

---

## Scripts in the browser structure

| Level | Folder / File | Description |
| ----------- | ----------- | ----------- |
| + | **BaatDiagramScripts** | A number of scripts which can be activated on a Diagram. Right-click in an empty space on the diagram and select the appropriate script from "*Specialize => Scripts*". |
| ++ | *DImExExportConnectorsJAVA* *DImExExportElementsJAVA* *DImExImportConnectorsJAVA* *DImExImportElementsJAVA* | A set of JavaScripts using the *ImExExport* scripts in *BaatScriptLib* to im/export Connector and Element information for the Elements present on the Diagram selected. |
| + | **BaatPackageScripts** | A number of scripts which can be activated on a Package. Right-click on a package and select the appropriate script from "*Specialize => Scripts*". |
| ++ | *PImExExportConnectorsJAVA* *PImExExportElementsJAVA* *PImExImportConnectorsJAVA* *PImExImportElementsJAVA* | A set of JavaScripts using the *ImExExport* scripts in *BaatScriptLib* to im/export Connector and Element information for the Elements present in the Package selected. The scripts check recursively all sub-packages as well. |
| ++ | *ModelCountTaggedValuesVBS* | A VBScript to count the number of all Enum TaggedValues which are written to an Excel file including the date. Each time the script is run, the values for today are added or updated on the worksheet such that a history of values is created which can be processed by e.g. PowerBI. |
| ++ | *ModelCountDuplicateTaggedValues* | A JavaScript to count the number of duplicate TaggedValues found in the complete model repository. |
| ++ | *ModelDeleteDuplicateTaggedValues* | A JavaScript to delete the duplicate TaggedValues found in the complete model repository. |
| + | **BaatScriptLib** | A number of scripts which cannot be used directly but are to be used in other scripts. |
| ++ | *JavaScript-Connector* | A JavaScript based on JavaScript-TaggedValue contains some methods to manipulate Connectors. |
| ++ | *JavaScript-EXCEL* | A JavaScript based on JavaScript-CSV contains some methods to manipulate Excel files. |
| ++ | *JavaScript-ImEx* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to support the ImEx functionality as provided in other scripts. |
| ++ | *ImExExportConnectorsJAVA* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to export Connector information to an Excel file. |
| ++ | *ImExExportElementsJAVA* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to export Element information to an Excel file. |
| ++ | *ImExImportConnectorsJAVA* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to import Connector information from an Excel file. |
| ++ | *ImExImportElementsJAVA* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to import Element information from an Excel file. |


## Remarks

These scripts are good enough for my needs. Feel free to reuse it, fork it, modify it, etc. but please mention it when you encounter any quircks or bugs or other discrepancies.


## How to install

There are several ways to install these scripts:
1. Create empty script and copy/paste file contents
1. Import xml file
Check the EA User Guide for more information on [Scripting](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/the_scripter_window.html).

### Manual installation

Follow these steps to install a package script from a source file:
1. Create new Project Browser group "BaatPackageScripts"
1. Create new JavaScript "ModelCountDuplicateTaggedValues"
1. Open newly created script
1. Open file "ModelCountDuplicateTaggedValues.js"
1. Replace all existing code in the open script with the code in the JS file
Repeat these steps for all files, check the group to put them in.
The file "ImExScripts.xml" contains all scripts related to the Im/Export of Elements and Connectors, including the necessary lib scripts.

### Import XML file

Follow these steps to import a script from an XML file:
1. Use menu "Settings => Transfer => Import Reference Data"
1. Select the XML file to import, e.g. "ModelCountDuplicateTaggedValues.xml"
1. Select "Automation Scripts" and import the contents of the XML file

