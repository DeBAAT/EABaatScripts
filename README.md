# EABaatScripts
A number of scripts for Sparx Enterprise Architect.

---

## Scripts in the browser structure

| Level | Folder / File | Description |
| ----------- | ----------- | ----------- |
| + | **BaatDiagramScripts** | A number of scripts which can be activated on a Diagram. Right-click in an empty space on the diagram and select the appropriate script from "*Specialize => Scripts*". |
| ++ | *DImExExportConnectors* *DImExExportElements* *DImExImportConnectors* *DImExImportElements* | A set of JavaScripts using the *ImExExport* scripts in *BaatScriptLib* to im/export Connector and Element information for the Elements present on the Diagram selected. |
| ++ | *DiagramAnalysisLayout* | Dynamically draw elements on the diagram as indicated by the TaggedValues defined in the selected element. |
| ++ | *ImportDataBricksAttributes* *ImportDataBricksConnectors* *ImportDataBricksElements* | A set of JavaScripts using the *ImportDataBricksCommon* script in *BaatScriptLib* to import Attributes, Connector and Element information for the DataBricks information present in the CSV files selected. |
| ++ | *ImportDataBricksDALGeneratie* | Generate a *DiagramAnalysisLayout* diagram for all DAL Elements present on the Diagram selected. |
| ++ | *RefreshDynamicLegend* | Refresh the set of color definitions of all LegendElements present on the Diagram selected. |
| + | **BaatPackageScripts** | A number of scripts which can be activated on a Package. Right-click on a package and select the appropriate script from "*Specialize => Scripts*". |
| ++ | *PImExExportConnectors* *PImExExportElements* *PImExImportConnectors* *PImExImportElements* | A set of JavaScripts using the *ImExExport* scripts in *BaatScriptLib* to im/export Connector and Element information for the Elements present in the Package selected. The scripts check recursively all sub-packages as well. |
| ++ | *ModelCountTaggedValues* | A VBScript to count the number of all Enum TaggedValues which are written to an Excel file including the date. Each time the script is run, the values for today are added or updated on the worksheet such that a history of values is created which can be processed by e.g. PowerBI. |
| ++ | *ModelCountDuplicateTaggedValues* | A JavaScript to count the number of duplicate TaggedValues found in the complete model repository. |
| ++ | *ModelDeleteDuplicateTaggedValues* | A JavaScript to delete the duplicate TaggedValues found in the complete model repository. |
| + | **BaatScriptLib** | A number of scripts which cannot be used directly but are to be used in other scripts. |
| ++ | *BaatScript-Database* | A JavaScript based on JavaScript-Database to try to optimise memory usage. |
| ++ | *BaatScript-Logging* | A JavaScript based on JavaScript-Logging added the name of the calling function to the updated log format. |
| ++ | *BaatScript-XML* | A JavaScript based on JavaScript-XML to try to optimise memory usage. |
| ++ | *ImExExportConnectors* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to export Connector information to an Excel file. |
| ++ | *ImExExportElements* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to export Element information to an Excel file. |
| ++ | *ImExImportConnectors* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to import Connector information from an Excel file. |
| ++ | *ImExImportElements* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to import Element information from an Excel file. |
| ++ | *JavaScript-Connector* | A JavaScript based on JavaScript-TaggedValue contains some methods to manipulate Connectors. |
| ++ | *JavaScript-EXCEL* | A JavaScript based on JavaScript-CSV contains some methods to manipulate Excel files. |
| ++ | *JavaScript-ImEx* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to support the ImEx functionality as provided in other scripts. |


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


## Tips on how to use these scripts

### ImExport Excel files

As with the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens, these scripts allow to im/export information on Connectors or Elements using an Excel sheet. The main difference is in the scope of objects exported: the *DimEx* scripts export only the information related to the Elements present on the diagram the script is running on. The *PimEx* scripts export only the information related to the Elements present in the package the script is running on, including any sub-packages below. Also, when exporting, all TaggedValues defined for the objects are exported, no need to define them in the export query.

### ImportDataBricks

Follow these steps to import information from a set of CSV files generated from DataBricks:
1. Use *ImportDataBricksElements* to import the elements as tables into the current package, a sub-package is created based on the value of the Stage TaggedValue.
1. Use *ImportDataBricksConnectors* to import the relationships between the tables imported earlier.
1. Use *ImportDataBricksAttributes* to import the Attributes for the tables imported earlier.
1. Use *ImportDataBricksDALGeneratie* to generate *DiagramAnalysisLayout* diagrams for each of the DAL Elements available on the selected diagram.

Some additional remarks to consider:
1. Each object related to DataBricks is identified by a unique IDBKey stored as TaggedValue, this counts for both Elements and Connectors.
1. In order for a DAL Element to be used to generate a *DiagramAnalysisLayout* diagram, the Keyword of the Element needs to contain the string "DiagramAnalysisLayout", it may be case-insensitive. Next to that, there are a number of TaggedValues defining the layout of the diagram, e.g. DALNumLevels to define the number of levels to analyse.
1. In order to limit the Elements shown on a *DiagramAnalysisLayout* diagram, each Connector is filtered using the TaggedValue DALFilter. Most values are defined as "DALLineage".

