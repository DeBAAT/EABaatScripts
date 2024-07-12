# EABaatScripts
A number of scripts for Sparx Enterprise Architect.

---

## File structure

BaatDiagramScripts
: A number of scripts which can be activated on a Diagram. Right-click in an empty space on the diagram and select the appropriate script from "*Specialize => Scripts*".

- **BaatPackageScripts**
A number of scripts which can be activated on a Package. Right-click on a package and select the appropriate script from "Specialize => Scripts".
- **BaatScriptLib**
A number of scripts which cannot be used directly but are to be used in other scripts.

| Folder | Description |
| ----------- | ----------- |
| + **BaatDiagramScripts** | A number of scripts which can be activated on a Diagram. Right-click in an empty space on the diagram and select the appropriate script from "*Specialize => Scripts*". |
| + **BaatPackageScripts** | A number of scripts which can be activated on a Package. Right-click on a package and select the appropriate script from "*Specialize => Scripts*". |
| ++ *ModelCountTaggedValuesVBS* | A VBScript to count the number of all Enum TaggedValues which are written to an Excel file including the date. Each time the script is run, the values for today are added or updated on the worksheet such that a history of values is created which can be processed by e.g. PowerBI. |
| + **BaatScriptLib** | A number of scripts which cannot be used directly but are to be used in other scripts. |
| ++ *JavaScript-Connector* | A JavaScript based on JavaScript-TaggedValue contains some methods to manipulate Connectors. |
| ++ *JavaScript-EXCEL* | A JavaScript based on JavaScript-CSV contains some methods to manipulate Excel files. |
| ++ *JavaScript-ImEx* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to support the ImEx functionality as provided in other scripts. |
| ++ *ImExExportConnectorsJAVA* | A JavaScript based on the [EA Excel import-export](https://bellekens.com/ea-excel-import-export/) tool from Geert Bellekens. It contains some methods to export Connector information to an Excel file. |

