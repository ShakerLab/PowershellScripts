# PowershellScripts
Powershell scripts for lab tasks

## Script Usage

### structureQuPathProject.ps1

This script takes a directory with `.jpg` images and their respective `.qpdata` files. Note that the name of the `.jpg` and `.qpdata` file must match, and ideally files should be named uniquely (even if in different folders) and not contain spaces. Ideally, create a QuPath project and import at least one image manually so the `.qpproj` file generates. Then, delete the first entry and set `.lastID` to `0`. Note this script depends on an installation of `jq`, which can be installed with Winget via `winget install -e --id stedolan.jq`. Note that by default, this installation can only be called with `jq-1.6`. You can simply rename the jq executable or modify the script. Once all requirements are met, invoke when the current directory is set to the root of the QuPath project, which is the directory containing the `.qpproj` file.

### groupOrganoidData.ps1

After exporting all measurments from QuPath:
- Open the QuPath project
- Select `Measure > Export measurements...`
- Use the `>>` button to select all images
- Select an output file
- Choose `Annotations` for the export type
- Choode `Comma (.csv)` for the sperator
- Select `Export`

Place the script in the same directory as `measurements.csv` and run. The script assumes that the image filenames are broken down by group, then day, then replicate seperated by underscores. This can be modified if not desirable. Three CSV files will be output: one seperated by group, day, and replicate (`groupedDataGDR.csv`); one seperated by group and day (`groupedDataGD.csv`); and one seperated by group (`groupedDataG.csv`).
