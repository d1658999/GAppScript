---
mode: 'agent'
---
## Requirements:

### Google App script is simulated under JAVASCRIPT

### ini file: #file: DPD_Smart_Char_LTE_NR_Vpa_Search.ini


### Build a JAVASCRIPT script named `get_dc2dc_level.js` that processes the ini file `DPD_Smart_Char_LTE_NR_Vpa_Search.ini` and extracts, for every band section, the following keys:
- Pa control table0 dc2dc level Num0
- Pa control table1 dc2dc level Num0
- Pa control table0 dc2dc level Num1
- Pa control table1 dc2dc level Num1
- Pa control table0 dc2dc level Num2
- Pa control table1 dc2dc level Num2
- Pa control table0 dc2dc level Num3
- Pa control table1 dc2dc level Num3
- Pa control table0 dc2dc level Num4
- Pa control table1 dc2dc level Num4
- Pa control table0 dc2dc level Num5
- Pa control table1 dc2dc level Num5

### The script should output a CSV file with the following columns:
1. **Band Name**: The section name from the ini file (e.g., `LTE BAND1 CID0 BW0 DPD`).
2. **DC2DC Level Table**: The key name (e.g., `Pa control table0 dc2dc level Num0`).
3. **DC2DC Level Value**: The value for that key (e.g., `1.2V,1.2V,1.2V,1.2V,1.2V,1.2V,1.2V,1.2V`).

Each row in the CSV should correspond to a single key for a single band, so there will be 12 rows per band.

### The Python file name must be: `get_dc2dc_level.js`
### The google drive link for ini file is at gsheet 'A1' for the first tab
### capture the file link and start to parse the ini file and the result is output at the second tab, whose tab name is 'DC2DC Level'