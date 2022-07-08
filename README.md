# CK3Tools-Titles-to-MediaWiki
A tool that reads titledata from game saves and mod files to produce a MediaWiki table full of data.

Installation: Just extract it. Make sure it is in its own folder.

- Step 1: Start up the mod in CK3. Entering into observer mode, save the state of the game at the immediate start.
-- **Note - The only way this works with multiple start dates is if you use this tool individually on each. Merging is not possible.**
- Step 2: Copy the save into the root folder of this portable tool.
- Step 3: Run. After a few seconds it will show a prompt that goes, "Enter output code or path to file with codes." Currently the way the tool is told the format in which to write the table is through a string of code. This is a CLI app. I'm sorry.
- Step 4: Paste in either: 
  1 - The individiual code 
  2 - Or the *file path* of a text file that contains multiple codes for batch processing. 
- The output files will be deposited on your desktop.

Dictionary of commands:

Firstly, type what type of title that you need the data of:

| Command | Definition |
| --- | --- |
| `DJE` | De Jure Empires |
| `TE` | Titular Empires |
| `DJK` | De Jure Kingdoms |
| `TK` | Titular Kingdoms |
| `DJD` | De Jure Duchies |
| `TD` | Titular Duchies |
| `C` | Counties |

Secondly, add parameters as reasonable:

| Parameter | Definition |
| --- | --- |
| `t` | Title IDs |
| `n` | Names |
| `le` | Empire Liege |
| `lk` | Kingdom Liege |
| `ld` | Duchy Liege |
| `vk` | Vassal Kingdoms |
| `cvk` | Number of Vassal Kingdoms |
| `vd` | Vassal Duchies |
| `cvd` | Number of Vassal Duchies |
| `vc` | Vassal Counties |
| `cvc` | Number of Vassal Counties |
| `vb` | Vassal Baronies |
| `cvb` | Number of Vassal Baronies |
| `lc` | Size of Largest County by Baronies |
| `cpt` | Capitals (County) |
| `clt` | Culture (Counties-only) |
| `f` | Faith (Counties-only) |
| `dt` | Total Development (non-Counties) |
| `dh` | Development of Highest Dev. County Vassal |
| `dc` | Development (Counties-only) |
| `sb` | Special Buildings in Baronies |
| `m` | Modifiers in Counties |
| `sbm` | `sb` and `m` combined |

**Note: Parameters that don't make sense for a title rank or type shouldn't be added. Adding Duchy Liege for Empires will just output blanks. It won't break but there's no point.**

These are the personal settings that I use (with batch instead of one by one):
```
DJE n cpt vk cvc t
TE n cpt vk t
DJK n cpt vd le cvc t
TK n cpt le t
DJD n lk le cpt vc cvc lc cvb dt dh sbm t
TD n lk le cpt t
C n ld lk le cvb dc sbm f clt t
```
I don't advise adding more options to these. You might even remove a few. Too much information doesn't work well with MediaWiki. At best it's difficult to upload, at worst the page won't show properly or nicely.
