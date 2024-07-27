The accompanied program is to evaluate rating based on Fuzzy TOPSIS method. The program is primarily intended for vendor assessment, but it can be used for evaluating any other alternatives.

The program uses three files:

fzTopsis.py   -- This is the analysis program. It is written for python 3. 
ParamWt.xlsx  -- Contains the list of parameters to be used for analysis and its weights. 
rating.xlsx   -- Contains evaluation ratings. Data needs to be arranged in Expert, Alternative order. 

Result of the analysis is written in TOPSIS_Result.xlsx.

All the file names are hard coded in the python program.

Note: 
1. The order and number of parameters in ParamWt and rating has to match. Any mismatch will cause unintended results. 
2. All the xlsx files are Microsoft Excel files. 

License:  Use of the Source Code and software package is subject to the terms of the Mozilla Public License, v. 2.0. If a copy of the MPL was not distributed with this package, You can obtain one at https://mozilla.org/MPL/2.0/.
