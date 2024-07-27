#!/usr/bin/python3
'''
* fzTopsis - program to evaluate Fuzzy TOPSIS ratings
*
* Author: Soumyanath Chatterjee
* Year:   2024
*
* This Source Code Form is subject to the terms of the Mozilla Public
* License, v. 2.0. If a copy of the MPL was not distributed with this
* file, You can obtain one at https://mozilla.org/MPL/2.0/.
*
'''
import numpy as np
import re
import openpyxl 
from openpyxl.styles import Font

import pandas as pd
import pprint
import math
import copy

pp = pprint.PrettyPrinter(width=41, compact=True)

# Functions  ----------------------------------------
def D(a,b):
    # Measures distance between A and B
    d = math.sqrt(((a[0]-b[0])*(a[0]-b[0]) + (a[1]-b[1])*(a[1]-b[1]) + (a[2]-b[2])*(a[2]-b[2]))/3)
    return d
    
# Config --------------------------------------------
sParamFile  = "ParamWt.xlsx"
sRatingFile = "ratings.xlsx"
sResultFile = "TOPSIS_Result.xlsx"

#--------- Credit Screen ------------------------------
print(80*"*")
print("*",76*" ","*")
print("*                           Fuzzy TOPSIS Evaluation                            *") 
print("*",76*" ","*")
print(80*"*")

# Initialize ----------------------------------------
rScale = {"A":(0.75,0.9,1.0), "B":(0.5, 0.75, 0.9), "C":(0.3, 0.5, 0.75),\
          "D":(0.1, 0.3, 0.5), "F":(0.0, 0.1, 0.3)}
#pp.pprint(rScale)
# Read Wt
attrib = pd.read_excel(sParamFile)

print("Analyzing Data:") 
        
# Read ratings 
ratings = pd.read_excel(sRatingFile)
#pp.pprint(ratings)

# Filter alternatives and Experts
alternatives = list(set(ratings["Alternative"]))
alternatives.sort()
experts = list(set(ratings["Expert"]))
experts.sort()
#pp.pprint(alternatives)
#pp.pprint(experts)

# Conv to FN
rFN = {}
for j, t in ratings.iterrows():
    tFN = {}
    for i, v in t.items():
        tFN[i]=v if i in ["Expert", "Alternative"] else list(rScale[v])
    rFN[j]=tFN
#pp.pprint(rFN)

# Combine ratings
cFN = {}
attrib_dim = attrib["Parameter"]
#pp.pprint(attrib_dim)
for i in alternatives:
    cAttr = {}
    for j in attrib_dim:
        a = []
        b = []
        c = []
        # print(i,j)
        for k in rFN:
            if rFN[k]["Alternative"]== i:
                a.append(rFN[k][j][0])
                b.append(rFN[k][j][1])
                c.append(rFN[k][j][2])
        ai = min(a)
        bi = sum(b)/len(b)
        ci = max(c)
        cAttr[j] = [ai,bi,ci]
    cFN[i] = cAttr
print("1. Combined Table")    
dFcFN = pd.DataFrame(cFN)
#pp.pprint(dFcFN)
       
# Normalize
cStar = {}
for a in attrib_dim:
    cj = []
    for i in alternatives:
        cj.append(cFN[i][a][2])
    cStar[a] = max(cj)
#print(cStar)    
nFN = copy.deepcopy(cFN)
for a in attrib_dim:
    for i in alternatives:
        for j in [0,1,2]:
            nFN[i][a][j] = nFN[i][a][j]/cStar[a]
print("2. Normalized Table")                
#pp.pprint(nFN)



# GET FPIS, FNIS
fpis = {}
fnis = {}
for a in attrib_dim:
    fpis[a] = nFN[alternatives[0]][a]
    fnis[a] = nFN[alternatives[0]][a]
    for i in alternatives:
        fpis[a] = fpis[a] if fpis[a][1] >= nFN[i][a][1] else nFN[i][a]
        fnis[a] = fnis[a] if fnis[a][1] <= nFN[i][a][1] else nFN[i][a]

print("3. FPIS & FNIS computed")


# Compute CC
cc = {}
ranks = {}
posn = []
dp = {}
dn = {}
attrib.set_index('Parameter', inplace=True)
#pp.pprint(attrib) 
for i in alternatives:
    dplus = 0
    dminus = 0
    for a in attrib_dim:
        dplus += attrib.loc[a,"Wt"]*D(nFN[i][a], fpis[a])
        dminus += attrib.loc[a,"Wt"]*D(nFN[i][a], fnis[a])
#         print(a, attrib.loc[a,"Wt"])


    cci = dminus/(dplus+dminus)
    posn.append(cci)
    cc[i] = cci
    dp[i] = dplus
    dn[i] = dminus
#pp.pprint(cc)  
#pp.pprint(dp)
#pp.pprint(dn)
posn.sort(reverse=True)
for i in alternatives:
    ranks[i] = posn.index(cc[i]) + 1
#pp.pprint(ranks) 
print("4. Ranks computed")
#print("\n",80*"~")

# Result
#fOut = pd.ExcelWriter(sResultFile)
with pd.ExcelWriter(sResultFile) as writer: 
    ratings.to_excel(writer, sheet_name = 'rating', index=False)
    dfRFN = pd.DataFrame(rFN).T
    dfRFN.to_excel(writer, sheet_name = 'fuzzyRating', index=False)
    # Format fuzzy values to 3 decimal places
    for i in attrib_dim:
        for j in alternatives:
            for k in [0,1,2]:
                dFcFN[j][i][k] = round(dFcFN[j][i][k],3)
    dFcFN.T.to_excel(writer, sheet_name = 'combined')
    dfNFN = pd.DataFrame(nFN).T
    for i in attrib_dim:
        for j in alternatives:
            for k in [0,1,2]:
                dfNFN[i][j][k] = round(dfNFN[i][j][k],3)
    
    dfNFN.to_excel(writer, sheet_name = 'normalized')
    dfRes = dfNFN
    dfRes['D_plus'] = dp
    dfRes['D_minus'] = dn
    dfRes['CCI'] = cc
    dfRes['Rank'] = ranks
    #dfRes.loc[len(dfRes.index)] = attrib['Wt'].T
    dfRes.loc['Wt'] = attrib['Wt'].T
    dfRes.loc['FPIS'] = fpis
    dfRes.loc['FNIS'] = fnis
    dfRes.to_excel(writer, sheet_name = 'Result', float_format = "%.3f")
    
print("\n** FINISHED** Result saved in file: ",sResultFile)
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~