
# Reqs
1) S column value to AH column in "mmm-yy" format
2) AI: Vlookup from other file and match AF column in this file to AF value in that file and get corresponding AI value from there. Here, instead of directly pasting the value, adding VLOOKUP is expected bcz change in input cell will update this cell automatically (Formula expected)
3) AJ: Subtract today's date with column J and add number in INTEGER
4) AK: Age range
  - 0 to 3 days
  - 4 to 9 days
  - 10 to 19 days
  - 20 to 29 days
  - more than a month
5) AL: (CC Status): Vlookup AD column from other file named cc\ details.xlsx
6) AM: Today's date
7) AE: control column width: 35
8) AF, AG, AH: Autofit column width
9) AG, AH: extra rows
10) AG Value mismatch: Bug
11) AC: 
 - Two digit decimal value.
 - Background yellow if greater than 40000.
 - W: ZEWP, ZGOW, ZGW1: background orange for AC field
 - When both conditions matched, then light blue
12) D: convert to number
13) F: (Cluster) Vlookup using column D from another file name SAPuser.xlsx
14) main pivot table
 - C to P: column width: 7.5
 - Number columns are center aligned
 - currency column are right aligned
 - B: left aligned
 - P: round off to INTEGER
 - heading merge and center upto column P
 - whole sheet middle align vertically
 - box border whole table
