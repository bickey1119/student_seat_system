import openpyxl

kr = [ #国理34席
    [1,3,5,7,9,11,14,16,18,19,21,23,2,4,6,8,10,12,13,15,17,20,22,24,26,28,30,32,34,25,27,29,31,33],
    ['B1','D1','F1','H1','J1','L1','C10','E10','G10','I10','K10','M10','C1','E1','G1','I1','K1','M1','B10','D10','F10','J10','L10','B13','D13','F13','I13','K13','M13','C13','E13','G13','J13','L13'],
]
kb = [ #国文29席
    [35,37,39,40,42,44,45,47,49,50,52,54,36,38,41,43,46,48,51,53,55,56],
    ['C24','E24','G24','I24','K24','M24','C27','E27','G27','I27','K27','M27','D24','F24','J24','L24','D27','F27','J27','L27','C35','D35'],
]
sr = [ #私理7席
    [57,59,60,62,58,61,63],
    ['E35','G35','J35','L35','F35','K35','M35'],
]
sb = [ #私文11席
    [64,66,68,69,71,73,65,67,70,72,74],
    ['C38','E38','G38','C49','E49','G49','D38','F38','D49','F49','H49'],
]


file_path_1 = '6月模試.xlsx' #ここを変える

wb1=openpyxl.load_workbook(file_path_1,data_only=False) 
ws1 = wb1['フォーマット']


file_path_2 = '座席表編集用2.xlsx' #ここを変える

wb2=openpyxl.load_workbook(file_path_2,data_only=False) 
ws2 = wb2['SVL']

i = 4
j = 0
k = 0
l = 0
m = 0
while True:
    if ws1['E'+str(i)].value == None:
        break
    
    if ws1['E'+str(i)].value == '国理':
        ws1['Y'+str(i)].value = kr[0][j]
        ws2[kr[1][j]].value = ws1['A'+str(i)].value
        j = j+1
    elif ws1['E'+str(i)].value == '国文':
        ws1['Y'+str(i)].value = kb[0][k]
        ws2[kb[1][k]].value = ws1['A'+str(i)].value
        k = k+1
    elif ws1['E'+str(i)].value == '私理':
        ws1['Y'+str(i)].value = sr[0][l]
        ws2[sr[1][l]].value = ws1['A'+str(i)].value
        l = l+1
    elif ws1['E'+str(i)].value == '私文':
        ws1['Y'+str(i)].value = sb[0][m]
        ws2[sb[1][m]].value = ws1['A'+str(i)].value
        m = m+1

    i = i+1
    wb1.save(file_path_1) 
    wb2.save(file_path_2) 

print('complete')