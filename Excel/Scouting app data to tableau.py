import csv
import pandas as pd
import numpy as np
import os
from collections import defaultdict
import json
import math
import pprint
import math
# NOTE: LOOK AT SPREADSHEET DATA FROM LAST YEAR to understand i-values
# Gets data from spreadsheet
allData = pd.read_excel(r"https://bendlapine-my.sharepoint.com/personal/lairdtrout_student_bend_k12_or_us/Documents/ScoutingPASS_Excel_Example.xlsm?web=1")
allTeams = allData.t.unique()
allDefenders = allData.who.unique()
# Making dicts to hold data
teamAvgScore = {}
teamAvgAutoScore = {}
teamAvgEndgameScore ={}
dvoa = {}
teamTeleScoreList = {}
teamAutoScoreList = {}
reliability = {}
# changes from letter to charging score
def chargeTranslation(letter, auto):
    attempted = 0
    score = 0
    if letter != "x":
        attempted = 1
    if letter == "d":
        score = 6
        if auto:
            score += 2
    elif letter == "e":
        score = 10
        if auto:
            score += 2
    if auto: 
        if attempted != 0:
            return score
        else:
            return 0
    else:
        return score
        
# goes from cone/cube location to points    
def scoreTranslation(asdf, teleop):
    scoreTot = 0
    listOfInts = json.loads(asdf)
    if teleop:
        for i in listOfInts:
            if i <= 9:
                scoreTot += 5 + 1.67
            elif i <= 18:
                scoreTot += 3 + 1.67
            else:
                scoreTot += 2 + 1.67
    else:
        for i in listOfInts:
            if i <= 9:
                scoreTot += 6 + 1.67
            elif i <= 18:
                scoreTot += 4 + 1.67
            else:
                scoreTot += 3 + 1.67
    return scoreTot
# def defensiveAnalysis(team):
#     global dvoa    

#     asdf = allData.groupby(allData.who)
#     try:
#         group = asdf.get_group(team)
#         total = 0
#         counter = 0
#         for i in group.values:
#             total += (scoreTranslation(i[7], True) - teamAvgScore[i[4]])
#             counter += 1
#         dvoa[team] = total/counter
#     except:
#         pass
 
# fills in data for each team
def dataPopulation(team):
    global teamAvgAutoScore
    global teamAvgEndgameScore
    global teamAvgScore
    global reliability
    # Creates different groups of data
    asdf = allData.groupby(allData.t)
    group = asdf.get_group(team)
    defense = group.groupby(group.wd)
    loopNum = 0
    scoreTot = 0
    timeTot = 0
    teamTeleScoreList[team] = []
    # translating grid placements to score
    for i in group.values:
        loopNum += 1
        scoreTot += scoreTranslation(i[8], True)
        teamTeleScoreList[team].append(scoreTranslation(i[8], True))
    scoreTot /= loopNum
    removeList = ""
    for i in teamTeleScoreList[team]:
        removeList += (str(i) + ",") 
    print(removeList)
    teamTeleScoreList[team] = removeList
    teamAvgScore[team] = round(scoreTot, 0)
    loopNum = 0
    scoreTot = 0
    chargeTot = 0
    chargeAttempt = 1
    acc = 0
    for i in group.values:
        loopNum += 1
        acc += i[6]*3
        scoreTot += scoreTranslation(i[5], False)
        a = chargeTranslation(i[6], True)
        chargeTot += a
        if a != 0 and loopNum != 1:
            chargeAttempt += 1
    teamAvgAutoScore[team] = (scoreTot/loopNum) #+ (chargeTot/chargeAttempt)
    chargeTot = 0
    loopNum = 0
    for i in group.values:
        loopNum += 1
        chargeTot += chargeTranslation(i[11], False)
    teamAvgEndgameScore[team] = chargeTot/loopNum
    # calcualtes reliability
def offensiveAnalysis(team):
    global reliability
    asdf = allData.groupby(allData.t)
    group = asdf.get_group(team)
    rScore = 0
    for i in group.values:
        rScore -= i[14]
    reliability[team] = rScore
#writes data to a file   
def fileWriter():
    global teamAvgAutoScore
    global teamAvgEndgameScore
    global teamAvgScore
    global reliability
    dd = defaultdict(list)
    ff = 0
    # going through and converitng values
    for i in teamAvgScore.values():
        bb = list(teamAvgScore.keys())
        # fixing null values
        if math.isnan(i):
            asdfadsf = bb[ff]
            teamAvgScore[asdfadsf] = 0
        ff +=1 
    # writing different dictionaries to main dictionary
    for d in (teamAvgAutoScore, teamAvgEndgameScore, teamAvgScore, reliability, teamTeleScoreList, {}): # you can list as many input dicts as you want here
        for key, value in d.items():
            dd[key].append(value)
    # writing data to csv
    path = r""C:\Users\user\Desktop\ScoutingAppData.csv""
    assert os.path.isfile(path) 
    with open(path, 'w') as g:
        holderArray = []
        # creates CSV with these headers
        headers = ["team", "autonomous", "endgame", "teleop", "reliability", "TeamTeleopScoregraphed"]
        # using csv.writer method from CSV package
        write = csv.writer(g)
        write.writerow(headers)
        for i in range(len(list(dd.keys()))):
            x = []
            x.append(list(dd.keys())[i])
            x.extend(list(dd.values())[i])
            holderArray.append(x)
        for row in holderArray:
            write.writerow(row)
def allAnalysis():
    for i in allTeams:
        if type(i) == np.int64:
            dataPopulation(i)
    for i in allTeams:

        if type(i) == np.int64:
            offensiveAnalysis(i)
           # defensiveAnalysis(i)
    fileWriter()
allAnalysis()
        
