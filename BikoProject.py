import os
import pandas as pd
import xlsxwriter

# I set this program up to be pretty easy to update in the future, I'll leave instructions here on modification.
# If you want to add a new worksheet or change a worksheet name then add it into the worksheets list.
#
# Within the Person class each Unnamed value corresponds to a column index in team roles
#
#
#
#

# creates the final excel file
outputName = input("What do you want your file name to be?\n")
seasonInput = input("Spring or Fall?\n")

finalExcel = xlsxwriter.Workbook(outputName + ".xlsx")

# creates a list with the worksheet names
worksheets = ["contents" ,"key" ,"RAW SCORE DATA", "scoreADX", "scoreACX", "scoreARanks", "ScoreWDX",
 "scoreWCX", "scoreWRanks", "scoreASpeech", "rawDiffData", "diffACX", "diffWCX", "diffASpeech"]

# dynamically creates worksheet variables
for sheet in worksheets:
    locals()[sheet] = finalExcel.add_worksheet(sheet)

# searches directory for Excel files parses them and then adds them to a dataset

dataset = {}
path = os.getcwd() + "\\Data"
for file in os.listdir(path):
    if file.endswith(".xlsx") or file.endswith(".xlsm"):
        dataset.update({os.path.splitext(file)[0]: pd.read_excel(path +"\\" + file, na_values="--", sheet_name=None, skiprows=[0])})
path += "\\"+seasonInput
print("Loading Excel Sheets...")
for tournamentFolder in os.listdir(path):
    tmp = {}
    for file in os.listdir(path+"\\"+tournamentFolder):
        print(file)
        if file.endswith(".xlsx") or file.endswith(".xlsm"):
            assert os.path.isfile(path + "\\"+tournamentFolder + "\\" + file)
            tmp.update({os.path.splitext(file)[0]: pd.read_excel(path + "\\"+tournamentFolder + "\\" + file,
                                                                 na_values="--", sheet_name=None,skiprows=[0])})
    dataset.update({tournamentFolder : tmp})
#for key in dataset.get('Black Lion - 2 REG').get('Preliminary Data').keys():
#   print(key + "\n")
#   print(dataset.get('Black Lion - 2 REG').get('Preliminary Data').get(key))


# populates the worksheets
worksheet = locals().get(worksheets[2])
worksheet.set_column('A:A', 20)
bold = finalExcel.add_format({'bold': True})
worksheet.write('A1', 'Hello')
worksheet.write('A2', 'World', bold)
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

#worksheets[0].add_format('')

# class that contains values for each person
class Person:

    def __str__(self):
        return ": % s, % s, % s, % s, % s, % s, % s" % (
            self.teamName, self.sideAsA, self.witnessDX, self.witnessCX,
            self.sideAsW, self.witnessMain, self.witnessConting)

    def __init__(self, this):
        # Datasheet
        self.teamName = this.keys()[0]
        self.name = this.keys()[1]
        self.sideAsA = this.keys()[2]
        self.witnessDX = this.keys()[3]
        self.witnessCX = this.keys()[4]
        self.sideAsW = this.keys()[5]
        self.witnessMain = this.keys()[6]
        self.witnessConting = this.keys()[7]

        # Ballot Data
        self

        # more properties

pass
seasonList = dataset.get('Team Roles')
container = {}
for season in seasonList:
    tmp = []
    counter = len(seasonList.get(season))
    while counter:
        tmpPerson = Person(pd.DataFrame(seasonList.get(season), index=[counter]))
        print(pd.DataFrame(seasonList.get(season), index=[counter]).keys())
        counter -= 1
        tmp.insert(counter-1, {tmpPerson.name : tmpPerson})
    container.update({season: tmp})
# struct
# container = {Spring{Person{general stuffs, {tournament 1}, {tournament 2}}, Person{}, Person{}}, Fall{Person{}, Person{}, Person{}}}

print(container)
#for season in container:
    #for person in container.get(season):
        #for a in person:
            #print(a)
finalExcel.close()