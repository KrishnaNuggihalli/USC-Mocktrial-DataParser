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

# creates the final Excel file
outputName = input("What do you want your file name to be?\n")
seasonInput = input("Spring or Fall?\n")

finalExcel = xlsxwriter.Workbook(outputName + ".xlsx")

# creates a list with the worksheet names
worksheets = ["contents", "key", "RAW SCORE DATA", "scoreADX", "scoreACX", "scoreARanks", "ScoreWDX",
              "scoreWCX", "scoreWRanks", "scoreASpeech", "rawDiffData", "diffACX", "diffWCX", "diffASpeech"]

# dynamically creates worksheet variables
for sheet in worksheets:
    locals()[sheet] = finalExcel.add_worksheet(sheet)

# searches directory for Excel files parses them and then adds them to a dataset

dataset = {}
teamRoles = {}
path = os.getcwd() + "\\Data"
print("Loading Team Roles...")
for file in os.listdir(path):
    if file.endswith(".xlsx") or file.endswith(".xlsm"):
        teamRoles = pd.read_excel(path + "\\" + file, na_values="--", sheet_name = seasonInput + " Roles")
        print(file)
path += "\\" + seasonInput
print("Loading Tournament Sheets...")
for tournamentFolder in os.listdir(path):
    indivScores = {}
    differences = {}
    for file in os.listdir(path + "\\" + tournamentFolder):
        print(file)
        if file.endswith(".xlsx") or file.endswith(".xlsm"):
            assert os.path.isfile(path + "\\" + tournamentFolder + "\\" + file)
            differences.update({os.path.splitext(file)[0]: pd.read_excel(path + "\\" + tournamentFolder + "\\" + file,
                                                            na_values="--", sheet_name="Differences", skiprows=[0])})
            indivScores.update({os.path.splitext(file)[0]: pd.read_excel(path + "\\" + tournamentFolder + "\\" + file,
                                                            na_values="--", sheet_name="Indiv Scores")})
    dataset.update({tournamentFolder: [indivScores, differences]})
# for key in dataset.get('Black Lion - 2 REG').get('Preliminary Data').keys():
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


# worksheets[0].add_format('')

# class that contains values for each person
class Person:

    def __str__(self):
        return ": % s, % s, % s, % s, % s, % s, % s, % s" % (
            self.name, self.squad, self.roleOnP, self.roleOnD, self.witDX,
            self.attyCX, self.attyDX, self.witConting
        )
    def __init__(self, this):
        # Datasheet
        self.squad = this.get(this.keys()[0]).array[0]
        self.name = this.get(this.keys()[1]).array[0]
        self.roleOnP = this.get(this.keys()[2]).array[0]
        self.roleOnD = this.get(this.keys()[3]).array[0]
        self.attyDX = this.get(this.keys()[4]).array[0]
        self.attyCX = this.get(this.keys()[5]).array[0]
        self.witDX = this.get(this.keys()[6]).array[0]
        self.witConting = this.get(this.keys()[7]).array[0]

        # Ballot Data

        self.tournaments = {}

pass
class Tournament:

    def __str__(self):
        return ": % s, % s, % s, % s, % s, % s, % s, % s, % s, % s, % s, % s," \
               " % s, % s, % s, % s, % s, % s, % s, % s, % s, % s, % s, % s" % (
            self.DXP, self.CXP, self.speechP, self.RPP,
            self.DXD, self.CXD, self.speechD, self.RPD,
            self.plaintiffCXRound1Judge1, self.plaintiffCXRound1Judge2,
            self.plaintiffCXRound2Judge1, self.plaintiffCXRound2Judge2,
            self.plaintiffSpeechRound1Judge1, self.plaintiffSpeechRound1Judge2,
            self.plaintiffSpeechRound2Judge1, self.plaintiffSpeechRound2Judge2,
            self.defenseCXRound1Judge1, self.defenseCXRound1Judge2,
            self.defenseCXRound2Judge1, self.defenseCXRound2Judge2,
            self.defenseSpeechRound1Judge1, self.defenseSpeechRound1Judge2,
            self.defenseSpeechRound2Judge1, self.defenseSpeechRound2Judge2
        )

    def __init__(self, index, folder, file):
        def infoParser(worksheet, index, x, y):
            modifier = 6 if worksheet[0] == "Indiv Values" else 8
            return worksheet[1].at[(modifier * index) + x, "Unnamed: %s" % y]

        sheet = ["Indiv Values", folder[0].get(file)]

        # Identification
        self.name = infoParser(sheet, index, 1, 2)

        # Indiv Scores

        # Plaintiff
        self.DXP = infoParser(sheet, index, 2, 3)
        self.CXP = infoParser(sheet, index, 2, 4)
        self.speechP = infoParser(sheet, index, 2, 5)
        self.RPP = infoParser(sheet, index, 2, 6)

        # Defense
        self.DXD = infoParser(sheet, index, 3, 3)
        self.CXD = infoParser(sheet, index, 3, 4)
        self.speechD = infoParser(sheet, index, 3, 5)
        self.RPD = infoParser(sheet, index, 3, 6)

        # Differeneces
        sheet = ["Differences", folder[1].get(file)]
        # Plaintiff
        self.plaintiffCXRound1Judge1 = infoParser(sheet, index, 4, 3)
        self.plaintiffCXRound1Judge2 = infoParser(sheet, index, 5, 3)
        self.plaintiffCXRound2Judge1 = infoParser(sheet, index, 4, 4)
        self.plaintiffCXRound2Judge2 = infoParser(sheet, index, 5, 4)
        self.plaintiffSpeechRound1Judge1 = infoParser(sheet, index, 4, 5)
        self.plaintiffSpeechRound1Judge2 = infoParser(sheet, index, 5, 5)
        self.plaintiffSpeechRound2Judge1 = infoParser(sheet, index, 4, 6)
        self.plaintiffSpeechRound2Judge2 = infoParser(sheet, index, 5, 6)

        # Defense
        self.defenseCXRound1Judge1 = infoParser(sheet, index, 6, 3)
        self.defenseCXRound1Judge2 = infoParser(sheet, index, 7, 3)
        self.defenseCXRound2Judge1 = infoParser(sheet, index, 6, 4)
        self.defenseCXRound2Judge2 = infoParser(sheet, index, 7, 4)
        self.defenseSpeechRound1Judge1 = infoParser(sheet, index, 6, 5)
        self.defenseSpeechRound1Judge2 = infoParser(sheet, index, 7, 5)
        self.defenseSpeechRound2Judge1 = infoParser(sheet, index, 6, 6)
        self.defenseSpeechRound2Judge2 = infoParser(sheet, index, 7, 6)

pass
# struct
# People = {Person1{general stuffs, {tournament 1}, {tournament 2}}, Person2{}, Person3{}...}print("People: ")
People = {}

counter = len(teamRoles)



while counter:
    counter-=1
    currPerson = Person(pd.DataFrame(teamRoles, index=[counter]))
    # classic Del Gaudio workaround
    People[currPerson.name.split(" ", 1)[1]] = currPerson

for folder in dataset:
    for file in dataset.get(folder)[0]:
        for index in range(0,7):
            currTournament = Tournament(index, dataset.get(folder), file)
            if(People.keys().__contains__(currTournament.name)):
                People[currTournament.name].tournaments[folder] =  currTournament


print(People)
print("Del gaudio", People.get("Del Gaudio") , "| Tournament 1 ", People.get("Del Gaudio").tournaments.get("Tournament 1"),
      "| Tournament 2 ", People.get("Del Gaudio").tournaments.get("Tournament 2"))

#finalExcel.close()
