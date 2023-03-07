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
            self.attyCX, self.attyDX, self.witConting)

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

        # more properties

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
i=0
for folder in dataset:
    for file in dataset.get(folder)[0]:
        # Identification
        currName = dataset.get(folder)[0].get(file).at[1,"Unnamed: %s" % 2]

        # Indiv Scores

        # Plaintiff
        currDXP = dataset.get(folder)[0].get(file).at[2,"Unnamed: %s" % 3]
        currCXP = dataset.get(folder)[0].get(file).at[2,"Unnamed: %s" % 4]
        currSpeechP = dataset.get(folder)[0].get(file).at[2,"Unnamed: %s" % 5]
        currRPP = dataset.get(folder)[0].get(file).at[2,"Unnamed: %s" % 6]

        # Defense
        currDXD = dataset.get(folder)[0].get(file).at[3,"Unnamed: %s" % 3]
        currCXD = dataset.get(folder)[0].get(file).at[3, "Unnamed: %s" % 4]
        currSpeechD = dataset.get(folder)[0].get(file).at[3, "Unnamed: %s" % 5]
        currRPD = dataset.get(folder)[0].get(file).at[3, "Unnamed: %s" % 6]

        # Differeneces
        #print(dataset.get(folder)[1].get(file))
        # Plaintiff
        plaintiffCXRound1Judge1 = dataset.get(folder)[1].get(file).at[4,"Unnamed: %s" % 3]
        plaintiffCXRound1Judge2 = dataset.get(folder)[1].get(file).at[5,"Unnamed: %s" % 3]
        plaintiffCXRound2Judge1 = dataset.get(folder)[1].get(file).at[4,"Unnamed: %s" % 4]
        plaintiffCXRound2Judge2 = dataset.get(folder)[1].get(file).at[5,"Unnamed: %s" % 4]
        plaintiffSpeechRound1Judge1 = dataset.get(folder)[1].get(file).at[4,"Unnamed: %s" % 5]
        plaintiffSpeechRound1Judge2 = dataset.get(folder)[1].get(file).at[5,"Unnamed: %s" % 5]
        plaintiffSpeechRound2Judge1 = dataset.get(folder)[1].get(file).at[4,"Unnamed: %s" % 6]
        plaintiffSpeechRound2Judge2 = dataset.get(folder)[1].get(file).at[5,"Unnamed: %s" % 6]

        # Defense
        defenseCXRound1Judge1 = dataset.get(folder)[1].get(file).at[6,"Unnamed: %s" % 3]
        defenseCXRound1Judge2 = dataset.get(folder)[1].get(file).at[7,"Unnamed: %s" % 3]
        defenseCXRound2Judge1 = dataset.get(folder)[1].get(file).at[6,"Unnamed: %s" % 4]
        defenseCXRound2Judge2 = dataset.get(folder)[1].get(file).at[7,"Unnamed: %s" % 4]
        defenseSpeechRound1Judge1 = dataset.get(folder)[1].get(file).at[6,"Unnamed: %s" % 5]
        defenseSpeechRound1Judge2 = dataset.get(folder)[1].get(file).at[7,"Unnamed: %s" % 5]
        defenseSpeechRound2Judge1 = dataset.get(folder)[1].get(file).at[6,"Unnamed: %s" % 6]
        defenseSpeechRound2Judge2 = dataset.get(folder)[1].get(file).at[7,"Unnamed: %s" % 6]

        print(defenseCXRound1Judge1)
#print(People)

#finalExcel.close()
