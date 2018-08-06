#List of Symptoms
#Pain during Sex
#Bleeding after sex
#Bleeding in between periods
#Pain in back or pelvis
#abnormal vaginal discharge
#change in bleeding

import xlrd
import xlwt

class Algo:
    
    def __init__(self, background):
        self.mybackground = background
        self.sympList = {'sexPain':False, 'sexBleeding':False, 'bleedingBtwn':False,'pelvPain':False,\
                'backPain':False, 'vaginalDischarge':False, 'abCycle':False, 'abBleeding':False}
        self.abnormalFlag = False
        self.output = ""

    def sexPain(self):
        if self.mybackground.painDuringSex():
            self.abnormalFlag = True
            self.output += "\tPain during sex in conjunction with nausea\n"

    def sexBleeding(self):
        if self.mybackground.bleedingAfterSex():
            self.abnormalFlag = True
            self.output += "\tBleeding after sex\n"

    def bleedingBtwn(self):
        result = self.mybackground.bleedingBetween()
        concerning = result[0]
        listOfTimes = result[1]
        num = len(listOfTimes)
        if concerning:
            self.abnormalFlag = True
            printStr = "\tBled between period " + str(num) + " times within your last cycle:\n"
            for i in range(num):
                printStr += ("\t\tTime " + str(i+1) + ": " + str(listOfTimes[i]))
                if listOfTimes[i] == 1: printStr += " day\n"
                else: printStr += " days\n"
            self.output += printStr

    def pelvPain(self):
        concerning = self.mybackground.pelvicPain()[0]
        othSymps = self.mybackground.pelvicPain() [1]
        count = 0
        if self.mybackground.pelvicPain():
            self.output += "\tPain in pelvic region coupled with"
            for symp in othSymps:
                count += 1
                if len(othSymps)>1 and count < len(othSymps) and count> 1: self.output += ","
                if len(othSymps)>1 and count == len(othSymps): self.output += " and"
                self.output += str(symp)
            self.output += "\n"
            self.abnormalFlag = True

    def backPain(self):
        concerning = self.mybackground.backPain()[0]
        othSymps = self.mybackground.backPain() [1]
        count = 0
        if self.mybackground.pelvicPain():
            self.output += "\tBack pain coupled with"
            for symp in othSymps:
                count += 1
                if len(othSymps)>1 and count == len(othSymps): self.output += " and"
                self.output += str(symp)
            self.output += "\n"
            self.abnormalFlag = True

    def utiOutput(self):
        concerning = self.mybackground.UTI()
        if concerning:
            self.output += "\tBlood in urine and a burning sensation were recorded. These "
            self.output += "symptoms align with common symptoms of urinary tract infections. \n"
            self.abnormalFlag = True
    
    def vaginalDischarge(self):
        vDischarge = self.mybackground.vaginalDischarge()
        concerning = vDischarge[0]
        continuous = vDischarge[1]
        itching = vDischarge[2]
        descriptOfDis = vDischarge[3]
        if concerning:
            self.abnormalFlag = True
            self.output += "\t"
            if continuous:  self.output += "Continuous"
            for des in descriptOfDis:
                if continuous:  self.output += ", "
                self.output += des
                if des != descriptOfDis[len(descriptOfDis)-1]: self.output += ", "
            self.output += " vaginal discharge"
            if itching:
                self.output += " coupled with itching or burning in the pelvic area"
            self.output += "\n"
        
    def abCycle(self):
        cycle = self.mybackground.abnormalCycle()
        concerning = cycle[0]
        string = cycle[1]
        if concerning:
            self.output += "\t" + string + "\n"

    def abBleeding(self):
        bleeding = self.mybackground.abnormalBleeding()
        concerning = bleeding[0]
        string = bleeding[1]
        if concerning:
            self.output += "\t" + string + "\n"
    
def main():
    file_location = "C:/Users/SDG Solution Space/Documents/aycFakeData.xlsx"
    workbook = xlrd.open_workbook(file_location)
    dataSheet = workbook.sheet_by_index(0)
    myNormal = workbook.sheet_by_index(1)
    myData = Data(dataSheet, myNormal)
    person = Algo(myData)
    exec_methods = ["sexPain", "sexBleeding", "bleedingBtwn", "pelvPain", "backPain",\
                   "utiOutput", "vaginalDischarge", "abCycle", "abBleeding"]
    for method in exec_methods:
       getattr(person, method)()
    if person.abnormalFlag:
        out = "You should see a healthcare practitioner. The following concerning symptoms were recorded: \n" + \
                             person.output
        print(out)

class Data():
    def __init__(self, dataSheet, normalSheet):
        self.sheet = dataSheet
        self.currDay = self.sheet.nrows -1
        self.dictOfSymps = {'sexPain':1, 'sexBleeding':2, 'bleedingBtw':3, 'pelvPain':4, 'backPain':5,\
                           'vaginalDischarge':6, 'nausea':7, 'bloodInUrine':8, 'fever':9, 'burning':10,\
                            'excessiveBleeding':11, 'period':12}
        self.normal = normalSheet
        self.currMonth = self.normal.nrows-1
        self.dictOfNorm = {'numOfProds':1, 'lengthOfCycle':2}
        self.normalLine =  1
        self.period = self.sheet.cell_value(self.currDay, self.dictOfSymps['period'])
        
    #returns True if nausea occurs on same day as pain during sex
    def painDuringSex(self):
        concerning = False
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['sexPain']) == 1:
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['nausea']) == 1:
                flag = True;
        return concerning

    def bleedingAfterSex(self):
        past = False
        concerning = False
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['sexBleeding']) == 1:
            for day in range(self.sheet.nrows -1):
                if self.sheet.cell_value(day, self.dictOfSymps['sexBleeding']) == 1:
                    past = True
        if past and self.period == 0:
            numOfHours = input("How many hours did the bleeding last?\t")
            if float(numOfHours) > 2: concerning = True
        return concerning

    def bleedingBetween(self):
        concerning= False
        countTimes = 0
        countDays = 0
        listOfLengths = []
        flagBleeding = False
        for day in range(self.sheet.nrows):
            if self.sheet.cell_value(day, self.dictOfSymps['bleedingBtw']) == 1:
                if flagBleeding == False:
                    flagBleeding = True
                    countTimes += 1
                countDays += 1
            else:
                if flagBleeding:
                    flagBleeding = False
                    listOfLengths.append( countDays)
                    countDays = 0
        if countTimes > 1: concerning = True
        result = (concerning, listOfLengths)
        return result

    def pelvicPain(self):
        concerning = False
        othSymps = []
        outputTuple = (concerning, othSymps)
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['pelvPain']) == 1:
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['bleedingBtw']) == 1:
                concerning = True
                othSymps.append(" bleeding between periods")
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['fever']) == 1:
                concerning = True
                othSymps.append(" fever")
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['nausea']) == 1:
                concerning = True
                othSymps.append(" nausea")
        return outputTuple

    def backPain(self):
        concerning = False
        othSymps = []
        outputTuple = (concerning, othSymps)
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['backPain']) == 1:
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['bloodInUrine']) == 1 and self.period == 0:
                concerning = True
                othSymps.append(" blood in urine")
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['bleedingBtw']):
                concerning = True
                othSymps.append(" bleeding between periods")
        return outputTuple

    def UTI(self):
        concerning = False
        if (self.sheet.cell_value(self.currDay, self.dictOfSymps['bloodInUrine']) == 1 \
           and self.sheet.cell_value(self.currDay, self.dictOfSymps['burning']) == 1) :
            concerning = True
        return concerning

    def vaginalDischarge(self):
        concerning = False
        itching = False
        cont = False
        descriptOfDis = []
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['vaginalDischarge']) == 1:
            bloody = input("Was the discharge bloody (Yes or No)?\t")
            color = input("Was the discharge pale (Yes or No)?\t")
            continuous = input("Was the discharge heavier than normal (Yes or No)?\t")
            smell = input("Was the discharge foul smelling (Yes or No)?\t")
            if (continuous.lower() == "yes"):
                cont = True
                if (color.lower() == "yes"):
                    concerning = True
                    descriptOfDis.append("pale")
                if (bloody.lower() == "yes"):
                    concerning = True
                    descriptOfDis.append("bloody")
                if (smell.lower() == "yes"):
                    concerning = True
                    descriptOfDis.append("foul smelling")
            if self.sheet.cell_value(self.currDay, self.dictOfSymps['burning']) == 1:
                itching = True
                if color.lower() == "yes" and cont == False:
                    concerning = True
                    descriptOfDis.append("pale")
                if (bloody.lower() == "yes" and cont == False):
                    concerning = True
                    descriptOfDis.append("bloody")
                if (smell.lower() == "yes" and cont == False):
                    concerning = True
                    descriptOfDis.append("foul smelling")
        return (concerning, cont, itching, descriptOfDis)

    def abnormalCycle(self):
        currLength = self.normal.cell_value(self.currMonth, self.dictOfNorm['lengthOfCycle'])
        concerning = False
        count = 0
        norm = self.normal.cell_value(self.normalLine, self.dictOfNorm['lengthOfCycle'])
        row = self.currMonth
        string = ""
        while (self.normal.cell_value(row, self.dictOfNorm['lengthOfCycle']) - norm >= 7):
            count += 1
            row -= 1
        if count > 1:
            concerning = True
            string = "Your cycle has been more than 7 days longer than usual for the past "\
                     + str(count) + " cycles"
        elif currLength < 21 or currLength > 45:
            concerning = True
            string = "Your cycle was " + str(numOfDays) + " days"
        outputTuple = (concerning, string)
        return outputTuple

    def abnormalBleeding(self):
        numOfProds = self.normal.cell_value(self.currMonth, self.dictOfNorm['numOfProds'])
        concerning = False
        string = ""
        norm = self.normal.cell_value(self.normalLine, self.dictOfNorm['numOfProds'])
        row = self.currMonth
        howAbnormal = self.normal.cell_value(self.currMonth, self.dictOfNorm['numOfProds'])/norm
        count = 0
        while (self.normal.cell_value(row, self.dictOfNorm['numOfProds'])/norm >= 2 or \
            self.normal.cell_value(row, self.dictOfNorm['numOfProds'])/norm <= .5):
            count += 1
            row -= 1
        if self.sheet.cell_value(self.currDay, self.dictOfSymps['excessiveBleeding']) == 1:
            oneHour = input("Have you gone through or filled a menstrual product (tampon, pad, menstrual) in an hour (yes or no)?\t")
            if (oneHour.lower() == "yes"):
                print("hi")
                concerning = True
                string += "Filled a menstrual product in one hour"
        if count > 1:
            if concerning == True: string += "\n\t"
            else: concerning = True
            string += "In your last " + str(count) + " cycles, you have bled aproximately " \
                     + str(howAbnormal) + " times as much as usual"
        outputTuple = (concerning, string)
        return (outputTuple)

 
                
    
            
                       

main()

