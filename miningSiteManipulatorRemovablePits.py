import csv as c
import copy
import string
import sys
import os

class mSite:
    def __init__(self, data : list[str] = []):
        self.raw = []

        for x in list(data):
            try : 
                self.raw.append(int(x))
            except ValueError:
                if len(x) == 0:
                    self.raw.append([])
                elif "," in x:
                    self.raw.append([int(y) for y in x.split(",")])
                else:
                    self.raw.append(x)

        if type(self.raw[-2]) == int:
            self.raw[-2] = [self.raw[-2]]
            self.raw[-3] = [self.raw[-3]]

        self.raw.pop(-1)

        self.raw.append(None)

class h:
    def __init__(self, csv : str):
        f = open(csv, "r")
        d = [x for x in c.reader(f, delimiter = ",")]
        f.close()

        self.indexs = d.pop(0)

        self.indexs[-1] = "Old OMP Number"

        self.miningSites = []

        for site in d:
            self.miningSites.append(mSite(site))

        self.newMiningSites = []

    def genNewSites(self):
        i = 0

        lookUpTable = {}

        for site in self.miningSites:
            lookUpTable.update({site.raw[1] : i + site.raw[3] - 1})

            i += site.raw[3]

        i = 0

        for site in self.miningSites:
            for day in range(site.raw[3]):
                tmpSite = copy.deepcopy(site)
                
                tmpSite.raw[-1] = site.raw[1]
                tmpSite.raw[1] = i
                tmpSite.raw[3] = 1
                tmpSite.raw[2] = round(site.raw[2]/site.raw[3])
                
                if day == 0:
                    tmpPreds = []
                    for pred in site.raw[-3]:
                        #print(site.raw)
                        tmpPreds.append(lookUpTable[pred])
                else:
                    tmpPreds = [i-1]

                tmpSite.raw[-3] = tmpPreds
                tmpSite.raw[-4] = len(tmpPreds)      

                tmpSite.raw[-2] = [1 for x in range(len(tmpSite.raw[-3]))]

                self.newMiningSites.append(tmpSite)
                i += 1

    def col2num(self, col):
        try:
            return int(col)
        except ValueError:
            pass
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1

        if (num == 0 and col.upper() != "A") or col == '':
            print(f"Invalid column int, not proccesing : {col}")
        return num

    def output(self, location : str, toRemove : list[int] = [], sep : str = ","):
        f = open(location + ".csv", "w")

        for i in range(len(self.indexs)):
            if i+1 not in toRemove:
                f.write(self.indexs[i] + ",")

        f.write("\n")

        for site in self.newMiningSites:
            for i, data in enumerate(site.raw):
                if i+1 not in toRemove:
                    if type(data) != list:
                        f.write(f"{data},")
                    else:
                        f.write('"')
                        for num in data[:-1]:
                            f.write(f"{num},")
                        try:
                            f.write(f"{data[-1]}\",")
                        except IndexError:
                            f.write("\",")

            f.write("\n")

        print(location + ".csv" + f" Has been created at {os.getcwd()}\\{f.name}")

        f.close()

    def toExcel(self, location : str):
        import xlsxwriter

        wb = xlsxwriter.Workbook(location + ".xlsx")

        ws = wb.add_worksheet()

        for i, header in enumerate(self.indexs):
            ws.write(0, i, header)

        for i, mine in enumerate(self.newMiningSites):
            i += 1

            for ii, data in enumerate(mine.raw):
                if type(data) != list: 
                    ws.write(i, ii, data)
                
                else:
                    ws.write(i, ii, ",".join(str(x) for x in data))
        
        ws.conditional_format(f"A1:AA{i}", {"type" : "text", "criteria" : "containing", "value" : "foo"})

        print(location + ".xlsx" + f" Has been created at {os.getcwd()}\\{location + '.xlsx'}")

        wb.close()

    def fourFileOutput(self, name : str, toRemove : list[int] = [], sep = " "):

        blocks = open(f"{name}.blocks", "w")
        delay = open(f"{name}.delay", "w")
        fpp = open(f"{name}.fpp", "w")
        on = open(f"{name}.old", "w")

        toRemove.append(0)

        for mine in self.newMiningSites:
            
            for b in range(0, 21):
                if b not in toRemove:
                    blocks.write(f"{mine.raw[b]}{sep}")

            blocks.write("\n")

            delay.write(f"{mine.raw[1]}{sep}")
            delay.write(f'{mine.raw[-4]}{sep}')
            delay.write(" ".join(str(x) for x in mine.raw[-2]) + "\n")

            fpp.write(f"{mine.raw[1]}{sep}")
            fpp.write(f'{mine.raw[-4]}{sep}')
            fpp.write(" ".join(str(x) for x in mine.raw[-3]) + "\n")

            on.write(f"{mine.raw[1]}{sep}{mine.raw[-1]}\n")

        print(location + ".blocks" + f" Has been created at {os.getcwd()}\\{blocks.name}")
        print(location + ".delay" + f" Has been created at {os.getcwd()}\\{delay.name}")
        print(location + ".fpp" + f" Has been created at {os.getcwd()}\\{fpp.name}")
        print(location + ".on" + f" Has been created at {os.getcwd()}\\{on.name}")

        blocks.close()
        delay.close()
        fpp.close()
        on.close()


#a = h(r"C:\Users\olive\OneDrive\code\andreaBrickey\myCode\scriptinput - new.csv")

#a.genNewSites()

#a.fourFileOutput(r"testFourFile")

def openFile() -> str:
    try:
        toReadFunc = lambda : input("What is the location of the file you to use for the input data? (Must be a CSV) This is case sensitive! (Can be the full file path or just the name if its in the same direcotry as this script file, do include the file extension.)\n\n")
        toRead = toReadFunc()

        f = open(toRead, "r")

        f.close()

        return toRead
    
    except FileNotFoundError:
        print("That file does not exist or cannot be opened, please try again\n\n")
        return openFile()

def siteRemoval(location : str):
    response = input("""What mining sites do you want to remove from the input file? 
    You can select mining sites by specifying a pit number, specifying an OMP number, or specifying a range of OMP numbers.
    (1..5) would remove OMP 1 through 5 inclusive.
    You may select multiple OMP and pit numbers by seperating them with a space.
    
    """)

    tr = set()

    for remove in response.split(" "):
        if ".." in remove:
            for omp in range(int(remove.split("..")[0]), int(remove.split("..")[1]) + 1):
                tr.add(omp)
        else:
            try:
                tr.add(int(remove))
            except ValueError:
                tr.add(remove)

    miningSite = h(location)
    #miningSite.genNewSites()

    newSites = set()

    for i, site in enumerate(miningSite.miningSites):
        if site.raw[1] in tr or site.raw[0] in tr:
            pass
            #miningSite.miningSites.pop(i)
            #print(site.raw[-1] in tr, site.raw[0] in tr, i, site.raw)
        else:
            newSites.add(site.raw[1])

    for x in range(len(miningSite.miningSites)-1, -1, -1):
        if miningSite.miningSites[x].raw[1] not in newSites:
            miningSite.miningSites.pop(x)

    missingPreds = []

    for i, site in enumerate(miningSite.miningSites):
        for index in range(len(site.raw[-3]) - 1, -1, -1):
            pred = site.raw[-3][index]
            if pred not in newSites:
                #print(index, pred)
                miningSite.miningSites[i].raw[-3].pop(index)
                #print(site.raw, pred, site.raw[-4], site.raw[1])
                missingPreds.append(f'Precedence number {pred} for OMP number {site.raw[1]}')

    print(', '.join(missingPreds))
    
    response = input("Is it ok to remove the precedence's for these OMP activitys? y/n\n\n")

    if response.lower() == "y":
        miningSite.genNewSites()
        return miningSite
    return siteRemoval(location)

        

if __name__ == "__main__":

    try:

        os.chdir(sys.argv[0][:sys.argv[0].rfind("\\")])

    except FileNotFoundError:

        try:
            os.chdir(sys.argv[0][:sys.argv[0].rfind("/")])
        except FileNotFoundError:
            pass
    

    if len(sys.argv) == 1:
        outputLoc = input("What do you want the names of your files to be? Do not include the file extension!\n")

        while (input(f"\nAre you sure you want \"{outputLoc}\" to be your file names?\nEnter \"y\" to confirm: ").lower() != "y"):
            outputLoc = input("What do you want the names of your files to be?\n")

        validArgs = ['all', '1', 'csv', '2', 'excel', '3', 'omp', '4']

        toPerformFunc = lambda : input("""What files do you want to generate?
If you want to generate all files (An excel spreadsheet, a csv file, and the four OMP files) enter \"all\" or 1
If you to generate a csv file enter \"csv\" or 2
If you want to generate a Excel file enter \"Excel\" or 3
If you want to generate the OMP files enter \"OMP\" or 4\n\n""").lower()

        toPerform = toPerformFunc()

        while toPerform not in validArgs:
            
            toPerform = toPerformFunc()

        location = openFile()

        miningSite = siteRemoval(location)

        location = outputLoc

        toRemoveFunc = lambda : input("""Do you want to exclude any columns from the output? 
You can select columns individually or within a range (inclsive) through the syntax A..C, this would remove the columns A, B, and C.
You can include multiple selections by seperating them by a space, EX: A B..C This would remove columns A through C.
If you do not wish to remove any columns just hit enter.\n\n""")

        toRemove = toRemoveFunc()

        tr = []

        for remove in toRemove.split(" "):
            tmp = remove.split("..") if remove.split("..")[0] != "" else []

            if len(tmp) > 1:
                for num in range(miningSite.col2num(tmp[0]), miningSite.col2num(tmp[1]) + 1):
                    tr.append(num)
            elif len(tmp) == 1:
                tr.append(miningSite.col2num(tmp[0]))

        while (input(f"Are the columns you want to remove : {tr if len(tr) != 0 else 'None'}, enter \"y\" to confirm these are the columns you want to remove, if this is not accurate hit enter and re-enter the columns you want to remove: ")).lower() != "y":
            toRemove = toRemoveFunc()

            tr = []

            for remove in toRemove.split(" "):
                tmp = remove.split("..") if remove.split("..")[0] != "" else []

                if len(tmp) > 1:
                    for num in range(miningSite.col2num(tmp[0]), miningSite.col2num(tmp[1]) + 1):
                        tr.append(num)
                elif len(tmp) == 1:
                    tr.append(miningSite.col2num(tmp[0]))

        

        print("Your files will be output here: ", os.getcwd())

        if toPerform == "all" or toPerform == "1":
            miningSite.toExcel(location)
            miningSite.fourFileOutput(location, tr)
            miningSite.output(location, tr)

        elif toPerform == "csv" or toPerform == "2":
            miningSite.output(location = location, toRemove = tr)

        elif toPerform == "excel" or toPerform == "3":
            miningSite.toExcel(location)

        elif toPerform == "omp" or toPerform == "4":
            miningSite.fourFileOutput(location, tr)
        
        else:
            print("Uh Oh, something went wrong relaunch the program")

        input("Hit any key to close this window.")

    """
    elif len(sys.argv) == 2:

        location = sys.argv[1]

    
    elif len(sys.argv) <= 4:
        #run all on given file
        args = [sys.argv[0], sys.argv[1], ['--all']]

    else:
        #run based on cla
        validArgs = ['--all', '--csv', "--excel", "--ff"]

        args = [sys.argv[0], sys.argv[1]]

        for arg in sys.argv[2:]:
            if arg.lower() in validArgs:
                args.append([])
            
            args[-1].append(arg)
        
    miningSites = h(args[1])
    miningSites.genNewSites()

    for arg in args[2:]:
        if arg[0] == "--all":
            try:
                arg[1]
            except:
                arg[1] = input("What do you want the names of your files to be?")

            miningSites.toExcel(arg[1])
            miningSites.fourFileOutput(arg[1])
            miningSites.output(arg[1])
    """