import math
import os
import random
import sys

import xlrd
import xlwt
from xlwt import Workbook
from operator import itemgetter

# initialize excel sheet to read from
p1 = "chessgamesdata.xls"
path1 = os.path.expanduser(p1)
wb1 = xlrd.open_workbook(path1)
sheet = wb1.sheet_by_index(0)

# global list for column names
cols = []
for j in range(sheet.ncols):
    cols.append(sheet.cell_value(0, j))

# global variable for number of rows of data
dataSize = sheet.nrows - 1


# returns the average value of a column's elements
def columnAvg(col):

    vals = []

    for i in range(1, sheet.nrows):
        # try-except to account for the first row being the label
        try:
            int(sheet.cell_value(i, col))
            vals.append(int(sheet.cell_value(i, col)))
        except ValueError:
            vals.append(sheet.cell_value(i, col))

    avg = sum(vals) / dataSize

    return avg


# global variable for average turn count to increase efficiency
avgTurnCount = columnAvg(2)


# returns number of occurrences of the target value in the given column
def countOf(col, target):

    count = 0

    for i in range(1, sheet.nrows):
        if sheet.cell_value(i, col) == target:
            count += 1

    return count


# returns the probability of a longer than average game given that it's a draw
def probLong_Draw():

    # separate counters
    drawCount = 0
    longCount = 0

    for i in range(1, sheet.nrows):
        # first check if game is a draw
        if sheet.cell_value(i, 3) == "draw":
            # increment draw count
            drawCount += 1
            # if draw, then check if game is long
            if sheet.cell_value(i, 2) > avgTurnCount:
                longCount += 1

    prob = float(longCount) / drawCount
    return prob


# returns the probability of a game drawing given that it's a longer than average game
def probOfDraw_Long():

    count = countOf(3, "draw")      # num of draws
    # probability of a game drawing is equal to the # of draws divided by the # of entries (excluding the first row)
    Pdraw = (float(count) / dataSize)
    # probability of a game going more than the average turn count = 50%
    Plong = .5
    # probability of a game being longer than average given that it's a draw
    Plongdraw = probLong_Draw()

    # Bayes equation
    ans = (Plongdraw * Pdraw) / Plong

    return ans * 100


# returns the probability of the higher rated player being the victor
#   probability of winning (P(x)) and probability of being rated higher than your opponent P(y) are both assumed
#   to be 50% for the purposes of this program. Therefore they cancel each other out and are unnecessary
#   for this Baye's calculation.
def highRateWins():

    highWins = 0

    for i in range(1, sheet.nrows):
        # check if white's rating is higher than black's rating
        if sheet.cell_value(i, 5) > sheet.cell_value(i, 6):
            # if white is rated higher and wins
            if sheet.cell_value(i, 4) == "white":
                # increment the higher rated wins count
                highWins += 1
        # else check if black is rated higher than white
        elif sheet.cell_value(i, 6) > sheet.cell_value(i, 5):
            if sheet.cell_value(i, 4) == "black":
                # increment the higher rated wins count
                highWins += 1
        # if ratings are equal, do nothing
        else:
            continue

    ans = (float(highWins) / dataSize) * 100
    return ans


# returns the probability of resigning given that the losing player is rated lower
# Baye's variables: X = resigning, Y = being rated lower than the opponent
def probResign_Low():

    # variables to count resigned game count vs games where the lower rated player resigned
    resignCount = 0
    lowResignCount = 0

    # calculate probability of being the lower player given that someone resigned
    for i in range(1, sheet.nrows):
        # check if game results in a resign
        if sheet.cell_value(i, 3) == "resign":
            resignCount += 1
            # if yes, check if lower rated player lost
            # if white is rated lower and black wins
            if (sheet.cell_value(i, 5) < sheet.cell_value(i, 6)) and (sheet.cell_value(i, 4) == "black"):
                # then the lower rated player resigned
                lowResignCount += 1
            # else if black is rated lower and white wins
            elif (sheet.cell_value(i, 5) > sheet.cell_value(i, 6)) and (sheet.cell_value(i, 4) == "white"):
                # then the lower rated player resigned
                lowResignCount += 1

    # low player given that game ends as resign
    probLowResign = float(lowResignCount) / resignCount

    # % of games that end in resign = total count of resign games / total games
    resignRate = float(resignCount) / dataSize

    # answer for Baye's equation. (the probability of being rated lower than your opponent P(Y) is assumed to be .5)
    ans = float(probLowResign * resignRate) / .5
    return ans * 100


# returns the probability of winning given that the white player opened with the queen's gambit
def probWinQG():

    QGwins = 0
    whiteWins = 0
    QGcount = 0

    # calculate probability of white having opened with the Queen's Gambit given that they won
    for i in range(1, sheet.nrows):
        # if white won add to white wins
        if sheet.cell_value(i, 4) == "white":
            whiteWins += 1
            # if the opening position includes 'Queen's Gambit' (accounts for variations of the QG)
            if "Queen's Gambit" in sheet.cell_value(i, 9):
                QGwins += 1

    # percent of white wins that started with Queen's Gambit
    probQGwin = float(QGwins) / whiteWins

    # probability of winning any game assumed to be .5
    probWin = .5

    # percent of Queen's Gambit being used
    for i in range(1, sheet.nrows):
        # if the opening position includes 'Queen's Gambit' (accounts for variations of the QG)
        if "Queen's Gambit" in sheet.cell_value(i, 9):
            QGcount += 1

    probQG = float(QGcount) / dataSize

    # answer for Baye's equation.
    ans = float(probQGwin * probWin) / probQG

    return ans * 100


# function to print out Bayes values calculated by other functions
def bayesCalcs():
    print("\nBayes probabilities based on " + str(dataSize) + " games of chess:\n\n"
          "Chances of a draw when a game lasts longer than the average (59 turns): %.2f percent\n" % probOfDraw_Long() +
          "Chances of the higher rated player winning: %.2f percent\n" % highRateWins() +
          "Chances of the lower rated player resigning: %.2f percent\n" % probResign_Low() +
          "Chances of the white player winning after they opened with the Queen's Gambit: %.2f percent\n" % probWinQG())


# A random sampling of the data set. sample size determined by input percent (per)
def randomSampling(per):
    # initialize excel sheet to write to
    wb2 = Workbook(encoding="utf-8")
    newSheet = wb2.add_sheet("Random Sampling")
    # sample size = desired percent * number of rows
    sampsize = per * dataSize

    for i in range(int(sampsize)):
        ran = random.randint(1, sheet.nrows)
        ranRow = sheet.row_values(ran)
        for j in range(sheet.ncols):
            newSheet.write(i, j, str(ranRow[j]))

    wb2.save("chessgamesRandomSample.xls")
    print("Binning complete. Exported data saved to 'chessgamesRandomSample.xls'.")


# top sampling of the data set. Sample size determined by input percent(per). Target attribute determined by column(col)
def topSampling(per, col):

    # initialize excel sheet to write to
    wb2 = Workbook(encoding="utf-8")
    newSheet = wb2.add_sheet("Equal Frequency Binning")

    # row data stored as list
    rowList = []
    for i in range(1, sheet.nrows):
        rowList.append(sheet.row_values(i))

    # sort data by target attribute column (reverse true so highest value at top)
    rowList = sorted(rowList, key=itemgetter(col), reverse=1)

    # sample size = desired percent * number of rows
    sampsize = per * dataSize

    for i in range(int(sampsize)):
        for j in range(sheet.ncols):
            newSheet.write(i, j, str(rowList[i][j]))

    wb2.save("chessgamesTopSample.xls")
    print("Top Sample complete. Exported data saved to 'chessgamesTopSample.xls'.")


# function to equal frequency binning on data set. col = feature attribute column. b = number of bins
# bins displayed by alternating excel cell background color between withe and gray
def equalFreqBinning(col, b):

    # initialize excel sheet to write to
    wb2 = Workbook(encoding="utf-8")
    newSheet = wb2.add_sheet("Equal Frequency Binning")

    # width of each bin
    binWidth = dataSize / b

    # row data stored as list
    rowList = []


    for i in range(1, sheet.nrows):
        rowList.append(sheet.row_values(i))

    rowList = sorted(rowList, key=itemgetter(col))


    gray = xlwt.easyxf('pattern: pattern solid, fore_colour gray40;')
    white = xlwt.easyxf('pattern: pattern solid, fore_colour white;')
    style = True

    for i in range(len(rowList)):
        if i % binWidth == 0:
            style = not style
        for j in range(sheet.ncols):
            if style:
                newSheet.write(i, j, str(rowList[i][j]), white)
            else:
                newSheet.write(i, j, str(rowList[i][j]), gray)
    wb2.save("chessgamesEqualFreqBinning.xls")
    print("Binning complete. Exported data saved to 'chessgamesEqualFreqBinning.xls'.")


# calculates entropy of a target column of the data set
def entropyCol(col):
    itemList = []           # list for column items
    uniqueItemList = []     # list for unique items
    uniqueFreqs = []        # list for frequencies of unique items
    totalEntropy = 0        # entropy

    # loop to add column values to item list (starts at 1 to exclude column label)
    for i in range(1, sheet.nrows):
        item = sheet.cell_value(i, col)
        # check to see if item is already in list.
        if item not in itemList:
            # if not, add 1 to unique count
           uniqueItemList.append(item)
        # add item to complete item list
        itemList.append(item)

    # loop through each unique item
    for i in range(len(uniqueItemList)):
        count = 0
        # check number of occurences of each unique item in the complete item list
        for j in range(len(itemList)):
            # when there is an occurence, add 1 to count
            if uniqueItemList[i] == itemList[j]:
                count += 1
        # unique freq holds the unique item occurence counts
        uniqueFreqs.append(count)

    # calculate entropy by summing individual probabilities of unique item occurences / data size
    for i in range(len(uniqueFreqs)):
        totalEntropy += entropyCalc(float(uniqueFreqs[i]) / len(itemList))

    print("Number of unique items in column '" + cols[col] + "': " + str(len(uniqueItemList)))
    print("Entropy of column '" + cols[col] + "' is: " + str(totalEntropy))


# calculates and returns entropy of a given value
def entropyCalc(val):
    entropy = 0
    entropy += -(val * math.log(val, 2))
    return entropy


# starting function for menu options
def start():
    print("\nThis is a program that performs various operations on the dataset 'chessgamesdata.xls'")
    print("Choose one of the following operations to perform on the dataset:\n"
          "1. Calculate various Bayes' Probabilities on the data\n"
          "2. Calculate the Shannon Entropy of a target column\n"
          "3. Perform an Equal-Frequency Binning operation to be exported as a new Excel sheet\n"
          "4. Perform a sampling of the dataset to be exported as a new Excel sheet")

    # original choice of operation
    while True:
        try:
            ans = int(input("Enter the number of your choice: "))
            if ans not in (1, 2, 3, 4):
                raise ValueError
        except ValueError:
            print("Error: choice must be an integer between 1 - 4")
        except NameError:
            print("Error: choice must be an integer between 1 - 4")
        else:
            break

    if ans == 1:
        bayesCalcs()

    elif ans == 2:
        print("Columns:")
        for i in range(len(cols)):
            print(str(i) + ". " + str(cols[i]))
        while True:
            try:
                col = int(input("Enter the column number to calculate the entropy of: "))
                if col not in range(len(cols)):
                    raise NameError
            except NameError:
                print("Error: choice must be an integer between 0 and " + str(len(cols) - 1))
            else:
                break
        entropyCol(col)

    elif ans == 3:
        print("Columns:")
        for i in range(len(cols)):
            print(str(i) + ". " + str(cols[i]))
        while True:
            try:
                col = int(input("Enter the column number to sort by for binning: "))
                if col not in range(len(cols)):
                    raise NameError
            except NameError:
                print("Error: choice must be an integer between 0 and " + str(len(cols) - 1))
            else:
                break
        while True:
            try:
                bins = int(input("Enter the number of bins you want: "))
                if bins < 1:
                    raise ValueError
            except ValueError:
                print("Error: Number of bins must be a positive integer")
            except NameError:
                print("Error: Number of bins must be an integer")
            else:
                break

        equalFreqBinning(col, bins)
    else:
        print("Choose Sampling Method:\n"
              "1. Random Sampling\n"
              "2. Top Sampling")
        while True:
            try:
                ans = int(input("Enter choice of sampling: "))
                if ans not in (1, 2):
                    raise ValueError
            except ValueError:
                print("Error: choice must be 1 or 2")
            except NameError:
                print("Error: choice must be 1 or 2")
            else:
                break
        while True:
            try:
                percent = input("Enter size of sample as a decimal percent: ")
                float(percent)
                if percent > 1 or percent < 0:
                    raise ValueError
            except ValueError:
                print("Error: percentage must be a decimal between 0 and 1")
            except NameError:
                print("Error: percentage must be a decimal between 0 and 1")
            else:
                break

        if ans == 1:
            randomSampling(percent)
        else:
            print("Columns:")
            for i in range(len(cols)):
                print(str(i) + ". " + str(cols[i]))
            while True:
                try:
                    col = int(input("Enter the target column number for Top Sampling: "))
                    if col not in range(len(cols)):
                        raise NameError
                except NameError:
                    print("Error: choice must be an integer between 0 and " + str(len(cols) - 1))
                else:
                    break
            topSampling(percent, col)
    end()


# function to quit or restart
def end():
    print("Would you like to perform another operation?")

    while True:
        try:
            # weird error with names not being defined
            # had to use raw_input to input a string
            ans = raw_input("Enter 'y' for yes or 'n' to quit: ")
            if ans not in ('y', 'Y', 'n', 'N'):
                raise ValueError
        except ValueError:
            print("Error: please enter 'y' or 'n'")
        except NameError:
            break
        else:
            break
    if ans in ('y', 'Y'):
        start()
    else:
        print("Bye!")
        sys.exit()


start()



