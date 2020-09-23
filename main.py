
import xlrd
import xlwt
import os
import random
from xlwt import Workbook

# initialize excel sheet to read from
p1 = "~/PycharmProjects/weekly-log-4/chessgamesdata.xls"
path1 = os.path.expanduser(p1)
wb1 = xlrd.open_workbook(path1)
sheet = wb1.sheet_by_index(0)

# global variable for number of rows of data
rows = sheet.nrows - 1

# initialize excel sheet to write to
wb2 = Workbook(encoding="utf-8")
newSheet = wb2.add_sheet("Sheet 2")


# returns the average value of a column's elements
def columnAvg(col):

    vals = []

    for i in range(rows):
        # try-except to account for the first row being the label
        try:
            int(sheet.cell_value(i, col))
            vals.append(int(sheet.cell_value(i, col)))
        except ValueError:
            vals.append(sheet.cell_value(i, col))

    avg = sum(vals[1:]) / rows

    return avg


# global variable for average turn count to increase efficiency
avgTurnCount = columnAvg(2)


# returns number of occurrences of the target value in the given column
def countOf(col, target):

    count = 0

    for i in range(rows):
        if sheet.cell_value(i, col) == target:
            count += 1

    return count


# returns the probability of a longer than average game given that it's a draw
def probLong_Draw():

    # separate counters
    drawCount = 0
    longCount = 0

    for i in range(rows):
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
    Pdraw = (float(count) / rows)
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

    for i in range(rows):
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

    ans = (float(highWins) / rows) * 100
    return ans


# returns the probability of resigning given that the losing player is rated lower
# Baye's variables: X = resigning, Y = being rated lower than the opponent
def probResign_Low():

    # variables to count resigned game count vs games where the lower rated player resigned
    resignCount = 0
    lowResignCount = 0

    # calculate probability of being the lower player given that someone resigned
    for i in range(rows):
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
    resignRate = float(resignCount) / rows

    # answer for Baye's equation. (the probability of being rated lower than your opponent P(Y) is assumed to be .5)
    ans = float(probLowResign * resignRate) / .5
    return ans * 100


# returns the probability of winning given that the white player opened with the queen's gambit
def probWinQG():

    QGwins = 0
    whiteWins = 0
    QGcount = 0

    # calculate probability of white having opened with the Queen's Gambit given that they won
    for i in range(rows):
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
    for i in range(rows):
        # if the opening position includes 'Queen's Gambit' (accounts for variations of the QG)
        if "Queen's Gambit" in sheet.cell_value(i, 9):
            QGcount += 1

    probQG = float(QGcount) / rows

    # answer for Baye's equation.
    ans = float(probQGwin * probWin) / probQG

    return ans * 100


# A random sampling of the data set, sample size determined by input percent (per)
def randomSampling(per):
    # sample size = desired percent * number of rows
    sampsize = per * rows-1
    ranSample = []


    for i in range(int(sampsize)):
        print(str(i))
        ran = random.randint(1, rows)
        ranRow = sheet.row_slice(ran)
        for j in range(sheet.ncols):
            newSheet.write(i, j, str(ranRow[j].value))



print("\nBayes probabilities based on " + str(rows) + " games of chess:\n\n"
      "Chances of a draw when a game lasts longer than the average (59 turns): %.2f percent\n" % probOfDraw_Long() +
      "Chances of the higher rated player winning: %.2f percent\n" % highRateWins() +
      "Chances of the lower rated player resigning: %.2f percent\n" % probResign_Low() +
      "Chances of the white player winning after they opened with the Queen's Gambit: %.2f percent\n" % probWinQG())

randomSampling(1)


wb2.save("chessgamesTEST.xls")
