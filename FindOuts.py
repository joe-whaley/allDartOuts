import numpy as np
import pandas as pd
import openpyxl
import time

# Functions:
# estimateDPerm: Estimates the number of total permutations to finish an n-Point checkout in any number of darts
# recursiveCheck: Will return a list of checkout options (curSols) for a given starting nPoint in less than nDarts
# buildDFs: Returns a DataFrame filled with all curRowSols for the current nPoint
# getSols: Find all the checkouts for nPoints using nDarts or less


excelPath = 'DartThrows.xlsx'

class FindOuts():
    def __init__(self, path=excelPath) -> None:
        self.path = path
        self.dartTable = pd.read_excel(path)
        self.fullTable = self.dartTable.values
        self.table = self.fullTable[1:self.fullTable.shape[0],
                                    1:self.fullTable.shape[1]]

        self.startingPoints = self.fullTable[1:self.fullTable.shape[0], 0]
        self.throwsMnemonic = self.dartTable.columns[1:self.dartTable.shape[1]]

    # estimateDPerm: Estimates the number of total permutations to finish an n-Point checkout in any number of darts
    def estimateDPerm(self, nPoint, fullList=False):
        # The following equation is my conjecture to estimate the number of permutations to finish an n-point checkout
        # D(n) = D(n-1) + summation(2*D(m), from m = 2, to m = n-2) + EO for n > 3
        # See that D(n<4) are trivial solutions: D(0) = DNE, D(1) = DNE, D(2) = 1, D(3) = 1
        # Where D(n) would return the number of permutations to finish an n-Point checkout
        # Where n is the n-Points remaining
        # Where EO is equal to 1 if n is even, and is equal to 0 if n is odd for n<41 or n==50

        # Inputs:
        # nPoint - Calculate the number of permutations up-to and including nPoint
        # fullList - Return number of permutations for n = 0 to nPoint if true, or only nPoint if false

        # Output:
        # D - The number of permutations

        D = [0]  # D(1): There are no 1-point checkouts since you must end on a double
        curPoint = 1 # Start at n = 1 since the number of n-Point checkouts is based on number of n-1, n-2, n-etc. checkouts before it
        for i in range(1, nPoint): # Iterate up to nPoint
            curPoint = curPoint + 1
            D.append(D[i-1])
            if curPoint > 3: # Execute summation if n > 3
                for j in range(1, curPoint-2):
                    D[i] = D[i] + 2*D[j]
            if all([curPoint % 2 == 0, any([curPoint<41, curPoint==50])]): # Execute EO function if n is even and n<41 or n==50
                D[i] = D[i] + 1

        if fullList:
            return D # Return estimated number of permutations from 1 to nPoints
        else:
            return D[-1] # Return only the estimated number of permutations for nPoints

    # recursiveCheck: Will return a list of checkout options (curSols) for a given starting nPoint in less than nDarts
    def recursiveCheck(self, nPointRow, nextDart, nDarts, lastThrow):
        # Inputs:
        # nPointRow - The row from our table associated with the number of remaining points
        # nextDart - The number of darts you've already thrown, plus one more
        # nDarts - The maximum number of darts (or throws) allowed to check out nPoint

        # Output:
        # curSols - An array of strings that are checkout solutions for nPoint

        col = 0
        curSols = []
        while nextDart < nDarts + 1: # While our nextDart is less than the total number of darts (or throws) allowed
            for col in range(nPointRow.size):
                curThrow = self.throwsMnemonic[col] # The Mnemonic of the current throw (ex. S19 = Single 19)
                remainingPoints = nPointRow[col] # The remaining number of points after executing your curThrow
                if remainingPoints == 0: # If you have 0 points left, then you checked out and found a solution
                    curSolRun = [lThrow + ' ' + curThrow for lThrow in lastThrow]
                    curSols.extend(curSolRun)
                elif remainingPoints == 999: # If you have 999 points left, then you busted and should skip this curThrow
                    pass
                else: # If you have some other number remaining, call recursiveCheck again on the row for the remaining points
                    nextIdx = np.where(self.startingPoints == remainingPoints)[0][0]
                    nextRow = self.table[nextIdx]
                    nextThrow = self.recursiveCheck(nextRow, nextDart+1, nDarts, [curThrow])
                    curSolRun = [lThrow + ' ' + throw for throw in nextThrow for lThrow in lastThrow]
                    curSols.extend(curSolRun)
            break

        return curSols

    # buildDFs: Returns a DataFrame filled with all curRowSols for the current nPoint
    def buildDFs(self, curSols, nPoint):
        # Inputs:
        # curSols - An array of strings that are checkout solutions for nPoint
        # nPoint - The remaining points associated with the curSols checkout solutions

        # Output:
        # solDFs - A DataFrame with columns for each solution

        numSol = 1
        for sol in curSols:
            colSol = '{}{}'.format('Solution ', numSol)
            curDF = pd.DataFrame({'Remaining Points': [nPoint], colSol: sol})
            if numSol < 2:
                solDFs = curDF # Initialize solDFs if this is first loop
            else:
                solDFs = solDFs.merge(curDF) # Else, merge the curDF with all solDFs

            numSol = numSol + 1

        return solDFs
    
    # getSols: Find all the checkouts for nPoints using nDarts or less
    def getSols(self, nPoints=range(2,6), nDarts = 10001):
        # Inputs:
        # nPoints - The points you are interested in finding the checkout solutions for
        # nDarts - The number of darts (or less) you want to checkout the nPoints with

        # Output:
        # allSols - A DataFrame with all checkout solutions for each nPoints

        allSols = pd.DataFrame()
        for nPoint in nPoints:
            nIdx = np.where(self.startingPoints == nPoint)[0][0]
            nPointRow = self.table[nIdx]
            lastThrow = ['']
            curSols = self.recursiveCheck(nPointRow, 1, nDarts, lastThrow)
            solDFs = self.buildDFs(curSols, nPoint)
            allSols = pd.concat([allSols, solDFs])

        return allSols
    

if __name__ == '__main__':
    app = FindOuts()
    n = 501
    print(f'Estimated Permutations to checkout {n} Points: {app.estimateDPerm(n)}')
    tic = time.perf_counter()
    b = app.getSols(range(2,11),3)
    toc = time.perf_counter()
    b.to_excel('answer.xlsx')
    print(f'Elapsed Time: {toc-tic:0.4f} seconds')
