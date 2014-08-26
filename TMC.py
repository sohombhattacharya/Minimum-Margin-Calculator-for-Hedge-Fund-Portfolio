import numpy
from numpy import *
from scipy import optimize as op
import xlrd


def main():
    
    #docName = raw_input("Enter the location of the portfolio workbook")
    docName = "c:\users\sohom\documents\Book1.xlsx"
    portfolioExcelDoc = xlrd.open_workbook(docName, 'r')
    worksheet = portfolioExcelDoc.sheet_by_name("Sheet2")
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = -1
    groupsAndTradeInfo = []
    groupsAndTradeCountsNames = []
    saveName = ''
    groupsDict = {}

    
    while curr_row < num_rows:
        curr_row+=1
        if curr_row != 0:
            tradeInfo = []
#            row = worksheet.row(curr_row)
            curr_cell = -1
            while curr_cell < num_cells:
                curr_cell += 1
#                cell_type = worksheet.cell_type(curr_row, curr_cell)
                cell_value = worksheet.cell_value(curr_row, curr_cell)
                if curr_cell < 2:
                    tradeInfo.append(str(cell_value))
                if curr_cell >= 2:
                    tradeInfo.append(float(cell_value))
            groupsAndTradeInfo.append(tradeInfo)    
    
    print groupsAndTradeInfo
   
    for x in xrange(0, len(groupsAndTradeInfo)):
        dictGroupName = groupsAndTradeInfo[x][0]
        tradeName = groupsAndTradeInfo[x][1]
        if dictGroupName not in groupsDict: 
            groupsDict[dictGroupName] = {}
            groupsDict[dictGroupName][tradeName] = []
            for i in xrange(2, len(groupsAndTradeInfo[x])):
                groupsDict[dictGroupName][tradeName].append(groupsAndTradeInfo[x][i])
        else:
            if tradeName not in groupsDict[dictGroupName]:
                groupsDict[dictGroupName][tradeName] = []
                for i in xrange(2, len(groupsAndTradeInfo[x])):
                    groupsDict[dictGroupName][tradeName].append(groupsAndTradeInfo[x][i])
            
            
        
    print groupsDict
    
    
    for x in xrange(0, len(groupsAndTradeInfo)-1):
        testGroupName = groupsAndTradeInfo[x][0]
        if testGroupName != saveName:
            saveName = testGroupName
            groupCount = 0
            groupAndTradeNum = []
            tradeNames = []
            for y in xrange(0, len(groupsAndTradeInfo)):
                if testGroupName == groupsAndTradeInfo[y][0]:
                    groupCount+=1
                    tradeNames.append(groupsAndTradeInfo[y][1])
            groupAndTradeNum.append(testGroupName)
            groupAndTradeNum.append(groupCount)
            groupAndTradeNum.append(tradeNames)
            groupsAndTradeCountsNames.append(groupAndTradeNum)
    
    print groupsAndTradeCountsNames
    
    
    for x in xrange(0, len(groupsAndTradeCountsNames)):
        numTrades = groupsAndTradeCountsNames[x][1]
        groupName = groupsAndTradeCountsNames[x][0]
        groupTrades = groupsAndTradeCountsNames[x][2]
        saveLastConstraintNum = 0
        constraintDict = {}
        constraintFuncs = []
        
        print groupName
        print "-----------------"
     
        def obj(k):
            answer1 = 0
            answer2 = 0
            finalAnswer = 0
            
            
            for b in xrange(0, numTrades):
                #counter = 0
                for c in xrange(0, numTrades):
                    if b != c:
                        trade = groupTrades[b]
                        #counter+=1
                        tradePrice = float(groupsDict[groupName][trade][1])
                        answer1 += k[b * numTrades  + c ]*tradePrice
            answer1 *= 0.05
            
            for e in xrange(0, numTrades):
                counter = 0
                for f in xrange(0, numTrades):
                    if e == f:
                        trade = groupTrades[counter]
                        counter+=1
                        tradePrice = float(groupsDict[groupName][trade][1])
                        margin = float(groupsDict[groupName][trade][4])
                        answer2 += k[e*numTrades + c]*tradePrice*margin
                            
            finalAnswer += answer1 + answer2
            
            return finalAnswer    
        
        def makeConstr1(constrName, tradeIndex, numTrades, group, groupTrades):
            def f(k):
                answer3 = 0
                testStr=''
                tradeName = groupTrades[tradeIndex]
                for z in xrange(0, numTrades):
                    answer3 += k[tradeIndex*numTrades + z]
                    testStr += "k[" + str(tradeIndex) + "," + str(z) + "] + "
                units = groupsDict[group][tradeName][0]
                answer3 -= units
                testStr += str(-units)
                #print testStr
                return answer3
            constraintDict[constrName] = lambda k: f(k)
                
        
          
        for y in xrange(0, len(groupTrades)):
            constraintName = "constraint" + str(y+1)
            makeConstr1(constraintName, y, numTrades, groupName, groupTrades)
        saveLastConstraintNum = y+1
        
        def makeConstr2(constrName, b, c, groupTrades, group, tradeNum):
            def f(k):
                firstTradeName = groupTrades[b]
                secondTradeName = groupTrades[c]
                firstTradeHedgeRatio = groupsDict[group][firstTradeName][2]
                firstTradePrice = groupsDict[group][firstTradeName][1]
                secondTradeHedgeRatio = groupsDict[group][secondTradeName][2]
                secondTradePrice = groupsDict[group][secondTradeName][1]
                expression = k[b*tradeNum + c]*firstTradeHedgeRatio*firstTradePrice + k[c*tradeNum + b]*secondTradeHedgeRatio*secondTradePrice
                #print "k[" + str(b) + "," + str(c) + "]*" + str(firstTradeHedgeRatio)+"*"+ str(firstTradePrice) + " + k[" + str(c) + "," + str(b) + "]*" + str(secondTradeHedgeRatio) + "*" + str(secondTradePrice)
                return expression
            constraintDict[constrName] = lambda k: f(k)
            
        xyValues = []
        
        for y in xrange(0, numTrades):
            for z in xrange(y+1, numTrades):
                constraintName = "constraint" + str(saveLastConstraintNum+1)
                makeConstr2(constraintName, y, z, groupTrades, groupName, numTrades)
                saveLastConstraintNum+=1

        def makeConstr3(constrName, sameIndex, tradeNum):
            def f(k):
                return k[sameIndex*tradeNum + sameIndex]
            constraintDict[constrName] = lambda k: f(k)
        
        def makeConstr4(constrName, sameIndex, tradeNum):
            def f(k):
                return groupsDict[groupName][groupTrades[sameIndex]][3] - k[sameIndex*tradeNum + sameIndex]
            constraintDict[constrName] = lambda k: f(k)
        
        def makeConstr5(constrName, sameIndex, tradeNum):
            def f(k):
                return k[sameIndex*tradeNum + sameIndex]*(-1)
            constraintDict[constrName] = lambda k: f(k)
        
        def makeConstr6(constrName, sameIndex, tradeNum):
            def f(k):
                return k[sameIndex*tradeNum + sameIndex] - groupsDict[groupName][groupTrades[sameIndex]][3]
            constraintDict[constrName] = lambda k: f(k)
        
        
        for trade in groupTrades:
            marketValue = groupsDict[groupName][trade][3] 
            if marketValue > 0:
                constraintName = "constraint" + str(saveLastConstraintNum+1)
                makeConstr3(constraintName, groupTrades.index(trade), numTrades)
                saveLastConstraintNum+=1
                constraintName = "constraint" + str(saveLastConstraintNum+1)
                makeConstr4(constraintName, groupTrades.index(trade), numTrades)
                saveLastConstraintNum+=1
            if marketValue < 0:
                constraintName = "constraint" + str(saveLastConstraintNum+1)
                makeConstr5(constraintName, groupTrades.index(trade), numTrades)
                saveLastConstraintNum+=1
                constraintName = "constraint" + str(saveLastConstraintNum+1)
                makeConstr6(constraintName, groupTrades.index(trade), numTrades)
                saveLastConstraintNum+=1
            
        '''
        def makeConstr7(constrName, b, c, numOfTrades):
            def f(k):
                return
            
        for x in xrange(0, numTrades):
            for y in xrange(0, numTrades):
                
        '''
        for x in constraintDict:
            constraintFuncs.append(constraintDict[x])
        print "constraint funcs length", len(constraintFuncs)
        print constraintFuncs
        
        values = numpy.empty(numTrades**2)
        values.fill(0)
        
        finalAnswerForGroup = op.fmin_cobyla(obj, values, cons=constraintFuncs)
       
        
            
        '''
        trade 1 vs trade 2 
        if trade1.marketvalue*trade2.marketvalue <0 
            (equation)
        else
            =0
        '''
        
    
        
        
       
   
    
   
   
    
                
                
                        
        
    
    
    
if __name__ == '__main__':
    main()
        
    
