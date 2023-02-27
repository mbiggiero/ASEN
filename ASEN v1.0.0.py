#ASEN - v1.0.0
import networkx as nx
import pandas as pd
import numpy as np
import xlsxwriter
import warnings
import datetime
import pickle
import math
import sys
import os

#GUI Imports:
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile 
from tkinter.filedialog import askdirectory
from tkinter import messagebox
from tkinter.ttk import Progressbar

#Global settings
warnings.filterwarnings("ignore")
np.set_printoptions(threshold=np.inf)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)

##Global variables##
GM = None
isD = True
nodeData=False
nodeValues=[]
groupData=False
groupValues=[]
betaData=False
betaValues=[]
excelMatrix=[]
assortativityData=False
assortativityValues=[]
attString=False
colSize=20
matrixName = ""
directory = "."
waitForMatrix = True
totalInd=0
prog=0
phi=0.1
rp=1

m1saved=False
m2saved=False
m3saved=False

#xlsxWriter
writer=None
workbook=None
worksheetMain=None
worksheetCC = None
worksheetSCC = None
worksheetKatz= None
worksheetGxRC = None
worksheetRW= None
#xlsx Formats
main_format=None
side_format=None
mixed_format=None

##Helper Functions##
def isSquare (m): return all (len (row) == len (m) for row in m)
def isDirected():
    global GM
    directed = False
    for x,y in GM.edges():
        if GM.get_edge_data(x,y)!=GM.get_edge_data(y,x): 
            directed = True        
    return directed
def ClearMatrix(matrix):
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            matrix.iloc[i,j]=0
def hasSelfLoops():
    if nx.number_of_selfloops(GM)==0:
        return False
    else:
        return True
def isWeighted():
    isW=nx.is_weighted(GM)
    if isW:
        isReallyW=False
        for x,y in GM.edges():
            if GM.get_edge_data(x,y)['weight']!=1: 
                isReallyW=True
                break
        if isReallyW==False:
            isW=False
    return isW
#I/O Functions
def OpenFile(name): sys.stdout = open(name, "w")
def CloseFile(): sys.stdout.close()
def read_pickle(filename):
   G = pickle.load(open(filename, 'rb'))
   return G
def read_dl(filename):
    header=0
    try:            
        with open(filename) as myFile:
            for num, line in enumerate(myFile, 1):
                if ("data:") in line.lower():
                    header=num
        dl=pd.read_csv(filename,sep=" ",header=None,skiprows=header,skipinitialspace=True)
    except pd.errors.ParserError as e:        
        print("ParserError")
        return None    
    return dl
def write_xlsx(table, sheet):
    global workbook
    #table.replace(0, float("NaN")).to_excel(workbook,sheet_name=sheet, header=True)
    table.replace(0, float("NaN")).to_excel(sheet, header=True)
def write_dl(G,filename):
    G=G.to_directed()
    isW=nx.is_weighted(G)
    if isW:
        isReallyW=False
        for x,y in G.edges():
            if G.get_edge_data(x,y)['weight']!=1: 
                isReallyW=True
                break
        if isReallyW==False:
            isW=False

    with open(filename,'w') as f:
        f.write("DL n=" + str(len(G))+"\n")
        f.write("format = edgelist1\n")
        f.write("labels embedded:\n")
        f.write("data:\n\n")
        
        if isW:
            for line in nx.generate_edgelist(G,data=['weight'] ):
                f.write(str(line))
                f.write("\n")
        else:        
            for line in nx.generate_edgelist(G):
                f.write(str(line.split(' {')[0]))
                f.write("\n")
        for node in G.nodes():
            if G.degree[node]==0:
                f.write(str(node))
                f.write("\n")
def write_edgelist(G,filename):
    #XlsxWriter Start
    global worksheet
    global main_format
    global side_format
    global mixed_format

    G=G.to_directed()
    isW=nx.is_weighted(G)

    if isW:
        isReallyW=False
        for x,y in G.edges():
            if G.get_edge_data(x,y)['weight']!=1: 
                isReallyW=True
                break
        if isReallyW==False:
            isW=False

    workbook = xlsxwriter.Workbook(filename, {'nan_inf_to_errors': True})  
    worksheet = workbook.add_worksheet("Graph")   

    main_format = workbook.add_format()
    main_format.set_bold()
    main_format.set_border(2)
    side_format = workbook.add_format()    
    side_format.set_border(1)
    mixed_format = workbook.add_format()    
    mixed_format.set_bold()  
    mixed_format.set_border(1)

    #Formatting
    worksheet.set_column(2, 3, colSize) 

    #Excel code    
    worksheet.write(0,0,"Source", main_format)  
    worksheet.write(0,1,"Target", main_format) 
    if isW:
        worksheet.write(0,2,"Weight", main_format)  

    lineCounter=1
    for line in nx.generate_edgelist(G,delimiter=';!%',data=['weight'] ):
        edge=line.split(';!%')
        worksheet.write(lineCounter,0,edge[0], side_format)  
        worksheet.write(lineCounter,1,edge[1], side_format)
        if isW:
            worksheet.write(lineCounter,2,float(edge[2]), side_format)  
        lineCounter+=1
    for node in G.nodes():
        if G.degree[node]==0:
            worksheet.write(lineCounter,0,str(node), side_format)  
            lineCounter+=1

    #XlsxWriter End
    excelSaved=False
    while not excelSaved:
        try:
            workbook.close()
            excelSaved=True  
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo("Output error", "Excel File already open! Close it and click OK")##TODO eigenvector 10*n iterantions??
def OpenMatrix(filepath):
    global nodeData
    global nodeValues
    global groupData
    global groupValues
    global betaData
    global betaValues
    global excelMatrix
    global assortativityData
    global assortativityValues
    global waitForMatrix
    excelMatrix = pd.read_excel(filepath, sheet_name=0, header=0,index_col=0)
    #np.fill_diagonal(excelMatrix.values, 0)
    
    try:
        groupValues = pd.read_excel(filepath, sheet_name="group",index_col=0, header=0)
        groupValues=groupValues.fillna(0) 
        groupValues=groupValues.index.tolist()        
        groupData=True  
    except Exception as e:#(IndexError, xlrd.biffh.XLRDError)
        groupData=False

    try:
        betaValues = pd.read_excel(filepath, sheet_name="beta",index_col=0, header=0)
        betaValues=betaValues.fillna(0)  
        betaData=True  
    except Exception as e:
        betaData=False

    try:
        assortativityValues = pd.read_excel(filepath, sheet_name="assortativity",index_col=0, header=0)
        assortativityValues=assortativityValues.fillna(0)      
        assortativityData=True  
    except Exception as e:#(IndexError, xlrd.biffh.XLRDError)
        assortativityData=False

    return excelMatrix
def InitializeExcel():
    global writer    
    global main_format
    global side_format
    global mixed_format

    main_format = writer.add_format()
    main_format.set_bold()
    main_format.set_border(2)
    side_format = writer.add_format()    
    side_format.set_border(1)
    mixed_format = writer.add_format()    
    mixed_format.set_bold()  
    mixed_format.set_border(1)

def CreateGraph(filepath):
    global waitForMatrix
    global excelMatrix
    global matrixName
    global nodeData
    global nodeValues
    global groupData
    global groupValues
    global betaData
    global betaValues
    global assortativityData
    global assortativityValues
    global attString
    global metaGraph
    global GM
    global isD
    global m1saved
    global m2saved
    global m3saved

    m1saved=False
    m2saved=False
    m3saved=False
    root.update() 
    analyzeBtnText.set("Loading matrix file...") 
    xTemp=matrixName.lower()

    excelMatrix=None
    GM=None
    usingEdgeList=False
    waitForMatrix=True
    error=False

    if xTemp.endswith(".xlsx") or xTemp.endswith(".xls"):    
        excelMatrix=OpenMatrix(filepath)
        if isSquare(excelMatrix.to_numpy()): 
            excelMatrix.fillna(0, inplace=True)
            GM = nx.from_pandas_adjacency(excelMatrix, create_using=nx.DiGraph()) 
            if assortativityData:
                attList={}
                for index, row in assortativityValues.iterrows(): 
                    attList[index]=row.values[0]
                    if isinstance(row.values[0], str):
                        attString=True                        
                nx.set_node_attributes(GM, name='assortativity', values=attList)
            if groupData:
                pass
            if betaData:            
                betaValues = betaValues.T.to_dict('list')
                for x in sorted(betaValues): 
                    betaValues[x]=betaValues[x][0]                    
        else:
            excelMatrix.reset_index(inplace=True)
            try:
                excelMatrix.columns = ['from','to','weight',]
            except ValueError:
                try:
                    excelMatrix.columns = ['from','to']
                    excelMatrix['weight'] = 1
                except ValueError:
                    ShowError("Input Error","Input isn't a matrix or an edgelist!")
                    error=True
                    matrixText.set("[None]")
                    analyzeBtn.config(state=NORMAL)  
                    waitForMatrix = False    
                    error=True
                    analyzeBtnText.set("Analyze")  
                    return 
            usingEdgeList=True   
    elif matrixName.endswith(".nxg"):
        GM=read_pickle(filepath)  
        usingEdgeList=False   
    else:        
        excelMatrix=read_dl(filepath) 
        if excelMatrix is None:
            error=True
            return
        try:
            excelMatrix.columns = ['from','to','weight']
        except ValueError:
            try:
                excelMatrix.columns = ['from','to']
                excelMatrix['weight'] = 1
            except ValueError:
                ShowError("Input Error","Input isn't a matrix or an edgelist!")
                error=True
                matrixText.set("[None]")
                analyzeBtn.config(state=NORMAL)  
                waitForMatrix = False    
                error=True
                analyzeBtnText.set("Analyze")  
                return
        usingEdgeList=True
    if usingEdgeList:
        try:
            GM=nx.from_pandas_edgelist(excelMatrix, 'from','to','weight', create_using=nx.DiGraph())
        except ValueError:
            ShowError("Input Error","Input isn't a matrix or an edgelist!")
            error=True
            matrixText.set("[None]")
            analyzeBtn.config(state=NORMAL)  
            waitForMatrix = False    
            analyzeBtnText.set("Analyze") 
            return            

        edgesToRemove = []     
        nodesToRemove = []
        for x,y in GM.edges():
            if pd.isnull(y):
                edgesToRemove.append([x,y])
                nodesToRemove.append(y)
            GM.add_node(x)             
        for z in edgesToRemove:
            GM.remove_edge(z[0], z[1])         
        for z in nodesToRemove:
            try:
                GM.remove_node(z)    
            except nx.exception.NetworkXError:
                pass

    if isDirected():
        isD=True
        geo = nx.DiGraph() 
        for x in sorted(GM.nodes()):
            if assortativityData:
                geo.add_node(x,assortativity=GM.nodes[x]['assortativity'])
            else:
                geo.add_node(x)
        for x in sorted(GM.edges(data=True)):
            geo.add_edge(x[0],x[1],weight=x[2]['weight'])
        GM=geo
    else:
        isD=False
        geo = nx.Graph() 
        for x in sorted(GM.nodes()):
            if assortativityData:
                geo.add_node(x,assortativity=GM.nodes[x]['assortativity'])
            else:
                geo.add_node(x)
        for x in sorted(GM.edges(data=True)):
            geo.add_edge(x[0],x[1],weight=x[2]['weight'])
        GM=geo

    #ClearMatrix(excelMatrix)
    if error==False:        
        matrixText.set(matrixName)
        analyzeBtn.config(state=NORMAL)  
        waitForMatrix = False    
        analyzeBtnText.set("Analyze") 

##Basic Indicators
def Size():
    print("Size: %d\n" % GM.number_of_nodes())
def AbsoluteDensity():
    print("Absolute Density: %d\n" % GM.number_of_edges())
def NormalizedDensity():
    if isD:
        print("Normalized Density: {:3.2f}".format(GM.number_of_edges() / (GM.number_of_nodes() * (GM.number_of_nodes() - 1)) * 100) + "%\n")
    else:
        print("Normalized Density: {:3.2f}".format(2*GM.number_of_edges() / (GM.number_of_nodes() * (GM.number_of_nodes() - 1)) * 100 )+ "%\n")
def ValueOfNetworkLinks():
    print("Value Of Network Links: {:1.2f}".format(GM.size(weight = 'weight')) + "\n")
def DisconnectednessDegree():
    print("Disconnectedness Degree: {:3.2f}".format((nx.number_weakly_connected_components(GM.to_directed()) - 1) / (GM.number_of_nodes() * (GM.number_of_nodes() - 1))) + "\n")
def AverageDegreeCentralityBinary():
    print("Average Degree Centrality Binary: {:.3f}".format(GM.number_of_edges() / GM.number_of_nodes()) + "\n")
def AverageDegreeCentralityValued():
    print("Average Degree Centrality Valued: {:1.3f}".format(GM.size(weight = 'weight') / GM.number_of_nodes()) + "\n")
def AverageLinkWeight():
    print("Average Link Weight: {:1.3f}".format(GM.size(weight = 'weight') / GM.number_of_edges()) + "\n")
def GlobalClusteringCoefficient():
    totalSum = 0
    totalCount = 0
    G=GM.to_directed()
    nodes_nbrs = ((n, G._pred[n], G._succ[n]) for n in G.nodes())
    for i, preds, succs in nodes_nbrs:        
        ipreds = set(preds) - {i}
        isuccs = set(succs) - {i}
        sum = 0
        count = 0
        for x in isuccs.union(ipreds):
            count += 1
            for y in isuccs.union(ipreds):
                if (x != y) and (x != i) and (y != i) and G.has_edge(x, y):                        
                    sum += 1
        if (count - 1) > 0:
            localClusteringCoefficient = sum / (count * (count - 1))
            totalCount += 1
            totalSum += localClusteringCoefficient
    G=None
    print("Global Clustering Coefficient: {:1.3f}".format((totalSum / totalCount)) + "\n")
def ReciprocityG():
    print("Arc Reciprocity of the graph: {:1.3f}".format(nx.overall_reciprocity(GM)) + "\n")
def ReciprocityWeightedG():
    totLinks=GM.size(weight = 'weight')
    undirLinks=0
    for x,y in GM.edges():
        if GM.get_edge_data(x,y) is not None and GM.get_edge_data(y,x)is not None:   
            undirLinks+=abs(GM.get_edge_data(x,y)['weight']-GM.get_edge_data(y,x)['weight'])
    print("Arc Reciprocity of the weigthed graph: {:1.3f}".format(1-(undirLinks)/totLinks) + "\n")  

#Miscellaneous Indicators
def AssortativityBinary():
    print("Assortativity Binary:")
    print("\tIN-IN\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='in', y='in', weight=1)))
    print("\tIN-OUT\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='in', y='out', weight=1)))
    print("\tOUT-IN\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='out', y='in', weight=1)))
    print("\tOUT-OUT\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='out', y='out', weight=1), "\n"))
    print("")
def AssortativityValued():
    print("Assortativity Valued:")      
    print("\tIN-IN\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='in', y='in', weight='weight')))
    print("\tIN-OUT\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='in', y='out', weight='weight')))
    print("\tOUT-IN\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='out', y='in', weight='weight')))
    print("\tOUT-OUT\t\t{:0.3f}".format(nx.degree_pearson_correlation_coefficient(GM, x='out', y='out', weight='weight'), "\n"))
    print("")
def SnijderCentralization():
    print("Snijder's Centralization:")
    g = GM.number_of_nodes()
    s = GM.number_of_edges()    
    h = GM.number_of_nodes()-1
    g2= g**2
    h2= h**2
    I = (g*3+2)/4

    if isD:
        E=(s*(g*h-s)*h)/(g*g*(g*h-1))  
        vmax=0
        if (g % 2) == 0:
            vmax=h2/4
        else:
            vmax=((1-(1/g2))*h2)/4

        sumIn = 0
        for x in GM.in_degree():
            sumIn = sumIn+x[1]
        avgIn=sumIn/g    
        sumInDiff = 0
        for x1 in GM.in_degree():
            sumInDiff = sumInDiff+(x1[1]-avgIn)*(x1[1]-avgIn)
        varIn = sumInDiff/g

        sumOut = 0
        for y in GM.out_degree():
            sumOut = sumOut+y[1]
        avgOut=sumOut/g
        sumOutDiff = 0
        for y1 in GM.out_degree():
            sumOutDiff = sumOutDiff+(y1[1]-avgOut)*(y1[1]-avgOut)
        varOut = sumOutDiff/g

        print("\tJ-IN\t\t{:0.3f}".format(math.sqrt(varIn/vmax)))
        print("\tJ-OUT\t\t{:0.3f}".format(math.sqrt(varOut/vmax)))
        print("\tH-IN\t\t{:0.3f}".format((math.sqrt(varIn)-math.sqrt(E))/(math.sqrt(vmax)-math.sqrt(E))))
        print("\tH-OUT\t\t{:0.3f}".format((math.sqrt(varOut)-math.sqrt(E))/(math.sqrt(vmax)-math.sqrt(E))), "\n")
    else:
        s = s*2 
        Eun= (s*((g**2)-g-s))/(g*g*(g+1))
        vmaxUn = I*((I-1)**2)*(g-I)/(g**2)

        sumUn = 0
        for x in GM.degree():
            sumUn = sumUn+x[1]
        avgUn=sumUn/g    
        sumUnDiff = 0
        for x1 in GM.degree():
            sumUnDiff = sumUnDiff+(x1[1]-avgUn)*(x1[1]-avgUn)
        varUn = sumUnDiff/g

        print("\tJ\t\t{:0.3f}".format(math.sqrt(varUn/vmaxUn)))
        print("\tH \t\t{:0.3f}".format((math.sqrt(varUn)-math.sqrt(Eun))/(math.sqrt(vmaxUn)-math.sqrt(Eun))))
        print("")
def Triads():
    print("Triadic Census:")
    for item, amount in nx.triadic_census(GM.to_directed()).items():  # dct.iteritems() in Python 2
        print("'{}'\t: {}".format(item, amount))
    print("")

#Centralization Indicators
def InDegreeCentralityBinary():
    global worksheetMain    
    nc = GM.number_of_nodes()

    if isD:
        worksheetMain.write(0,getRP(),"IDCBA", main_format)      
        setRP()
        worksheetMain.write(0,getRP(),"IDCBN", main_format)    
    else:
        worksheetMain.write(0,getRP(),"DCBA", main_format)   
        setRP()   
        worksheetMain.write(0,getRP(),"DCBN", main_format)   

    counter=1
    temp=[]
    if isD:
        for x in nx.in_degree_centrality(GM).items():
            worksheetMain.write(counter,getRP()-1,x[1]*(nc-1), side_format)  
            worksheetMain.write(counter,getRP(),x[1], side_format)  
            temp.append(x[1])
            counter+=1  
    else:
        for x in nx.degree_centrality(GM).items():
            worksheetMain.write(counter,getRP()-1,x[1]*(nc-1), side_format)  
            worksheetMain.write(counter,getRP(),x[1], side_format)  
            temp.append(x[1])
            counter+=1  
    setRP()

    #InDegreeCentralizationBinary
    max = 0
    sum = 0
    for v in temp:
        if (max < v):
            max = v
    for v in temp:
        sum = sum + (max - v)
    den=0
    if isD:
        den=nc-1
        print("In Degree Centralization Binary Absolute: {:1.3f}".format(sum*(nc-1))  + "\n")
        print("In Degree Centralization Binary Normalized: {:1.3f}".format((sum / den )) + "\n")
    else:  
        den=nc-1  	
        print("Degree Centralization Binary Absolute: {:1.3f}".format(sum*(nc-1))  + "\n")
        print("Degree Centralization Binary Normalized: {:1.3f}".format((sum / den )) + "\n")
def OutDegreeCentralityBinary():
    global worksheetMain    
    nc = GM.number_of_nodes()

    worksheetMain.write(0,getRP(),"ODCBA", main_format) 
    setRP() 
    worksheetMain.write(0,getRP(),"ODCBN", main_format) 
    counter=1
    temp=[]    
    for x in nx.out_degree_centrality(GM).items():
        worksheetMain.write(counter,getRP()-1,x[1]*(nc-1), side_format)  
        worksheetMain.write(counter,getRP(),x[1], side_format)  
        temp.append(x[1])
        counter+=1    
    setRP() 

    #OutDegreeCentralizationBinary
    max = 0
    sum = 0
    for v in temp:
        if (max < v):
            max = v
    for v in temp:
        sum = sum + (max - v)
    den=0
    if isD:
    	den=nc-1
    else:
    	den=nc-2
    print("Out Degree Centralization Binary Absolute: {:1.3f}".format(sum*(nc-1)) + "\n")
    print("Out Degree Centralization Binary Normalized: {:1.3f}".format((sum / den )) + "\n")
def InDegreeCentralityValued():
    global worksheetMain   
    nc = GM.number_of_nodes()

    if isD:
        worksheetMain.write(0,getRP(),"IDCWA", main_format) 
        setRP()  
        worksheetMain.write(0,getRP(),"IDCWN", main_format) 
    else:
        worksheetMain.write(0,getRP(),"DCWA", main_format) 
        setRP()  
        worksheetMain.write(0,getRP(),"DCWN", main_format) 

    maxEdge = 0
    tempDistr= []
    temp=[]
    if isD:
        for k, v in GM.in_edges():
            tempDistr.append(GM.edges[k, v]['weight'])
            if (GM.edges[k, v]['weight'] > maxEdge):
                maxEdge = GM.edges[k, v]['weight']
        counter=1
        for x in GM.in_degree(weight = 'weight'):
            worksheetMain.write(counter,getRP()-1,x[1], side_format)  
            t=x[1]/((GM.number_of_nodes() - 1) * maxEdge)
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format)  
            counter+=1   
    else:
        for k, v in GM.edges():
            tempDistr.append(GM.edges[k, v]['weight'])
            if (GM.edges[k, v]['weight'] > maxEdge):
                maxEdge = GM.edges[k, v]['weight']
        counter=1
        for x in GM.degree(weight = 'weight'):
            worksheetMain.write(counter,getRP()-1,x[1], side_format)  
            t=x[1]/((GM.number_of_nodes() - 1) * maxEdge)
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format)  
            counter+=1   

    setRP()
    sum=0
    maxNorm=0
    for x in temp:
        if x > maxNorm:
            maxNorm = x  
    for y in temp:
        sum = sum + (maxNorm - y)
    den=0
    if isD:
        den=nc-1
        print("In Degree Centralization Weighted Absolute: {:1.3f}".format(sum*(nc-1)* maxEdge)  + "\n")
        print("In Degree Centralization Weighted Normalized: {:1.3f}".format((sum / den )) + "\n")
    else:
        den=nc-1
        print("Degree Centralization Weighted Absolute: {:1.3f}".format(sum*(nc-1)* maxEdge)  + "\n")
        print("Degree Centralization Weighted Normalized: {:1.3f}".format((sum / den )) + "\n")
def OutDegreeCentralityValued():
    global worksheetMain    
    nc = GM.number_of_nodes()
     
    worksheetMain.write(0,getRP(),"ODCWA", main_format) 
    setRP()     
    worksheetMain.write(0,getRP(),"ODCWN", main_format)     
    maxEdge = 0
    tempDistr= []
    for k, v in GM.in_edges():
        tempDistr.append(GM.edges[k, v]['weight'])
        if (GM.edges[k, v]['weight'] > maxEdge):
            maxEdge = GM.edges[k, v]['weight']
    counter=1
    temp=[]
    for x in GM.out_degree(weight = 'weight'):
        worksheetMain.write(counter,getRP()-1,x[1], side_format)  
        t=x[1]/((GM.number_of_nodes() - 1) * maxEdge)
        temp.append(t)
        worksheetMain.write(counter,getRP(),t, side_format)  
        counter+=1   
    setRP()   
    sum=0
    maxNorm=0
    for x in temp:
        if x > maxNorm:
            maxNorm = x 
    for y in temp:
        sum = sum + (maxNorm - y)
    den=0
    if isD:
    	den=nc-1
    else:
    	den=nc-2
    print("Out Degree Centralization Weighted Absolute: {:1.3f}".format(sum*(nc-1)* maxEdge) + "\n")
    print("Out Degree Centralization Weighted Normalized: {:1.3f}".format((sum / den )) + "\n")# * maxEdge
def BetweennessCentrality():
    global worksheetMain    
    nc = GM.number_of_nodes()

    worksheetMain.write(0,getRP(),"BCA", main_format) 
    setRP()    
    worksheetMain.write(0,getRP(),"BCN", main_format) 

    counter=1
    temp=[]
    for x in nx.betweenness_centrality(GM, normalized=False, weight=1).items():
        worksheetMain.write(counter,getRP()-1,x[1], side_format)  
        if isD:
            t=x[1]/((nc-1)*(nc-2))
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format) 
        else:
            t=x[1]*2/((nc-1)*(nc-2))
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format) 
        counter+=1   

    setRP()
    #BetweennessCentralization
    max = 0
    sum = 0
    for v in temp:
        if (max < v):
            max = v
    for v in temp:
        sum = sum + (max - v)
    print("Betweenness Centralization: {:1.3f}".format((sum / (GM.number_of_nodes() - 1))) + "\n")
def BetweennessCentralityW():
    global worksheetMain    
    nc = GM.number_of_nodes()

    worksheetMain.write(0,getRP(),"BCWA", main_format) 
    setRP()
    
    worksheetMain.write(0,getRP(),"BCWN", main_format) 

    counter=1
    temp=[]
    for x in nx.betweenness_centrality(GM, normalized=False, weight='weight').items():
        worksheetMain.write(counter,getRP()-1,x[1], side_format)  
        if isD:
            t=x[1]/((nc-1)*(nc-2))
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format) 
        else:
            t=x[1]*2/((nc-1)*(nc-2))
            temp.append(t)
            worksheetMain.write(counter,getRP(),t, side_format) 
        counter+=1   
    setRP()

    #BetweennessCentralizationW
    max = 0
    sum = 0
    for v in temp:
        if (max < v):
            max = v
    for v in temp:
        sum = sum + (max - v)
    print("Betweenness Centralization Weighted: {:1.3f}".format((sum / (GM.number_of_nodes() - 1))) + "\n")
def ClosenessCentralityIn():
    global worksheetMain
    size=GM.number_of_nodes()
    if isD:
        worksheetMain.write(0,getRP(),"CCI", main_format) 
    else:
        worksheetMain.write(0,getRP(),"CC", main_format) 
    max=0
    sum=0
    temp=[]
    counter=1
    for x in nx.closeness_centrality(GM).items():
        worksheetMain.write(counter,getRP(),x[1], side_format)  
        if x[1]>max:
            max=x[1]
        temp.append(x[1]) 
        counter+=1   
    for z in temp:
        sum+=(max-z)
    den=(size-1)
    den2=(((size-2)*(size-1))/((2*size)-3))
    if isD:
        print("Closeness Centralization In: {:1.3f}".format(sum/den) + "\n")  
    else:       
        print("Closeness Centralization: {:1.3f}".format(sum/den2) + "\n")  
    setRP()    
def ClosenessCentralityOut():
    global worksheetMain
    size=GM.number_of_nodes()  
    worksheetMain.write(0,getRP(),"CCO", main_format) 
    max=0
    sum=0
    temp=[]
    counter=1
    for x in nx.closeness_centrality(GM.reverse()).items():
        if x[1]>max:
            max=x[1]
        temp.append(x[1]) 
        worksheetMain.write(counter,getRP(),x[1], side_format)  
        counter+=1   
    for z in temp:
        sum+=(max-z)
    den=(size-1)
    den2=(((size-2)*(size-1))/((2*size)-3))
    if isD:
        print("Closeness Centralization Out: {:1.3f}".format(sum/den) + "\n")  
    else:      
        print("Closeness Centralization Out: {:1.3f}".format(sum/den2) + "\n")  
    setRP()       

#Eigenvector Indicators
def EVCentralityIn():
    global worksheetEV
    worksheetEV = writer.add_worksheet("Eigenvector") 
    worksheetEV.set_column(0, 0, colSize)
    counter=1
    for x in GM.nodes():
        worksheetEV.write(counter,0,str(x), main_format)  
        counter+=1  

    if isD:
        worksheetEV.write(0,1,"EICN", main_format)  
    else:  
        worksheetEV.write(0,1,"ECN", main_format)

    try:
        max=0
        sum=0
        temp=[]
        counter=1
        for x in nx.eigenvector_centrality(GM, max_iter=1000).items():
            worksheetEV.write(counter,1,x[1], side_format)  
            counter+=1  
            if x[1]>max:
                max=x[1]
            temp.append(x[1])
        for z in temp:
            sum+=(max-z)
        if isD:
            print("Eigenvector-In Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
        else:
            print("Eigenvector Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetEV.write(x+1,1,"N/A", side_format)  
        if isD:
            print("Eigenvector-In Centralization: N/A\n")
        else:
            print("Eigenvector Centralization: N/A\n")
def EVCentralityOut():
    global worksheetEV
    worksheetEV.write(0,2,"EOCN", main_format) 
    try:
        max=0
        sum=0
        temp=[]
        counter=1
        for x in nx.eigenvector_centrality(GM.reverse(), max_iter=1000).items():
            worksheetEV.write(counter,2,x[1], side_format)  
            counter+=1  
            if x[1]>max:
                max=x[1]
            temp.append(x[1])
        for z in temp:
            sum+=(max-z)
        print("Eigenvector-Out Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetEV.write(x+1,2,"N/A", side_format)   
        print("Eigenvector-Out Centralization: N/A\n")    
def EVCentralityInWeighted():
    global worksheetEV
    if isD:
        worksheetEV.write(0,3,"EICWN", main_format) 
    else:
        worksheetEV.write(0,2,"ECWN", main_format) 
    try:   
        max=0
        sum=0
        temp=[]
        counter=1
        if isD:
            for x in nx.eigenvector_centrality(GM,weight='weight', max_iter=1000).items():
                worksheetEV.write(counter,3,x[1], side_format)  
                counter+=1 
                if x[1]>max:
                    max=x[1]
                temp.append(x[1]) 
        else:
            for x in nx.eigenvector_centrality(GM,weight='weight', max_iter=1000).items():
                worksheetEV.write(counter,2,x[1], side_format)  
                counter+=1 
                if x[1]>max:
                    max=x[1]
                temp.append(x[1]) 

        for z in temp:
            sum+=(max-z)
        if isD:
            print("Eigenvector-In Centralization Weighted: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
        else:
            print("Eigenvector Centralization Weighted: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")

    except nx.PowerIterationFailedConvergence:
        if isD:
            for x in range(len(GM.nodes())):
                worksheetEV.write(x+1,3,"N/A", side_format)  
            print("Eigenvector-In Centralization Weighted: N/A\n")  
        else:  
            for x in range(len(GM.nodes())):
                worksheetEV.write(x+1,2,"N/A", side_format)  
            print("Eigenvector Centralization Weighted: N/A\n")  
def EVCentralityOutWeighted():
    global worksheetEV 
    worksheetEV.write(0,4,"EOCWN", main_format) 
    try:
        counter=1
        max=0
        sum=0
        temp=[]
        for x in nx.eigenvector_centrality(GM.reverse(),weight='weight', max_iter=1000).items():
            worksheetEV.write(counter,4,x[1], side_format)  
            counter+=1  
            if x[1]>max:
                max=x[1]
            temp.append(x[1])    
        for z in temp:
            sum+=(max-z)
        print("Eigenvector-Out Centralization Weighted: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetEV.write(x+1,4,"N/A", side_format)  
        print("Eigenvector-Out Centralization Weighted: N/A\n")

#Katz Indicators
def KatzCentralityIn():
    global worksheetKatz    
    global betaData
    global betaValues
    worksheetKatz = writer.add_worksheet("Katz") 
    worksheetKatz.set_column(0, 0, colSize)
    counter=1
    for x in GM.nodes():
        worksheetKatz.write(counter,0,str(x), main_format)  
        counter+=1  

    beta2=1
    if betaData:
        beta2=betaValues
    else:
        beta2=float(beta.get()) 

    worksheetKatz.write(0,1,"KICA", main_format)
    worksheetKatz.write(0,2,"KICN", main_format) 

    try:
        max=0
        sum=0
        temp=[]
        norm=[]
        counter=1
        for x in nx.katz_centrality(GM,alpha=float(alpha.get()), beta=beta2, normalized=False, max_iter=1000).items():
            worksheetKatz.write(counter,1,x[1], side_format)  
            norm.append(x[1]**2)
            counter+=1     
        newSum=0
        for y in range(len(norm)):
            newSum+=norm[y]
        normalizer=math.sqrt(newSum)
        counter=1
        for zz in range(len(norm)):  
            normValue=math.sqrt(norm[zz])/normalizer
            worksheetKatz.write(counter,2,normValue, side_format) 
            counter+=1                  
            temp.append(normValue)  
            if normValue>max:
                max=normValue
        for z in temp:
            sum+=(max-z)

        print("Katz-In Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetKatz.write(x+1,1,"N/A", side_format)  
            worksheetKatz.write(x+1,2,"N/A", side_format) 
        print("Katz-In Centralization: N/A\n")
def KatzCentralityOut():
    global worksheetKatz
    global betaData
    global betaValues
    beta2=1
    if betaData:
        beta2=betaValues
    else:
        beta2=float(beta.get())

    worksheetKatz.write(0,3,"KOCA", main_format) 
    worksheetKatz.write(0,4,"KOCN", main_format) 
 
    try: 
        max=0
        sum=0
        temp=[]
        norm=[]
        counter=1
        for x in nx.katz_centrality(GM.reverse(), alpha=float(alpha.get()), beta=beta2, normalized=False, max_iter=1000).items():
            worksheetKatz.write(counter,3,x[1], side_format)  
            norm.append(x[1]**2)  
            counter+=1  
        newSum=0
        for y in range(len(norm)):
            newSum+=norm[y]
        normalizer=math.sqrt(newSum)
        counter=1
        for zz in range(len(norm)):  
            normValue=math.sqrt(norm[zz])/normalizer
            worksheetKatz.write(counter,4,normValue, side_format) 
            counter+=1                  
            temp.append(normValue)  
            if normValue>max:
                max=normValue
        for z in temp:
            sum+=(max-z)

        print("Katz-Out Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetKatz.write(x+1,3,"N/A", side_format)  
            worksheetKatz.write(x+1,4,"N/A", side_format) 
        print("Katz-Out Centralization: N/A\n")    
def KatzCentralityInWeighted():
    global worksheetKatz
    global betaData
    global betaValues
    beta2=1
    if betaData:
        beta2=betaValues
    else:
        beta2=float(beta.get())
    worksheetKatz.write(0,5,"KICWA", main_format)
    worksheetKatz.write(0,6,"KICWN", main_format)  

    try:
        max=0
        sum=0
        temp=[]
        norm=[]
        counter=1
        for x in nx.katz_centrality(GM,weight='weight', alpha=float(alpha.get()), beta=beta2, normalized=False, max_iter=1000).items():
            worksheetKatz.write(counter,5,x[1], side_format)  
            norm.append(x[1]**2)  
            counter+=1       
        newSum=0
        for y in range(len(norm)):
            newSum+=norm[y]
        normalizer=math.sqrt(newSum)
        counter=1
        for zz in range(len(norm)):  
            normValue=math.sqrt(norm[zz])/normalizer
            worksheetKatz.write(counter,6,normValue, side_format) 
            counter+=1                  
            temp.append(normValue)  
            if normValue>max:
                max=normValue
        for z in temp:
            sum+=(max-z)

        print("Katz-In Centralization Weighted: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetKatz.write(x+1,5,"N/A", side_format)  
            worksheetKatz.write(x+1,6,"N/A", side_format)         
        print("Katz-In Centralization Weighted: N/A\n")    
def KatzCentralityOutWeighted():
    global worksheetKatz
    global betaData
    global betaValues
    beta2=1
    if betaData:
        beta2=betaValues
    else:
        beta2=float(beta.get())
    worksheetKatz.write(0,7,"KOCWA", main_format) 
    worksheetKatz.write(0,8,"KOCWN", main_format) 

    try:
        max=0
        sum=0
        temp=[]
        norm=[]
        counter=1
        for x in nx.katz_centrality(GM.reverse(),weight='weight', alpha=float(alpha.get()), beta=beta2, normalized=False, max_iter=1000).items():
            worksheetKatz.write(counter,7,x[1], side_format)  
            norm.append(x[1]**2)  
            counter+=1        
        newSum=0
        for y in range(len(norm)):
            newSum+=norm[y]
        normalizer=math.sqrt(newSum)
        counter=1
        for zz in range(len(norm)):  
            normValue=math.sqrt(norm[zz])/normalizer
            worksheetKatz.write(counter,8,normValue, side_format) 
            counter+=1                  
            temp.append(normValue)  
            if normValue>max:
                max=normValue
        for z in temp:
            sum+=(max-z)

        print("Katz-Out Centralization Weighted: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")
    except nx.PowerIterationFailedConvergence:
        for x in range(len(GM.nodes())):
            worksheetKatz.write(x+1,7,"N/A", side_format)  
            worksheetKatz.write(x+1,8,"N/A", side_format) 
        print("Katz-Out Centralization Weighted: N/A\n")

#Random Walk Indicators   
def RandomWalk():
    global worksheetRW    
    worksheetRW = writer.add_worksheet("Random walks") 
    worksheetRW.set_column(0, 0, colSize)
    counter=1
    for x in GM.nodes():
        worksheetRW.write(counter,0,str(x), main_format)  
        counter+=1     
    worksheetRW.write(0,1,"RW", main_format) 
    counter=1
    for x in nx.approximate_current_flow_betweenness_centrality(GM, normalized=False).items():
        worksheetRW.write(counter,1,x[1], side_format) 
        counter+=1      
def RandomWalkNormalized():
    global worksheetRW
    worksheetRW.write(0,2,"RWN", main_format) 
    max=0
    sum=0
    temp=[]
    counter=1
    for x in nx.approximate_current_flow_betweenness_centrality(GM, normalized=True).items():
        if x[1]>max:
            max=x[1]
        temp.append(x[1])
        worksheetRW.write(counter,2,x[1], side_format)  
        counter+=1     
    for z in temp:
        sum+=(max-z)
    print("RWB Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")    
def WeightedRandomWalk():
    global worksheetRW
    worksheetRW.write(0,3,"RWW", main_format) 
    counter=1
    for x in nx.approximate_current_flow_betweenness_centrality(GM, normalized=False, weight = 'weight').items():
        worksheetRW.write(counter,3,x[1], side_format)  
        counter+=1     
def WeightedRandomWalkNormalized():
    global worksheetRW
    worksheetRW.write(0,4,"RWWN", main_format) 
    max=0
    sum=0
    temp=[]
    counter=1
    for x in nx.approximate_current_flow_betweenness_centrality(GM, normalized=True, weight = 'weight').items():
        if x[1]>max:
            max=x[1]
        temp.append(x[1])
        worksheetRW.write(counter,4,x[1], side_format) 
        counter+=1   
    for z in temp:
        sum+=(max-z)
    print("Weighted RWB Centralization: {:1.3f}".format((sum/(len(GM.nodes())-1))) + "\n")  

def Fragmentation(shortest):
    sumU = 0
    sumW = 0
    for k, v in shortest:
        for x, y in v.items():
            #if x < k: #Remove the 2 comments for the alternative formula
            if (len(y)-1) != 0:
                sumU+=1
                sumW+=(len(y)-1)
    print("Average Distance: {:.3f}\n".format((sumW/sumU)))
    print("Fragmentation: {:1.3f}\n".format(1 - (sumU / (GM.number_of_nodes() * (GM.number_of_nodes() - 1))))) 
    if isD:
        tempG=GM.to_undirected()
        sumP=0
        for x in GM.nodes():
            for y in GM.nodes():
                if x!=y and nx.has_path(tempG,x,y):
                    sumP+=1
        tempG=None
        print("Fragmentation of the undirected graph: {:1.3f}\n".format(1 - (sumP / (GM.number_of_nodes() * (GM.number_of_nodes() - 1))))) 
    print("Fragmentation Distance Weighted: {:1.3f}\n".format(1-1/(sumW/(sumU))))#Replace sumU with 2*sum for the alternative formula
def ReciprocityGeodesic(shortest):
    if isD:
        geo = nx.DiGraph()
    else: 
        geo = nx.Graph()
    for x in GM.nodes():
        geo.add_node(x)
    for k, v in shortest:
        for x,y in v.items():
            if x!=k:
                geo.add_edge(k, x, weight=len(y)-1) 
    if outputEdgelists.get() and len(GM.nodes())<16000:
        write_xlsx(nx.to_pandas_adjacency(geo),matrixName.rsplit( ".", 1 )[ 0 ] + " Geodesics (of Binary Graph).xlsx")
    if isD:
        print("Arc Reciprocity of the geodesic graph: {:1.3f}".format(nx.overall_reciprocity(geo)) + "\n")

    geo.clear()
    for x in GM.nodes():
        geo.add_node(x)
    for k, v in shortest:
        for x,y in v.items():
            prev=None
            if len(y)>1:
                tempSum=0
                for z in y:
                    if prev==None:
                        prev=z
                    else:        
                        tempSum+=GM.edges[prev,z]['weight']
                        prev=z    
                geo.add_edge(k, x, weight=tempSum)    
    if outputEdgelists.get() and len(GM.nodes())<16000:         
        write_xlsx(nx.to_pandas_adjacency(geo),matrixName.rsplit( ".", 1 )[ 0 ] + " Geodesic (of Weighted Graph).xlsx")
    if isD:
        undirLinks=0
        for x,y in geo.edges():
            if geo.get_edge_data(x,y) is not None and geo.get_edge_data(y,x)is not None:     
                undirLinks+=abs(geo.get_edge_data(x,y)['weight']-geo.get_edge_data(y,x)['weight']) 
        print("Arc Reciprocity of the weigthed geodesic graph: {:1.3f}".format(1-(undirLinks)/geo.size(weight = 'weight')) + "\n")
def GxRCentralizationDir(shortest):
    global worksheetGxRC
    worksheetGxRC = writer.add_worksheet("GxR Centrality") 
    counter=1
    worksheetGxRC.set_column(0, 0, colSize)
    for x in GM.nodes():
        worksheetGxRC.write(counter,0,str(x), main_format)  
        counter+=1  

    geo=nx.DiGraph()
    for x in GM.nodes():
        geo.add_node(x)
    for k, v in shortest:
        for x,y in v.items():
            if x!=k:
                geo.add_edge(k, x, weight=1) 
    size=len(geo.nodes())
    den=(size-1)**2

    maxI=0
    maxO=0
    outValI=[]
    outValO=[]
    for x in geo.nodes():
        tempI=geo.in_degree(x,weight='weight')
        tempO=geo.out_degree(x,weight='weight')
        outValI.append(tempI)    
        outValO.append(tempO)    
        if tempI>maxI:
            maxI=tempI      
        if tempO>maxO:
            maxO=tempO  

    worksheetGxRC.write(0,1,"ALIRCB", main_format)   
    worksheetGxRC.write(0,2,"NLIRCB", main_format)    
    counter=1
    for x in outValI:
        worksheetGxRC.write(counter,1,x, side_format) 
        worksheetGxRC.write(counter,2,x/(size-1), side_format)  
        counter+=1     
    worksheetGxRC.write(0,5,"ALORCB", main_format) 
    worksheetGxRC.write(0,6,"NLORCB", main_format)        
    counter=1
    for x in outValO:
        worksheetGxRC.write(counter,5,x, side_format)  
        worksheetGxRC.write(counter,6,x/(size-1), side_format) 
        counter+=1  

    nGorI=0
    nGorO=0  
    for x in range(0,size):
        nGorI+=(maxI-outValI[x])
        nGorO+=(maxO-outValO[x])
    print("Global In-Reaching Centralization Absolute (Geodesic of Binary Graph): {:1.3f}".format(nGorI) + "\n")  
    print("Global In-Reaching Centralization Normalized (Geodesic of Binary Graph): {:1.3f}".format(nGorI/den) + "\n") 
    print("Global Out-Reaching Centralization Absolute (Geodesic of Binary Graph): {:1.3f}".format(nGorO) + "\n")
    print("Global Out-Reaching Centralization Normalized (Geodesic of Binary Graph): {:1.3f}".format(nGorO/den) + "\n")

    IncrementBar("Calculating GxR of Weighted Graph")   
    geo.clear()
    for x in GM.nodes():
        geo.add_node(x)
    for k, v in shortest:
        for x,y in v.items():
            prev=None
            if len(y)>1:
                tempSum=0
                for z in y:
                    if prev==None:
                        prev=z
                    else:        
                        tempSum+=GM.edges[prev,z]['weight']
                        prev=z    
                geo.add_edge(k, x, weight=tempSum)  
    size=len(geo.nodes())  
    denW=(size-1)

    outValIW=[]
    outValOW=[]
    maxIW=0
    maxOW=0
    for x in geo.nodes():
        tempI=geo.in_degree(x,weight='weight')
        tempO=geo.out_degree(x,weight='weight')
        outValIW.append(tempI)   
        outValOW.append(tempO)  
        if tempI>maxIW:
            maxIW=tempI    
        if tempO>maxOW:
            maxOW=tempO

    worksheetGxRC.write(0,3,"ALIRCW", main_format)      
    counter=1
    for x in outValIW:
        worksheetGxRC.write(counter,3,x, side_format)  
        counter+=1  
    worksheetGxRC.write(0,7,"ALORCW", main_format)      
    counter=1
    for x in outValOW:
        worksheetGxRC.write(counter,7,x, side_format)  
        counter+=1      

    nGorIW=0
    nGorOW=0
    for x in range(0,size):
        nGorIW+=(maxIW-outValIW[x])
        nGorOW+=(maxOW-outValOW[x])  

    worksheetGxRC.write(0,4,"NLIRCW", main_format)      
    counter=1
    maxLorcIW=0
    for x,y in zip(outValIW,outValI):
        tempI=0
        if x!=0:
            tempI=y/x/(size-1)   
        if tempI>maxLorcIW:
            maxLorcIW=tempI
        outValIW[outValIW.index(x)]=tempI
        worksheetGxRC.write(counter,4,tempI, side_format)  
        counter+=1  

    worksheetGxRC.write(0,8,"NLORCW", main_format)      
    counter=1
    maxLorcOW=0
    for x,y in zip(outValOW,outValO):
        tempO=0
        if x!=0:
            tempO=y/x/(size-1) 
        if tempO>maxLorcOW:
            maxLorcOW=tempO 
        outValOW[outValOW.index(x)]=tempO
        worksheetGxRC.write(counter,8,tempO, side_format)  
        counter+=1  

    nLorcIW=0
    nLorcOW=0
    for x in range(0,size):
        nLorcIW+=(maxLorcIW-outValIW[x])
        nLorcOW+=(maxLorcOW-outValOW[x]) 

    print("Global In-Reaching Centralization Absolute (Geodesic of Weighted Graph): {:1.3f}".format(nGorIW) + "\n")  
    print("Global In-Reaching Centralization Normalized (Geodesic of Weighted Graph): {:1.3f}".format(nLorcIW/denW) + "\n")  
    print("Global Out-Reaching Centralization Absolute (Geodesic of Weighted Graph): {:1.3f}".format(nGorOW) + "\n") 
    print("Global Out-Reaching Centralization Normalized (Geodesic of Weighted Graph): {:1.3f}".format(nLorcOW/denW) + "\n") 
def GxRCentralizationUndir(shortest, geo):
    global worksheetGxRC
    worksheetGxRC = writer.add_worksheet("GxR Centrality") 
    worksheetGxRC.set_column(0, 0, colSize)
    counter=1
    for x in GM.nodes():
        worksheetGxRC.write(counter,0,str(x), main_format)  
        counter+=1 

    size=len(geo.nodes())
    den=(size-1)*(size-2)

    maxD=0
    outValD=[]
    for x in geo.nodes():
        temp=geo.degree(x,weight=1)
        outValD.append(temp)      
        if temp>maxD:
            maxD=temp    

    worksheetGxRC.write(0,1,"ALRCB", main_format)   
    worksheetGxRC.write(0,2,"NLRCB", main_format)    
    counter=1
    for x in outValD:
        worksheetGxRC.write(counter,1,x, side_format) 
        worksheetGxRC.write(counter,2,x/(size-1), side_format)  
        counter+=1     

    nGorD=0
    for x in range(0,size):
        nGorD+=(maxD-outValD[x])
    print("Global Reaching Centralization Absolute (Geodesic of Binary Graph): {:1.3f}".format(nGorD) + "\n")  
    print("Global Reaching Centralization Normalized (Geodesic of Binary Graph): {:1.3f}".format(nGorD/den) + "\n") 

    IncrementBar("Calculating GxR of Weighted Graph")    
    denW=(size-2) 

    outValDW=[]
    maxDW=0
    for x in geo.nodes():
        temp=geo.degree(x,weight='weight')
        outValDW.append(temp)   
        if temp>maxDW:
            maxDW=temp   

    worksheetGxRC.write(0,3,"ALRCW", main_format)      
    counter=1
    for x in outValDW:
        worksheetGxRC.write(counter,3,x, side_format)  
        counter+=1      

    nGorDW=0
    for x in range(0,size):
        nGorDW+=(maxDW-outValDW[x])

    worksheetGxRC.write(0,4,"NLRCW", main_format)      
    counter=1
    maxLorcDW=0
    for x,y in zip(outValDW,outValD):
        temp=0
        if x!=0:
            temp=y/x/(size-1)   
        if temp>maxLorcDW:
            maxLorcDW=temp
        outValDW[outValDW.index(x)]=temp
        worksheetGxRC.write(counter,4,temp, side_format)  
        counter+=1  

    nLorcDW=0
    for x in range(0,size):        
        nLorcDW+=(maxLorcDW-outValDW[x])

    print("Global Reaching Centralization Absolute (Geodesic of Weighted Graph): {:1.3f}".format(nGorDW) + "\n")  
    print("Global Reaching Centralization Normalized (Geodesic of Weighted Graph): {:1.3f}".format(nLorcDW/denW) + "\n")  

#Special Indicators
def CalculateCliques():
    worksheetCliques = writer.add_worksheet("Cliques") 
    worksheetCliques.write(0,0,"Cliques", main_format)  
    counter=1
    totalC=0    
    for x in sorted(list(nx.enumerate_all_cliques(GM)),key=len,reverse=True):
        if len(x)>2:
            totalC+=1
            if outputEdgelists.get() and len(x)>2:
                worksheetCliques.write(counter,0,", ".join(sorted(x)), side_format)  
            counter+=1     
    worksheetCliques.write(0,1,counter-1, side_format)
    print("Cliques count (3+ nodes): {:.3f}\n".format(totalC))      
def CalculateCycles():
    worksheetCycles = writer.add_worksheet("Cycles") 
    worksheetCycles.write(0,0,"Cycles", main_format)  
    counter=1
    counter2=1
    totalC=0    
    for x in sorted(list(nx.simple_cycles(GM)),key=len,reverse=True):
        if len(x)>2:
            totalC+=1
            if outputEdgelists.get() and len(x)>2: 
                worksheetCycles.write(counter2,0,", ".join(x), side_format) 
                counter2+=1 
            counter+=1     
    worksheetCycles.write(0,1,counter-1, side_format)
    print("Cycles count (3+ nodes): {:.3f}\n".format(totalC))  

#Communities
def CalculateGirvanNewman():
    worksheetKC = writer.add_worksheet("Girvan-Newman") 
    mgnc= None
    for x in nx.algorithms.community.centrality.girvan_newman(GM):
        counter=0
        mgnc=x
        for y in sorted(mgnc,key=len,reverse=True):
            if len(y)>0:
                worksheetKC.write(0,counter,"GN-C "+str(counter+1), main_format) 
                worksheetKC.write(1,counter,len(y), main_format) 
                worksheetKC.write(2,counter,GM.subgraph(y).size(), main_format) 
                counter2=3
                for z in sorted(y,key=lambda x: x):
                    worksheetKC.write(counter2,counter,str(z), side_format) 
                    counter2+=1
                if outputEdgelists.get() and len(y)>3:
                    write_dl(GM.subgraph(y), matrixName.rsplit( ".", 1 )[ 0 ] +"'s GN-C "+ str(counter+1) +" DL.txt") 
                counter+=1   
        print("Girvan-Newman Communities count: {:.0f}\n".format(counter))    
        break
    #contract nodes
    if isD:
        GS=nx.DiGraph()
    else:
        GS=nx.Graph()
    for x in range(0,len(mgnc)):
        GS.add_node("GN-C "+str(x+1))
    cx=0
    for x in mgnc:
        cy=0
        for y in mgnc:
            if x!=y:
                for z,c in GM.edges():
                    if z in x and c in y:
                        GS.add_edge("GN-C "+str(cx+1), "GN-C "+str(cy+1))
                        break
            cy+=1
        cx+=1      
    write_edgelist(GS, matrixName.rsplit( ".", 1 )[ 0 ] +"'s GN-C structure edge list.xlsx")
def CalculateCommunities():
    worksheetKC = writer.add_worksheet("Communities") 
    try:
        counter=0
        counter4=0
        comm=nx.algorithms.community.greedy_modularity_communities(GM, weight='weight')
        for y in comm:
            counter4+=1
            if len(y)>0:
                worksheetKC.write(0,counter,"Comm. "+str(counter+1), main_format) 
                worksheetKC.write(1,counter,len(y), main_format) 
                worksheetKC.write(2,counter,GM.subgraph(y).size(), main_format) 
                counter2=3
                for z in sorted(y,key=lambda x: x):
                    worksheetKC.write(counter2,counter,str(z), side_format) 
                    counter2+=1
                if outputEdgelists.get() and len(y)>3:
                    write_dl(GM.subgraph(y), matrixName.rsplit( ".", 1 )[ 0 ] +"'s Comm. "+ str(counter+1) +" DL.txt") 
                counter+=1   
        print("Greedy Modular Communities count: {:.0f}\n".format(counter4)) 
    except (KeyError, IndexError):
        worksheetKC.write(1,0,"NetworkX Error", side_format)  

#K-Stuff
def CalculateKCore():
    kcnd=nx.core_number(GM)
    worksheetKCore = writer.add_worksheet("K-core") 
    worksheetKCore.set_column(0, 0, colSize)
    worksheetKCore.write(0,0,"Nodes", main_format)  
    worksheetKCore.write(0,1,"K-core number", main_format) 
    counter=1
    for x in GM.nodes():
        worksheetKCore.write(counter,0,str(x), main_format)  
        counter+=1    
    counter=1
    for x,y in kcnd.items():
        worksheetKCore.write(counter,1,y, side_format) 
        counter+=1  

    kcoreGraph=nx.k_core(GM, k=None, core_number=kcnd)
    write_edgelist(kcoreGraph, matrixName.rsplit( ".", 1 )[ 0 ] +"'s k-core edge list.xlsx")

    kshellGraph=nx.k_shell(GM, k=None, core_number=kcnd)
    write_edgelist(kshellGraph, matrixName.rsplit( ".", 1 )[ 0 ] +"'s k-shell edge list.xlsx")   
def CalculateKComponents():
    worksheetKC = writer.add_worksheet("K-components") 
    worksheetKC.write(0,0,"K", main_format)  
    worksheetKC.write(0,1,"Count", main_format)  
    counter=1
    for x,y in sorted(nx.k_components(GM).items(), reverse=True):
        worksheetKC.write(counter,0,str(x), side_format)  
        counter2=2
        for z in y:
            worksheetKC.write(counter,counter2,", ".join(sorted(z)), side_format)  
            counter2+=1
        worksheetKC.write(counter,1,counter2-2, side_format)  
        counter+=1     

#Structural Holes
def CalculateStructuralHoles():
    worksheetSH = writer.add_worksheet("Structural Holes") 
    worksheetSH.write(0,0,"Nodes", main_format)  
    worksheetSH.set_column(0, 0, colSize)
    counter=1
    for x in GM.nodes():
        worksheetSH.write(counter,0,str(x), main_format)  
        counter+=1    

    IncrementBar("Calculating Constraint Binary")
    worksheetSH.write(0,1,"Constraint Binary", main_format)   
    counter=1
    for x,y in nx.constraint(GM,weight='None').items():
        if not math.isnan(y):
            worksheetSH.write(counter,1,y, side_format)  
        counter+=1    

    IncrementBar("Calculating Constraint Weighted")
    worksheetSH.write(0,2,"Constraint Weighted", main_format) 
    counter=1
    for x,y in nx.constraint(GM,weight='weight').items():
        if not math.isnan(y):
            worksheetSH.write(counter,2,y, side_format)  
        counter+=1     

    IncrementBar("Calculating Effective Size Binary")
    worksheetSH.write(0,3,"Effective Size Binary", main_format) 
    counter=1
    for x,y in nx.effective_size(GM,weight='None').items():
        if not math.isnan(y):
            worksheetSH.write(counter,3,y, side_format)  
        counter+=1     

    IncrementBar("Calculating Effective Size Weighted")
    worksheetSH.write(0,4,"Effective Size Weighted", main_format)  
    counter=1
    for x,y in nx.effective_size(GM,weight='weight').items():
        if not math.isnan(y):
            worksheetSH.write(counter,4,y, side_format)  
        counter+=1        

    IncrementBar("Calculating Local Constraint Binary")
    worksheetLCB = writer.add_worksheet("Local Constraint Binary") 
    counter=1
    for x in GM.nodes():
        worksheetLCB.write(counter,0,str(x), main_format)  
        worksheetLCB.write(0,counter,str(x), main_format) 
        counter+=1     
    xC=1
    yC=1
    for u in GM.nodes():
        for v in GM.nodes():
            if u!=v:
                lc=nx.local_constraint(GM, u, v, weight=None)
                if lc!=0:
                    worksheetLCB.write(xC,yC,lc, side_format) 
            yC+=1
        xC+=1
        yC=1

    IncrementBar("Calculating Local Constraint Weighted")
    worksheetLCW = writer.add_worksheet("Local Constraint Weighted") 
    counter=1
    for x in GM.nodes():
        worksheetLCW.write(counter,0,str(x), main_format)  
        worksheetLCW.write(0,counter,str(x), main_format) 
        counter+=1     

    xC=1
    yC=1
    for u in GM.nodes():
        for v in GM.nodes():
            if u!=v:
                lc=nx.local_constraint(GM, u, v, weight='weight')
                if lc!=0:
                    worksheetLCW.write(xC,yC,lc, side_format) 
            yC+=1
        xC+=1
        yC=1

#Components
def ConnectedComponents():
    worksheetCC = writer.add_worksheet("CCs") 
    counter=0
    totalC=0
    for y in reversed(sorted(nx.connected_components(GM.to_undirected()), key=len)):
        totalC+=1
        if len(y)>1:
            worksheetCC.write(0,counter,"CC "+str(counter+1), main_format) 
            worksheetCC.write(1,counter,len(y), main_format) 
            worksheetCC.write(2,counter,GM.subgraph(y).size(), main_format) 
            counter2=3
            for z in sorted(y,key=lambda x: x):
                worksheetCC.write(counter2,counter,str(z), side_format) 
                counter2+=1
            if outputEdgelists.get() and len(y)>3:
                write_dl(GM.subgraph(y), matrixName.rsplit( ".", 1 )[ 0 ] +"'s CC "+ str(counter+1) +" DL.txt") 
            counter+=1    
    print("Connected Components count (including isolated nodes): {:.0f}\n".format(totalC))                 
def StronglyConnectedComponents():
    worksheetSCC = writer.add_worksheet("SCCs")  
    counter=0
    totalC=0
    for y in reversed(sorted(nx.strongly_connected_components(GM), key=len)):
        if len(y)>1:
            totalC+=1
            worksheetSCC.write(0,counter,"SCC "+str(counter+1), main_format) 
            worksheetSCC.write(1,counter,len(y), main_format) 
            worksheetSCC.write(2,counter,GM.subgraph(y).size(), main_format) 
            counter2=3
            for z in sorted(y,key=lambda x: x):
                worksheetSCC.write(counter2,counter,str(z), side_format)  
                counter2+=1
            counter+=1      
    print("Strongly Connected Components count (excluding isolated nodes): {:.0f}\n".format(totalC))

#Additional Sheet Indicators
def AssortativityAttributes():
    global attString
    if attString:
        print("Attribute Assortativity Coefficient: {:.3f}\n".format(nx.attribute_assortativity_coefficient(GM,"assortativity")))    
    else:
        print("Numeric Assortativity Coefficient: {:.3f}\n".format(nx.numeric_assortativity_coefficient(GM,"assortativity")))  
def GroupDataCentralities():
    try:
        print("Group Betweenness Centrality: {:.3f}".format(nx.group_betweenness_centrality(GM,groupValues,normalized=False)) + "\n")
        print("Group Betweenness Centrality Normalized: {:.3f}".format(nx.group_betweenness_centrality(GM,groupValues)) + "\n")
        print("Group Betweenness Centrality Weighted: {:.3f}".format(nx.group_betweenness_centrality(GM,groupValues,weight='weight',normalized=False)) + "\n")
        print("Group Betweenness Centrality Weighted Normalized: {:.3f}".format(nx.group_betweenness_centrality(GM,groupValues,weight='weight')) + "\n")
    except KeyError as e:
        print("Can't calculate group betweenness! (Because of "+ str(e) +") Try expanding the group.\n")    
    if isD:    
        print("Group In-Closeness Centrality: {:.3f}".format(nx.group_closeness_centrality(GM,groupValues)) + "\n")
        print("Group In-Closeness Centrality Weighted: {:.3f}".format(nx.group_closeness_centrality(GM,groupValues,weight='weight')) + "\n")
        print("Group Out-Closeness Centrality: {:.3f}".format(nx.group_closeness_centrality(GM.reverse(),groupValues)) + "\n")
        print("Group Out-Closeness Centrality Weighted: {:.3f}".format(nx.group_closeness_centrality(GM.reverse(),groupValues,weight='weight')) + "\n")
        print("Group In-Degree Centrality: {:.3f}".format(nx.group_in_degree_centrality(GM,groupValues)) + "\n")
        print("Group Out-Degree Centrality: {:.3f}".format(nx.group_out_degree_centrality(GM,groupValues)) + "\n")
    else:
        print("Group Closeness Centrality: {:.3f}".format(nx.group_closeness_centrality(GM,groupValues)) + "\n")
        print("Group Closeness Centrality Weighted: {:.3f}".format(nx.group_closeness_centrality(GM,groupValues,weight='weight')) + "\n")
        print("Group Degree Centrality: {:.3f}".format(nx.group_degree_centrality(GM,groupValues)) + "\n")

#Special Indicators
def CumulativeFlow():
    worksheetC = writer.add_worksheet("Cumulative Flow") 
    xPointer=1
    yPointer=1 
    #isDirected()
    descendants=dict()
    ancestors=dict()
    dx=set()
    dy=set()
    dz=set()
    for x in GM.nodes():
        worksheetC.write(0,yPointer,str(x), main_format) 
        worksheetC.write(yPointer,0,str(x), main_format) 
        xPointer=1
        for y in GM.nodes():
            if x!=y and nx.has_path(GM,x,y):
                dx=nx.descendants(GM.subgraph([n for n in GM.nodes() if n != y]),x)
                dy=nx.ancestors(GM.subgraph([n for n in GM.nodes() if n != x]),y)
                dz=dx.intersection(dy)
                dz.add(x)
                dz.add(y)          
                worksheetC.write(yPointer,xPointer,GM.subgraph(dz).size(weight='weight'),side_format) 
            xPointer+=1
        yPointer+=1
def printHeader():
    print("Excel Matrix name: "+ str(matrixName))    
    print("Directed: "+ str(isDirected()))
    print("Weighted: " + str(isWeighted()))
    print("Connected: " + str(nx.is_connected(GM.to_undirected())))
    print("Self-links: " + str(hasSelfLoops()))
    print("")

#Main Function
def CalculateGeodesicsIndicators():
    IncrementBar("Calculating all shortest paths")
    shortest=list(nx.all_pairs_dijkstra_path(GM)) 
    IncrementBar("Calculating Fragmentation")
    Fragmentation(shortest)
    IncrementBar("Creating Geodesic Matrices")
    if isD:
        ReciprocityGeodesic(shortest)   
        IncrementBar("Calculating GxR Centralization")
        GxRCentralizationDir(shortest)
        shortest=None
        geo=None
    else:
        if outputEdgelists.get() and len(GM.nodes())<16000:
            ReciprocityGeodesic(shortest)   
        geo = nx.Graph()
        for x in GM.nodes():
            geo.add_node(x)
        for k, v in shortest:
            for x,y in v.items():
                if not geo.has_edge(x,k) and len(y)>1:
                    prev=None
                    for z in y:
                        if prev==None:
                            prev=z
                        else:
                            if geo.has_edge(k,x):
                                geo.edges[k,x]['weight']+=GM.edges[prev,z]['weight']
                            else:
                                geo.add_edge(k, x, weight=GM.edges[prev,z]['weight'])
                            prev=z 
        shortest=None
        IncrementBar("Calculating GxR Centralization")
        GxRCentralizationUndir(shortest,geo)
        geo=None

def ASEN(): 
    global prog
    global totalInd  
    global phi

    if algorithm.get()=="Standard Indicators":
        #Basic stuff
        IncrementBar("Calculating Basic Indicators")
        Size()
        AbsoluteDensity()
        NormalizedDensity()
        ValueOfNetworkLinks()
        DisconnectednessDegree()
        AverageDegreeCentralityBinary()
        AverageDegreeCentralityValued()        
        AverageLinkWeight()       
        GlobalClusteringCoefficient()
        if isD:
            ReciprocityG()
            ReciprocityWeightedG()
        AssortativityBinary()    
        AssortativityValued()
        SnijderCentralization() 
        Triads()

        #Centralities
        IncrementBar("Calculating Centrality Indicators")
        if isD:
            InDegreeCentralityBinary()
            OutDegreeCentralityBinary()
            InDegreeCentralityValued()
            OutDegreeCentralityValued()
            BetweennessCentrality()     
            BetweennessCentralityW()      
            ClosenessCentralityIn()     
            ClosenessCentralityOut()  
        else:
            InDegreeCentralityBinary()
            InDegreeCentralityValued()
            BetweennessCentrality()     
            BetweennessCentralityW()      
            ClosenessCentralityIn()    

        IncrementBar("Calculating Eigenvector Indicators")
        if isD:
            EVCentralityIn()
            EVCentralityOut()
            EVCentralityInWeighted()
            EVCentralityOutWeighted()
        else:
            EVCentralityIn()
            EVCentralityInWeighted()

        #Components
        IncrementBar("Calculating Components Indicators")
        ConnectedComponents()
        if isD:
            StronglyConnectedComponents()    

        #Calculate if possible
        IncrementBar("Calculating Extra Sheet Indicators")
        if assortativityData:
            AssortativityAttributes() 
        if groupData:
            GroupDataCentralities()    
    elif algorithm.get()=="Structural Holes Indicators":
        CalculateStructuralHoles()
    elif algorithm.get()=="Communities Indicators":
        IncrementBar("Calculating Girvan-Newman Indicator")
        CalculateGirvanNewman() 
        IncrementBar("Communities Indicator")
        CalculateCommunities()
    elif algorithm.get()=="Geodesics Indicators":
        CalculateGeodesicsIndicators()        
    elif algorithm.get()=="Cycles and Cliques":        
        if isD:
            IncrementBar("Calculating Cycles Indicator")
            CalculateCycles() 
        else:
            IncrementBar("Calculating Cliques Indicator")
            CalculateCliques()
    elif algorithm.get()=="Katz and RWBC":  
        if isD:
            IncrementBar("Calculating Katz' Alpha")              
            analyzeBtn.config(state=DISABLED)   
            if alpha.get()=="":
                phi=max(nx.adjacency_spectrum(GM,weight='weight'))        
                try:
                    if (1/float(phi))<0.1:
                        if (1/float(phi))<0:
                            alpha.delete(0, END)
                            alpha.insert(0,"0")
                        else:
                            alpha.delete(0, END)
                            alpha.insert(0,str(1/float(phi)))
                except ZeroDivisionError:
                    alpha.insert(0,"1/0!") 
                    return
                global barTick 
                global prog
                bar['value']+= barTick
                root.update()        
            IncrementBar("Calculating Katz Centralities")                
            analyzeBtn.config(state=DISABLED)   
            try:
                KatzCentralityIn()
                KatzCentralityOut()
                KatzCentralityInWeighted()
                KatzCentralityOutWeighted()
            except np.linalg.LinAlgError:
                 ShowError("Katz Error","Singular matrix error, change alpha!")                 
            analyzeBtn.config(state=DISABLED)   
        else:
            IncrementBar("Calculating Random Walk Indicators")
            if nx.is_connected(GM.to_undirected()):
                RandomWalk()
                RandomWalkNormalized()
                WeightedRandomWalk()
                WeightedRandomWalkNormalized() 
    elif algorithm.get()=="K-core and K-components":   
        if not hasSelfLoops():
            IncrementBar("Calculating K-core and K-shell")
            CalculateKCore()   
            if not isD:
                IncrementBar("Calculating K-Components")
                CalculateKComponents() 
    elif algorithm.get()=="Cumulative Flow":  
        IncrementBar("Calculating Cumulative Flow")
        if isD:
            CumulativeFlow()


#GUI code
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()
def OpenFileClick():
    global directory
    global matrixName
    bar['value']=0
    analyzeBtnText.set("Analyze")    
    path = askopenfilename(initialdir = directory, title = "Choose input file", filetypes =[("Graph input files", "*.xlsx *.xls *.txt *.nxg")]  )
    if path!="":
        matrixName = os.path.basename(path)
        directory = os.path.dirname(path)         
        CreateGraph(path) 
    CheckStatus()
def GetProgreessBarTick():
    global tCount
    global barTick
    global totalInd
    global prog
    prog=0
    steps=1

    if algorithm.get()=="Standard Indicators":
        steps+=5
    elif algorithm.get()=="Structural Holes Indicators":
        steps+=6
    elif algorithm.get()=="Communities Indicators":
        steps+=2
    elif algorithm.get()=="Geodesics Indicators":
        steps+=5
    elif algorithm.get()=="Cycles and Cliques":
        steps+=1
    elif algorithm.get()=="Katz and RWBC":
        if isD:
            steps+=2
        else:
            steps+=1
    elif algorithm.get()=="K-core and K-components":
        if isD:
            steps+=1
        else:
            steps+=2
    elif algorithm.get()=="Cumulative Flow":
        steps+=1
    
    totalInd=steps
    barTick=100/steps
def IncrementBar(string):
    global barTick 
    global prog
    prog+=1
    bar['value']+= barTick
    analyzeBtnText.set(string+" ("+str(prog-1)+"/"+str(totalInd)+")")     
    root.update() 
def getRP():
    global rp
    return int(rp)
def setRP():
    global rp
    rp+=1
def AnalyzeClick():
    global writer
    global rp
    global workbook
    global prog
    global main_format
    global side_format
    global mixed_format  
    global worksheetMain
    global groupData
    if waitForMatrix:
        messagebox.showinfo("Alert", "Open a matrix first!")
    else:
        suffix=""
        if algorithm.get()=="Standard Indicators":
            suffix="Analysis"
        elif algorithm.get()=="Structural Holes Indicators":
            suffix="Structural Holes"
        elif algorithm.get()=="Communities Indicators":
            suffix="Communities"
        elif algorithm.get()=="Geodesics Indicators":
            suffix="Geodesics"
        elif algorithm.get()=="Cycles and Cliques":
            suffix="Cycles and Cliques"
        elif algorithm.get()=="Katz and RWBC":
            suffix="Katz and RWBC"
        elif algorithm.get()=="K-core and K-components":
            suffix="K-core and K-components"
        elif algorithm.get()=="Cumulative Flow":
            suffix="Cumulative Flow"
        OpenFile(""+matrixName.rsplit( ".", 1 )[ 0 ] + " ASEN "+suffix+".txt") 
        printHeader()
        workbook = pd.ExcelWriter(""+matrixName.rsplit( ".", 1 )[ 0 ] + " ASEN "+suffix+".xlsx", engine='xlsxwriter')#xlsxwriter.Workbook("ASEN Analysis on "+matrixName.rsplit( ".", 1 )[ 0 ] + " Analysis.xlsx",{'nan_inf_to_errors': True})               
        writer=workbook.book
        InitializeExcel()
        if algorithm.get()=="Standard Indicators":
            worksheetMain = writer.add_worksheet("Dc-Bc-Cc")            
            worksheetMain.set_column(0, 0, colSize)
            counter=1
            for x in GM.nodes():
                worksheetMain.write(counter,0,str(x), main_format)  
                counter+=1  
        analyzeBtn.config(state=DISABLED)      
        GetProgreessBarTick()           
        analyzeBtnText.set("Analyzing...")
        rp=1
        prog=1        

        a = datetime.datetime.now() 
        ASEN() 
        b = datetime.datetime.now()    
        delta = b - a
        print("Elapsed time: " + str(delta))
        CloseFile()
        excelSaved=False
        while not excelSaved:
            try:
                workbook.save()
                excelSaved=True            
            except xlsxwriter.exceptions.FileCreateError:
                excelSaved=False        
                messagebox.showinfo("Output error", "Excel File already open! Close it and click OK")
        IncrementBar("All done! Time Elapsed: "+ str(delta))

def ShowError(title,message):
    global waitForMatrix
    global excelMatrix
    global excelMatrix2

    root.update() 
    analyzeBtn.config(state=NORMAL)
    analyzeBtnText.set("Analyze") 
    bar['value']=0   
    messagebox.showinfo(title,message)
def CheckStatus(*heh):    
    root.update() 
    if (matrixName is not None and matrixName !=""):
        analyzeBtn.config(state=NORMAL)   
    else:
        analyzeBtn.config(state=DISABLED)
    analyzeBtnText.set("Analyze") 
    bar['value']=0   
def ValidateRegex(string):
    regex = re.compile(r"^([0-9]|[1-9]([0-9])*)([.][0-9]*)?$")
    result = regex.match(string)
    return (string == ""
            or (string.count('.') <= 1
                and result is not None
                and float(string)<=100 and float(string)>=0                
                and result.group(0) != ""))
def ValidateEntry(P): 
    try: 
        analyzeBtn
    except NameError:
        pass
    else:
        CheckStatus()
    return True

root = Tk()
root.geometry("295x285") 
root.resizable(0, 0)
root.title("ASEN v1.0.0")
vldtNmbr = (root.register(ValidateEntry), '%P')

matrixLabel = Label(root, text ="Input graph topology:", width=26).grid(row=0, columnspan=4,sticky=EW)
matrixText=StringVar()
matrixText.set("[None]")
outputBtn=Button(root, textvariable=matrixText, command = OpenFileClick, width=10).grid(row=1, padx = 10, columnspan=4,sticky=EW)

algLabel = Label(root, text ="Indicators Group:", width=26).grid(row=2, pady=(10,0), columnspan=4,sticky=EW)
algorithm = StringVar()
algorithm.set("Standard Indicators") 
algorithmMenu = OptionMenu(root, algorithm, "Standard Indicators", "Katz and RWBC", "Geodesics Indicators", "Communities Indicators","Structural Holes Indicators", "Cumulative Flow",    "Cycles and Cliques",  "K-core and K-components",  command =CheckStatus)
algorithmMenu.grid(row=3, column=0, padx=10, columnspan=4, sticky=EW)

matrixLabel = Label(root, text ="Katz Settings (don't edit if unsure):", width=26).grid(row=10, columnspan=4, sticky=EW, pady=(10,5))
labelA = Label(root, text ="Katz Alpha:")
labelB = Label(root,text ="Katz Beta:")
alpha = Entry(root,  validate='key',width=8, validatecommand=vldtNmbr)
#alpha.insert(0,0.1)
beta = Entry(root, validate='key', width=6,validatecommand=vldtNmbr)
beta.insert(0,1.0)
labelA.grid(row=11, column=0, padx=10, sticky=W)
alpha.grid(row=11, column=1, padx =10, sticky=W)
labelB.grid(row=11, column=2, padx=10, sticky=W)
beta.grid(row=11, column=3, padx = 10, sticky=W)
outputEdgelists = IntVar()
outputEdgelists.set(False)
outputEdgelistsBtn = Checkbutton(root, text="Enable extra output?", variable=outputEdgelists,command=CheckStatus)
outputEdgelistsBtn.grid(row=13, padx = 10,pady=(10,0), columnspan=4, sticky=W)
bar = Progressbar(root, length=100) 
bar.grid(row=17, column=0, pady=(20,5), padx=10, columnspan=4,sticky=EW)

analyzeBtnText = StringVar()
analyzeBtnText.set("Analyze")
analyzeBtn = Button(root, textvariable=analyzeBtnText, command = AnalyzeClick,width=30)
analyzeBtn.grid(row=18,  column=0, padx=10, sticky=EW, columnspan=4)
analyzeBtn.config(state=DISABLED) 

center(root)
root.mainloop()

