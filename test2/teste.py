import os
import win32com.client
import pint
import numpy as np
import pandas as pd

SapObject = win32com.client.Dispatch("Sap2000v16.SapObject")
SapObject.ApplicationStart()
SapModel = SapObject.SapModel

def start_up():

    SapModel.InitializeNewModel()

    filename = r"C:\Users\User\Desktop\SAP\sap_interacao\teste2\ITERAÇÃO SOLO - ESTRUTURA\22.02.18.TESC silo.SDB"

    ret = SapModel.File.OpenFile(filename)
    if ret == 0:
        print('Model Loaded')

    ret = SapModel.SetModelIsLocked(False)

    ret = SapModel.View.RefreshView(0, False)

    ret = SapModel.File.Save(filename)

    print('Running Analysis...')
    ret = SapModel.Analyze.RunAnalysis()
    

    ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
    ret = SapModel.Results.Setup.SetComboSelectedForOutput("COMB1")

    print('Analysis finished.')
    

def springs_names():
    _, numberOfPoints, pointsName = SapModel.PointObj.GetNameList()

    spring_stats = []
    spring_name = []

    for pts in pointsName:
        spring = SapModel.PointObj.GetSpring(pts, [0,0,0,0,0,0])
        if spring[0] == 0:
            spring_stats.append(spring[1])
            spring_name.append(pts)
        else:
            pass

    return (spring_name, spring_stats)

def group_name():
    _, numberOfGroups, groupsName = SapModel.GroupDef.GetNameList()

    return groupsName

def group_points(group_name):
    _, n, groupObj_type, groupObj_name=SapModel.GroupDef.GetAssignments(group_name)

    return groupObj_type, groupObj_name


def results_single_joint(result_type, spring_name):
    ObjectElm = 0
    NumberResults = 0
    Obj = ObjSta = Elm = LoadCase = StepType = StepNum = []
    F1 = F2 = F3 = M1 = M2 = M3 = []

    side_inputs = (ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)

    single_spring = spring_name

    results = {'react': SapModel.Results.JointReact(single_spring, *side_inputs),
                'displ': SapModel.Results.JointDispl(single_spring, *side_inputs)}

    return results[result_type]
    
def results_byGroup(result_type, groupName):
    ObjectElm = 2
    NumberResults = 0
    Obj = ObjSta = Elm = LoadCase = StepType = StepNum = []
    F1 = F2 = F3 = M1 = M2 = M3 = []

    side_inputs = (ObjectElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)

    results = {'react': SapModel.Results.JointReact(groupName, *side_inputs),
                'displ': SapModel.Results.JointDispl(groupName, *side_inputs)}

    return results[result_type]



def change_single_spring(spring_name):
    pass