import comtypes.client
import random
import numpy as np
import os


#创建6个文件夹用于存储频率、应力比、层间位移的data和label

CirFre_Data = 'CirFre_Data'
if not os.path.exists(CirFre_Data):
    os.makedirs(CirFre_Data)

Ratio_Data = 'Ratio_Data'
if not os.path.exists(Ratio_Data):
    os.makedirs(Ratio_Data)

Displ_Data = 'Displ_Data'
if not os.path.exists(Displ_Data):
    os.makedirs(Displ_Data)

CirFre_Label = 'CirFre_Label'
if not os.path.exists(CirFre_Label):
    os.makedirs(CirFre_Label)

Ratio_Label = 'Ratio_Label'
if not os.path.exists(Ratio_Label):
    os.makedirs(Ratio_Label)

Displ_Label = 'Displ_Label'
if not os.path.exists(Displ_Label):
    os.makedirs(Displ_Label)


# 连接到运行中的 SAP2000
sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
SapModel = sap_object.SapModel

# 解锁模型
SapModel.SetModelIsLocked(False)

# 设置单位
SapModel.SetPresentUnits(9)  # 设定单位为 N/mm^2


def random_goods_distribution(num_z, num_x, num_y):
    # 创建一个 num_z 行，num_x * num_y 列的矩阵
    distribution_matrix = []
    
    for layer in range(num_z):
        # 随机生成当前层的货物分布，0 或 1
        distribution = [random.choice([0, 1]) for _ in range(num_x * num_y)]
        distribution_matrix.append(distribution)  # 将当前层的分布添加到矩阵中
        
    return distribution_matrix


# 循环进行随机货位分布并计算频率

num_trials = 5 # 生成数据的条数
surface_pressure_area = 0.001  # 均布面力
num_z = 3 #  3 for real , 3 for test
num_x = 11 # 11 for real , 7 for test
num_y = 14 # 14 for real , 7 for test


#获取框架信息，用于后续输出应力比
frame_info=sap_object.SapModel.FrameObj.GetNameList()

#if frame_info != 0:
    #print("GetNameList not successfully.")
#else:
    #print("GetNameList successfully.")

#获取框架编号list   
frame_name=frame_info[1]
#print(frame_name)


#根据货位分布矩阵，对结构施加均布面力，然后进行分析以及设计

print("**************** Start process ******************")
#print('—————————————————————————————————————————————————————————————————————————————————————————————————————————')

num_loc = 462  # 货位及车道位总数 147 for test   462 for real

for trial in range(num_trials):
    # 随机分布货位
    distribution_matrix = random_goods_distribution(num_z, num_x, num_y)
    
    #print(f"Trial {trial + 1}: Goods Distribution Matrix:")
    #for layer in distribution_matrix:
        #print(layer)

    #flatten货位分布矩阵，便于后续操作
    flattened_distribution = [item for sublist in distribution_matrix for item in sublist]

    # 按照货位分布矩阵，在 SAP2000 中添加均布面力
    for loc in range(num_loc):
        if flattened_distribution[loc] == 1:
            area_name = str(loc + 1)
            
            ret=sap_object.SapModel.AreaObj.SetLoadUniform(area_name,
                                                       "DEAD",
                                                       surface_pressure_area,
                                                       10,
                                                       True,
                                                       "Global",   #"local"
                                                       0)
            
            #ret = sap_object.SapModel.AreaObj.SetLoadUniform(area_name,"DEAD",surface_pressure_area,10,True,"Global",0)
            #if ret != 0:
                #print("Area load not assigned successfully.")
            #else:
                #print("Area load assigned successfully.")
            

    #执行模态分析
    ret=SapModel.Analyze.RunAnalysis()
    #ret = SapModel.Analyze.RunAnalysis()
    #if ret != 0:
        #print("Analysis did not run successfully.")
    #else:
        #print("Analysis run successfully.")

    #执行钢结构设计
    ret=SapModel.DesignSteel.StartDesign()
    #if ret != 0:
        #print("StartDesign failed.")
    #else:
        #print("StartDesign succeeded.")
    


######################### 获取模态(前12阶圆频率)数据 ####################

    NumberResults = 12
    LoadCase = []
    StepType = []
    StepNum = []
    Period = []
    Frequency = []
    CircFreq = []
    EigenValue = []

    # 调用函数获取模态数据
    result = SapModel.Results.ModalPeriod(NumberResults, 
                                          LoadCase, 
                                          StepType, 
                                          StepNum, 
                                          Period, 
                                          Frequency, 
                                          CircFreq, 
                                          EigenValue)
    #if result != 0:
        #print("ModalPeriod failed.")
    #else:
        #print("ModalPeriod succeeded.")

    #print(f"Trial {trial + 1} CirFreq: {result[6]}")  # 6 代表CircFreq

    np.save(os.path.join(CirFre_Data,f'data_cirfre_{trial}'), distribution_matrix)
    np.save(os.path.join(CirFre_Label,f'label_cirfre_{trial}'), result[6])



########################### 获取应力比数据 ############################

    max_ratio = 0  

    for i in frame_name:
        FrameName = i 
        NumberItems = 0
        FrameNames = []
        Ratios = []
        RatioTypes = []
        Locations = []
        ComboNames = []
        ErrorSummary = []
        WarningSummary = []

        result_ratio = SapModel.DesignSteel.GetSummaryResults(FrameName, 
                                                              NumberItems, 
                                                              FrameNames, 
                                                              Ratios, 
                                                              RatioTypes, 
                                                              Locations, 
                                                              ComboNames, 
                                                              ErrorSummary, 
                                                              WarningSummary)
        
        #print(f"frame_{i}'s ratio:{result_ratio[2][0]}")
        if result_ratio[2][0] > max_ratio:
            max_ratio = result_ratio[2][0]
            
        #这里会持续输出“failed”，但是实际上是成功的，应力比会正常输出
        #if ret != 0:
            #print("GetSummaryResults failed.")
        #else:
            #print("GetSummaryResults succeeded.")
    

    #print(f"Trial {trial + 1} max ratio: {max_ratio}")

    np.save(os.path.join(Ratio_Data,f'data_ratio_{trial}'), distribution_matrix)
    np.save(os.path.join(Ratio_Label,f'label_ratio_{trial}'), max_ratio)   
    

########################### 获取层间位移数据 ############################

    num_floor = 3
    result_rad=[]
    for floor in range(num_floor):
        Name=f"rad_{floor}"
        NumberResults=0
        LoadCase=[]
        GD=[]
        LoadCase=[]
        StepType=[]
        StepNum=[]
        DType=[]
        Value=[]

        result_displ = SapModel.Results.GeneralizedDispl(Name, 
                                                         NumberResults, 
                                                         GD, 
                                                         LoadCase, 
                                                         StepType, 
                                                         StepNum, 
                                                         DType, 
                                                         Value)
        max_displ = max(result_displ[6])
        result_rad.append(max_displ)
    
        #if result_displ != 0:
            #print("GeneralizedDispl failed.")
        #else:
            #print("GeneralizedDispl succeeded.")  
             
    #print(f"Trial {trial+1} DisplRad: {result_rad}")

    np.save(os.path.join(Displ_Data,f'data_displ_{trial}'), distribution_matrix)
    np.save(os.path.join(Displ_Label,f'label_displ_{trial}'), result_rad)
    


    #print('—————————————————————————————————————————————————————————————————————————————————————————————————————————')
     #解锁模型
    SapModel.SetModelIsLocked(False)

    # 删除分析以及设计结果
    ret1=sap_object.SapModel.Analyze.DeleteResults(" ", True)
    #if ret1 != 0:
        #print("Analyze_Results not deleted successfully.")
    #else:
        #print("Analyze_Results deleted successfully.")


    ret2=sap_object.SapModel.DesignSteel.DeleteResults()
    #if ret2 != 0:
        #print("Design_Results not deleted successfully.")
    #else:
        #print("Design_Results deleted successfully.")
    
    # 删除均布面力
    for load in range(num_loc):
        if flattened_distribution[load] == 1:
            load_name = str(load + 1)

            ret=sap_object.SapModel.AreaObj.DeleteLoadUniform(load_name, 
                                                              "DEAD", 
                                                              0)

            #ret= sap_object.SapModel.AreaObj.DeleteLoadUniform(load_name, "DEAD", 0)
            #if ret != 0:
                #print("Area load not deleted successfully.")
            #else:
                #print("Area load deleted successfully.")

    #再次解锁模型，确保模型解锁
    SapModel.SetModelIsLocked(False)
    
print("**************** ALL process done ******************")

# 关闭 SAP2000
#sap_object.ApplicationExit(False)
