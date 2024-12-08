import comtypes.client
import random
import numpy as np
import os

############################ Introduction ##############################
# Run this code while running SAP2000 project
# Sap2000 V23.3.1 is recommended
# You can find every api used in this program in CSI_OAPI_Documentation, 
# which is in the SAP2000 installation directory
# Version: 2.0
########################################################################


## Create 6 folders to store datas and labels for frequency, stress ratio, and inter-story displacement

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


# Connect to the running instance of SAP2000
sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
SapModel = sap_object.SapModel

# Unlock the sap2000 model
SapModel.SetModelIsLocked(False)

# Set the model units to N/mm^2
SapModel.SetPresentUnits(9)  # 9 for N/mm^2


#Randomly distribute goods in the warehouse
def random_goods_distribution(num_z, num_x, num_y):
    # create a num_z * num_x * num_y matrix to store the distribution of goods
    distribution_matrix = []
    
    for layer in range(num_z):
        # Randomly distribute goods in each layer,1 for goods, 0 for no goods
        distribution = [random.choice([0, 1]) for _ in range(num_x * num_y)]
        distribution_matrix.append(distribution)  # append the distribution of each layer
        
    return distribution_matrix


############# configurations #############
num_trials = 5 # epoch for data generation
surface_pressure_area = 0.001  # pressure for each area
num_z = 3 #  3 for real , 3 for test
num_x = 11 # 11 for real , 7 for test
num_y = 14 # 14 for real , 7 for test


# Get the frame information 
frame_info=sap_object.SapModel.FrameObj.GetNameList()

#if frame_info != 0:
    #print("GetNameList not successfully.")
#else:
    #print("GetNameList successfully.")

# Get the frame name list   
frame_name=frame_info[1]
#print(frame_name)


#Set area load  for each area according to the goods distribution matrix and run the analysis and design

print("**************** Start process ******************")
#print('—————————————————————————————————————————————————————————————————————————————————————————————————————————')

num_loc = 462  # Total goos position, storage&AGVway included (147 for test   462 for real)

for trial in range(num_trials):
    # Randomly distribute goods in the warehouse
    distribution_matrix = random_goods_distribution(num_z, num_x, num_y)
    
    #print(f"Trial {trial + 1}: Goods Distribution Matrix:")
    #for layer in distribution_matrix:
        #print(layer)

    # Flatten the distribution matrix, 1 for goods, 0 for no goods
    flattened_distribution = [item for sublist in distribution_matrix for item in sublist]

    # Set the area load for each area according to the goods distribution matrix
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
            

    # Run the analysis 
    ret=SapModel.Analyze.RunAnalysis()
    #if ret != 0:
        #print("Analysis did not run successfully.")
    #else:
        #print("Analysis run successfully.")

    # Run the design
    ret=SapModel.DesignSteel.StartDesign()
    #if ret != 0:
        #print("Design did not run successfully.")
    #else:
        #print("Design run successfully.")
    


######################### Get first 12 circ_frequences of the structure ####################

    NumberResults = 12
    LoadCase = []
    StepType = []
    StepNum = []
    Period = []
    Frequency = []
    CircFreq = []
    EigenValue = []

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

    #print(f"Trial {trial + 1} CirFreq: {result[6]}")  # 6 for CircFreq

    np.save(os.path.join(CirFre_Data,f'data_cirfre_{trial}'), distribution_matrix)
    np.save(os.path.join(CirFre_Label,f'label_cirfre_{trial}'), result[6])



########################### Get the stress ratio data ############################

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
            
        #The output may be "GetSummaryResults failed." , however, the data is still gotten successfully
        #if ret != 0:
            #print("GetSummaryResults failed.")
        #else:
            #print("GetSummaryResults succeeded.")
    

    #print(f"Trial {trial + 1} max ratio: {max_ratio}")

    np.save(os.path.join(Ratio_Data,f'data_ratio_{trial}'), distribution_matrix)
    np.save(os.path.join(Ratio_Label,f'label_ratio_{trial}'), max_ratio)   
    

########################### Get the interlayer displacement (angle) data ############################
# Generalized dispacement need to be set in Sap2000 model, this code use corner point of the frame to get the displacement
# For this case, the frame is 3 floors, rad_1 stands for the displacement angle of the first floor,etc.

    num_floor = 3
    result_rad=[]
    for floor in range(num_floor):
        Name=f"rad_{floor}"  # Name of the generalized displacement item, for this case: rad_1, rad_2, rad_3
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

        #The output may be "GeneralizedDispl failed." , however, the data is still gotten successfully
        #if result_displ != 0:
            #print("GeneralizedDispl failed.")
        #else:
            #print("GeneralizedDispl succeeded.")  
             
    #print(f"Trial {trial+1} DisplRad: {result_rad}")

    np.save(os.path.join(Displ_Data,f'data_displ_{trial}'), distribution_matrix)
    np.save(os.path.join(Displ_Label,f'label_displ_{trial}'), result_rad)
    


    #print('—————————————————————————————————————————————————————————————————————————————————————————————————————————')

     #Unlock the model
    SapModel.SetModelIsLocked(False)

    # Delete the analysis and design results
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
    

    # Delete the area load for each area
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


    #Unlock the model again, in case the model is locked
    SapModel.SetModelIsLocked(False)
    
print("**************** ALL process done ******************")

# Close SAP2000 
#sap_object.ApplicationExit(False)
