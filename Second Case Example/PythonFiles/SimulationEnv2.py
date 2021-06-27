#################################################################################################################
# The "os" library is used to direct Python to the location of the ASPEN File.
import os
# The "win32com.client" library is used as an alternative to VBA  for the communication between ASPEN+ and Python.
import win32com.client as win32
# The "gym" is imported since it facilitates the creation of the Simulation Environement.
from gym import Env, spaces
# "Numpy" is used as a mathematical extention to Python.
import numpy as np
#################################################################################################################

# STEP 1. Initialise Environement and Prerequisite Variables.


class Simulator(Env):
    # The "__init__" function is used to provide the simulation all the variables it needs to initialise.
    def __init__(self):
        self.PATH = 'C:/Users/s2199718/Desktop/Second Case Example/AspenSimulation/SimulationCaseFile.bkp'
        # We define the document type of the Aspen+ File.
        self.AspenSimulation = win32.gencache.EnsureDispatch("Apwn.Document")
        self.AspenSimulation.InitFromArchive2(os.path.abspath(self.PATH))    # Path to File
        self.AspenSimulation.Visible = False    # Not opening the Simulation expedites training

        ## We define the variables needed for the DQNAgent
        self.action_space = spaces.Discrete(5)  # The Agent has 2 options [CSTR, PFR]
        self.observation_space = spaces.Box(low=0, high=1, shape=(5,))  # CONVERSION Ɐ [0, 1]
        self.TEMP_CHANGER = ["TC1","TC2","TC3","TC4","TC5","TC6"]
        self.CONVERSION_LIST = [0]
        self.Get_Input_Temp = np.array([350,350,350,350])/600
        self.CONVERSION_STATE = np.array([int(self.CONVERSION_LIST[-1])])
        self.STATE = np.concatenate((self.CONVERSION_STATE,self.Get_Input_Temp))
        # The state of the Env equals to 0 at the start of the 
        self.DONE_COUNTER = 0    # Cycle counter used to make the program more light-weight
        self.N_STEPS = 0    # 0 Steps taken when the Simulation is initialised

        ## We define the variables needed for the Simulation
        self.CHEM = ["SO2", "SO3", "O2", "N2"] # Names of the Chemicals
        self.IN_FLOW = [16.7878, 0, 23.0833, 169.977]  # Initial Flowrates [kmol/s]
        self.T_IN = 350
        self.T_max = 600
        self.T_min = 300
        self.CHOICE_MEMORY = []  # All the choices made by the Agents
        self.CHOICE_MATRIX = []
        self.FEED_MEMORY = ["FEED"]   # The Feed Memory, used for connecting Streams to Blocks
        self.Name_BLK_Output = "FEED"
           # Short-term memory of Conversion at Block Output
        self.MAX_CONVERSION = 0 # The maximum reward is 0 at Initialisation
        self.BEST_CASE = []
        self.CONVERSION_MATRIX = [] 

        ## We define the Stream and Block Names


#################################################################################################################

# STEP 2. Implement the Step function for the Environement

    def step(self,action):

        ## First we need to update some variables 
        self.Feed_Stream_Name = self.FEED_MEMORY[-1] # Defines the Input Name
        
        ## The Agent makes its Choice and The Simulation is Updated
        self.Agent_Makes_Choice(action) # Find Custom Function Below

        ## Calculate the Reward as a function of Conversion
        REWARD = self.REWARD_SIGNAL 
        self.Get_Input_Temp = np.array(self.GET_FINAL_TEMP())/600
        self.CONVERSION_STATE = np.array([self.CONVERSION_LIST[-1]])
        self.STATE = np.concatenate((self.CONVERSION_STATE,self.Get_Input_Temp))
        if action == 0:
            self.CONVERSION_LIST.append(self.CONVERSION)
            #print(self.CONVERSION_LIST)
        else:
            pass

        ## Checkpoint: Is the Cycle Done? What is the best Case?
        if self.N_STEPS == 4: # Max Number of Reactors = 4
            done = True 
            self.CONVERSION_MATRIX.append(self.CONVERSION_LIST[-1])
            self.CHOICE_MATRIX.append(self.CHOICE_MEMORY)
            TC_Temp_End = self.GET_FINAL_TEMP()
            FORM_CONV = "{:.2f}".format(self.CONVERSION_LIST[-1])
            #print(f"CONV: {self.CONVERSION_LIST}")
            print(f" TC_TEMP: [{TC_Temp_End}||{FORM_CONV}]")
            ## Keep Track of the Best Solutions 
            if self.CONVERSION_LIST[-1] > self.MAX_CONVERSION:
                self.MAX_CONVERSION = self.CONVERSION_LIST[-1]
                self.BEST_CASE.append([TC_Temp_End,self.MAX_CONVERSION,self.DONE_COUNTER])
            else:
                pass
            ## To make the Simulation more light-weight we hard-Reset it every 100 Cycles.
            if (self.DONE_COUNTER%100)==0:
                self.AspenSimulation.Close()    # Close the Simulation
                # Completely restart the simulation Like Step 1
                self.AspenSimulation = win32.gencache.EnsureDispatch("Apwn.Document")
                self.AspenSimulation.InitFromArchive2(os.path.abspath(self.PATH)) 
                self.AspenSimulation.Visible = False
                self.AspenSimulation.Engine.Run2()
                #print(f"~ASPEN+ Restarted {self.DONE_COUNTER}~")
            else:
                pass
            self.DONE_COUNTER += 1 # End of one Full Cycle
        else:
            done = False 
        return self.STATE, REWARD, done, {}

##################################################################################################################

# STEP 3. Define the Agent Make Choice Function


    def Agent_Makes_Choice(self, action):
        if action == 0: 
            self.CHOICE_MEMORY.append(f"MOVE {self.N_STEPS}->{self.N_STEPS+1}")
            self.N_STEPS += 1   # The Step counter is updated
            self.Name_BLK = "R" + str(self.N_STEPS)  # Eg. Name_BLK = R1, R(Reactor)1(STEP)
            self.Name_BLK_Input = "S"+str(self.N_STEPS)+"IN" # Preset Block Input Eg. S1AIN
            self.Name_BLK_Output = "S" + str(self.N_STEPS)+ "OUT" # Preset Block Input Eg. S1AOUT
            self.FEED_MEMORY.append(self.Name_BLK_Output) # The Output of this step is the Input of the next.
            self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
            ## Next the needed Simulation are acquired
            self.REACTANT_OUTPUT = self.Get_Output(self.Name_BLK_Output) # Amount of Reactant at Blocks Output [kmol/s]
            self.CONVERSION = self.Get_Conversion(self.REACTANT_OUTPUT) # Amount of Reactant Converted at Output [0, 1]
            self.REWARD_SIGNAL = self.CONVERSION - self.CONVERSION_LIST[-1] 
        elif action == 1:
            self.CHOICE_MEMORY.append(f"T Change in {self.TEMP_CHANGER[self.N_STEPS]} + 5")
            new_temp = self.CHANGE_TEMP(self.TEMP_CHANGER[self.N_STEPS],5)
            PENALTY = 0
            if new_temp > self.T_max:
                self.RESET_TEMP(self.TEMP_CHANGER[self.N_STEPS],self.T_max)
                PENALTY = - 1
            else:
                pass
            self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
            self.CONVERSION = 0
            self.REWARD_SIGNAL = PENALTY
        elif action == 2:
            self.CHOICE_MEMORY.append(f"T Change in {self.TEMP_CHANGER[self.N_STEPS]} - 5")
            new_temp = self.CHANGE_TEMP(self.TEMP_CHANGER[self.N_STEPS],-5)
            PENALTY = 0
            if new_temp < self.T_min:
                self.RESET_TEMP(self.TEMP_CHANGER[self.N_STEPS],self.T_min)
                PENALTY = - 1
            else:
                pass
            self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
            self.CONVERSION = 0            
            self.REWARD_SIGNAL = PENALTY
        elif action == 3:
            self.CHOICE_MEMORY.append(f"T Change in {self.TEMP_CHANGER[self.N_STEPS]} + 10")
            new_temp = self.CHANGE_TEMP(self.TEMP_CHANGER[self.N_STEPS],+10)
            PENALTY = 0
            if new_temp < self.T_min:
                self.RESET_TEMP(self.TEMP_CHANGER[self.N_STEPS],self.T_min)
                PENALTY = - 1
            else:
                pass
            self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
            self.CONVERSION = 0            
            self.REWARD_SIGNAL = PENALTY
        elif action == 4:
            self.CHOICE_MEMORY.append(f"T Change in {self.TEMP_CHANGER[self.N_STEPS]} - 10")
            new_temp = self.CHANGE_TEMP(self.TEMP_CHANGER[self.N_STEPS],-10)
            PENALTY = 0
            if new_temp < self.T_min:
                self.RESET_TEMP(self.TEMP_CHANGER[self.N_STEPS],self.T_min)
                PENALTY = - 1
            else:
                pass
            self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
            self.CONVERSION = 0            
            self.REWARD_SIGNAL = PENALTY



##################################################################################################################

# STEP 4. Define the Reset Function for the Environment.

    def reset(self):
        # Now all Variables that need to be reset Every Cycle are Reset
        self.N_STEPS = 0    # reset the number of steps taken (inside a cycle (max=4))
        self.CHOICE_MEMORY = [] # reset the choice memory of the cycle
        self.CONVERSION_LIST = [0] # reset the convrsion list of the cycle
        self.FEED_MEMORY = ["FEED"] # the first feed is set again to be S1
        self.Name_BLK_Output = "FEED"
        self.Get_Input_Temp = np.array([350,350,350,350])/600
        self.CONVERSION_STATE = np.array([0])
        self.STATE = np.concatenate((self.CONVERSION_STATE,self.Get_Input_Temp))
        self.Reset_Temp()
        self.CONVERSION = 0
        self.REWARD_SIGNAL = 0 # Remake all the streams to how they were before
        return self.STATE 

##################################################################################################################

# STEP 5. Implement all Auxiliary Functions that are needed for the previous steps.
# All functions presented below, are writen in the same order as they were used in the Steps above.

# STEP 5.1. Get the Output [kmol/s] of Reactant
    def Get_Output(self,Name_BLK_OUT):
        # We Navigate to the ASPEN Node containing the Value of the Output Streams
        REAC_OUT = self.AspenSimulation.Tree.Elements("Data").Elements("Streams").Elements(Name_BLK_OUT).Elements("Output").Elements("MOLEFLOW").Elements("MIXED").Elements(self.CHEM[0]).Value
        return REAC_OUT

# STEP 5.2. Calculate the Conversion at the Output of Every Block
    def Get_Conversion(self,REAC_OUT):
        CONV = (self.IN_FLOW[0] - REAC_OUT)/self.IN_FLOW[0]  # CONVERSION Ɐ [0, 1]
        return CONV

# STEP 5.4. Define the Reset Streams Function

    def Reset_Temp(self):
        self.BLK = self.AspenSimulation.Tree.Elements("Data").Elements("Blocks") # Find the ASPEN Node for all Streams
        self.BLK.Elements(self.TEMP_CHANGER[0]).Elements("Input").Elements("TEMP").Value = self.T_IN
        self.BLK.Elements(self.TEMP_CHANGER[1]).Elements("Input").Elements("TEMP").Value = self.T_IN
        self.BLK.Elements(self.TEMP_CHANGER[2]).Elements("Input").Elements("TEMP").Value = self.T_IN
        self.BLK.Elements(self.TEMP_CHANGER[3]).Elements("Input").Elements("TEMP").Value = self.T_IN

    def CHANGE_TEMP(self,Name_Temp_Changer,Temperature_Change):
        current_Temp = self.BLK.Elements(Name_Temp_Changer).Elements("Input").Elements("TEMP").Value
        new_Temp = current_Temp + Temperature_Change
        self.BLK.Elements(Name_Temp_Changer).Elements("Input").Elements("TEMP").Value = new_Temp
        return new_Temp
    def RESET_TEMP(self,Name_Temp_Changer,Temperature):
        self.BLK.Elements(Name_Temp_Changer).Elements("Input").Elements("TEMP").Value = Temperature

    def GET_FINAL_TEMP(self):
        self.BLK = self.AspenSimulation.Tree.Elements("Data").Elements("Blocks")
        T1 = self.BLK.Elements(self.TEMP_CHANGER[0]).Elements("Input").Elements("TEMP").Value
        T2 =self.BLK.Elements(self.TEMP_CHANGER[1]).Elements("Input").Elements("TEMP").Value 
        T3 = self.BLK.Elements(self.TEMP_CHANGER[2]).Elements("Input").Elements("TEMP").Value 
        T4 =self.BLK.Elements(self.TEMP_CHANGER[3]).Elements("Input").Elements("TEMP").Value 
        return [T1, T2, T3, T4 ]