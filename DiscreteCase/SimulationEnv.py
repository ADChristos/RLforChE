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
        self.PATH = 'C:/Users/s2199718/Desktop/DiscreteCase/AspenSimulation/DiscreteExample.bkp'
        # We define the document type of the Aspen+ File.
        self.AspenSimulation = win32.Dispatch("Apwn.Document")
        self.AspenSimulation.InitFromArchive2(os.path.abspath(self.PATH))    # Path to File
        self.AspenSimulation.Visible = False    # Not opening the Simulation expedites training

        ## We define the variables needed for the DQNAgent
        self.action_space = spaces.Discrete(2)  # The Agent has 2 options [CSTR, PFR]
        self.observation_space = spaces.Box(low=0, high=1, shape=(1,))  # CONVERSION Ɐ [0, 1]
        self.STATE = 0  # The state of the Env equals to 0 at the start of the Simulation
        self.DONE_COUNTER = 0    # Cycle counter used to make the program more light-weight
        self.N_STEPS = 0    # 0 Steps taken when the Simulation is initialised

        ## We define the variables needed for the Simulation
        self.CHEM = ["N-BUT-01", "ISO-B-01", "2-MET-01"] # Names of the Chemicals
        self.IN_FLOW = [0.0099, 0.0001, 0.0769]  # Initial Flowrates [kmol/s]
        self.CHOICE_MEMORY = []  # All the choices made by the Agents
        self.FEED_MEMORY = ["S1"]   # The Feed Memory, used for connecting Streams to Blocks
        self.CONVERSION_LIST = [0]   # Short-term memory of Conversion at Block Output
        self.MAX_CONVERSION = 0 # The maximum reward is 0 at Initialisation
        self.BEST_CASE = []

        ## We define the Stream and Block Names
        self.STRM_INPUTS = ["S1AIN", "S2AIN", "S3AIN", "S4AIN", "S1BIN", "S2BIN", "S3BIN", "S4BIN"]
        self.BLK_NAMES = ["B1A", "B2A", "B3A", "B4A", "B1B", "B2B", "B3B", "B4B"]
        self.STRM_OUTPUTS = ["S1AOUT", "S2AOUT", "S3AOUT", "S4AOUT", "S1BOUT", "S2BOUT", "S3BOUT", "S4BOUT"]

#################################################################################################################

# STEP 2. Implement the Step function for the Environement

    def step(self,action):

        ## First we need to update some variables 
        self.Feed_Stream_Name = self.FEED_MEMORY[-1] # Defines the Input Name
        self.N_STEPS += 1   # The Step counter is updated
        
        ## The Agent makes its Choice and The Simulation is Updated
        self.Agent_Makes_Choice(action) # Find Custom Function Below
        self.AspenSimulation.Engine.Run2() # Run the ASPEN+ Simulation
        self.AspenSimulation.Engine.Run2() # Run the Simulation again to eliminate Errors

        ## Next the needed Simulation are acquired
        REACTANT_OUTPUT = self.Get_Output(self.Name_BLK_Output) # Amount of Reactant at Blocks Output [kmol/s]
        CONVERSION = self.Get_Conversion(REACTANT_OUTPUT) # Amount of Reactant Converted at Output [0, 1]
        ## Calculate the Reward as a function of Conversion
        REWARD = CONVERSION - self.CONVERSION_LIST[-1]
        self.STATE = np.array([self.CONVERSION_LIST[-1]])
        self.CONVERSION_LIST.append(CONVERSION)

        ## Checkpoint: Is the Cycle Done? What is the best Case?
        if len(self.CHOICE_MEMORY) == 4: # Max Number of Reactors = 4
            done = True 
            ## Keep Track of the Best Solutions 
            if self.CONVERSION_LIST[3] > self.MAX_CONVERSION:
                self.MAX_REWARD = self.CONVERSION_LIST[3]
                self.BEST_CASE.append([self.CHOICE_MEMORY,self.MAX_CONVERSION,self.DONE_COUNTER])
            else:
                pass
            ## To make the Simulation more light-weight we hard-Reset it every 100 Cycles.
            if (self.DONE_COUNTER%100)==0:
                self.AspenSimulation.Close()    # Close the Simulation
                # Completely restart the simulation Like Step 1
                self.AspenSimulation = win32.Dispatch("Apwn.Document")
                self.AspenSimulation.InitFromArchive2(os.path.abspath(self.PATH)) 
                self.AspenSimulation.Visible = False
                self.AspenSimulation.Engine.Run2()
                print(f"~ASPEN+ Restarted {self.DONE_COUNTER}~")
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
            ROUTE = "A" # Route(A) = CSTR
        else:
            ROUTE = "B" # Route(B) = PFR
        
        # Define the Names of the Block and Streams
        self.Name_BLK = "B" + str(self.N_STEPS) + ROUTE # Eg. Name_BLK = B1A, B(Block)1(STEP)A(CSTR)
        self.CHOICE_MEMORY.append(self.Name_BLK)
        self.Name_BLK_Input = "S"+str(self.N_STEPS)+ROUTE+"IN" # Preset Block Input Eg. S1AIN
        self.Name_BLK_Output = "S" + str(self.N_STEPS) + ROUTE + "OUT" # Preset Block Input Eg. S1AOUT
        self.FEED_MEMORY.append(self.Name_BLK_Output) # The Output of this step is the Input of the next.
        self.Connect_Feed() # We connect the Feed(step-1) to the chosen Block

##################################################################################################################

# STEP 4. Define the Reset Function for the Environment.

    def reset(self):
        # Now all Variables that need to be reset Every Cycle are Reset
        self.N_STEPS = 0    # reset the number of steps taken (inside a cycle (max=4))
        self.CHOICE_MEMORY = [] # reset the choice memory of the cycle
        self.CONVERSION_LIST = [0] # reset the convrsion list of the cycle
        self.FEED_MEMORY = ["S1"] # the first feed is set again to be S1
        self.STATE = np.array([0]) # reset the State 
        self.Reset_Streams() # Remake all the streams to how they were before
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

# STEP 5.3. Define the Connect Feed Function
    def Connect_Feed(self):
        BLK_INPUT = self.AspenSimulation.Tree.Elements("Data").Elements("Blocks") # Find the ASPEN Node for all Blocks
        STRM_INPUT = self.AspenSimulation.Tree.Elements("Data").Elements("Streams") # Find the ASPEN Node for all Streams
        STRM_INPUT.Elements.Remove(self.Name_BLK_Input) # Delete the Stream that is already connected to the Block
        BLK_INPUT.Elements(self.Name_BLK).Elements("Ports").Elements("F(IN)").Elements.Add(self.Feed_Stream_Name) # Connect the OutputStream[step-1]

# STEP 5.4. Define the Reset Streams Function

    def Reset_Streams(self):
        self.STRM = self.AspenSimulation.Tree.Elements("Data").Elements("Streams") # Find the ASPEN Node for all Streams
        self.STRM.RemoveAll() # This Deletes all the streams present in the Flowsheet
        # To re-create all the Deleted Streams the following 4-sub commands are used.
        self.Add_Input_Streams()
        self.Add_Output_Streams()
        self.Add_Feed_Stream()
        self.Connect_Streams()

# STEP 5.4.1. Create all the Input Streams -"S1"
    def Add_Input_Streams(self):
        for STRM_Name in self.STRM_INPUTS: # For every name in the Stream Inputs
            self.STRM.Elements.Add(STRM_Name) # Create a Stream with that Name and the following SPECS
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("TEMP").Elements("MIXED").Value = 298 # Stream Temp [K]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("PRES").Elements("MIXED").Value = 5e+06 # Pressure [N/m2]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("TOTFLOW").Elements("MIXED").Value = 0.0869 # Total Flow [kmol/s]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("FLOW").Elements("MIXED").Elements(self.CHEM[0]).Value = 0.0099 # [kmol/s]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("FLOW").Elements("MIXED").Elements(self.CHEM[1]).Value  = 0.0001 # [kmol/s]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("FLOW").Elements("MIXED").Elements(self.CHEM[2]).Value  = 0.0769 # [kmol/s]
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("NPHASE").Elements("MIXED").Value = 1 # Number of Phases
            self.STRM.Elements(STRM_Name).Elements("Input").Elements("PHASE").Elements("MIXED").Value = "L" # Chosen Liquid Phase

# STEP 5.4.2. Create All of the Output Streams 
    def Add_Output_Streams(self):
        for STRM_OUT in self.STRM_OUTPUTS: # For every name in the Stream Outputs
            self.STRM.Elements.Add(STRM_OUT) # create a stream with that name 

# STEP 5.4.3. Create The first Feed Stream [S1]
    def Add_Feed_Stream(self):
        FeedConfig = self.STRM.Elements.Add("S1") # Create "S1" and select the following SPECS
        FeedConfig.Elements("Input").Elements("TEMP").Elements("MIXED").Value = 298
        FeedConfig.Elements("Input").Elements("PRES").Elements("MIXED").Value = 5e+06
        FeedConfig.Elements("Input").Elements("TOTFLOW").Elements("MIXED").Value = 0.0869
        FeedConfig.Elements("Input").Elements("FLOW").Elements("MIXED").Elements("N-BUT-01").Value = 0.0099
        FeedConfig.Elements("Input").Elements("FLOW").Elements("MIXED").Elements("ISO-B-01").Value = 0.0001
        FeedConfig.Elements("Input").Elements("FLOW").Elements("MIXED").Elements("2-MET-01").Value = 0.0769
        FeedConfig.Elements("Input").Elements("NPHASE").Elements("MIXED").Value = 1
        FeedConfig.Elements("Input").Elements("PHASE").Elements("MIXED").Value = "L"

# STEP 5.4.4. Connect All Input and Output Streams to their Corresponding Blocks
    def Connect_Streams(self):
        BLK = self.AspenSimulation.Tree.Elements("Data").Elements("Blocks")
        for i in range(0,len(self.BLK_NAMES)): # for every Name in Blocks: Connect Input and Output 
            BLK.Elements(self.BLK_NAMES[i]).Elements("Ports").Elements("F(IN)").Elements.Add(self.STRM_INPUTS[i])
            BLK.Elements(self.BLK_NAMES[i]).Elements("Ports").Elements("P(OUT)").Elements.Add(self.STRM_OUTPUTS[i])