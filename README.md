# Imankulov_MDP

About
Python script for calculating maximum transmission power flow in branch group of power grid according with Russian PSO criterias (rus).

Solution is based on calculation engine Astra from the RastrWin3 library collection.

Requirments:

 - Python (x32) 3.7 (or upper)

 - Pywin32 installed in python enviroment

 - Installed RastrWin3 with registered COM AstraLib 1.0

Model of powergrid should be described in RastrWin3 regime.rg2 file
Faults and flowgate should be described in .json file
Trajectory of regime should be described in .csv file

RastrWin3 templates should be described in RastrWin3 appropriate file formats. They can be found in YOUR_PATH\RastrWin3\RastrWin3\SHABLON
