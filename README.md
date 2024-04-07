# Power-Theft-Mapper

--------- Files ---------
main.py - Script for calculation of theft location, plotting it and sending via email.
Power Theft Location.slx - Simulation file, for MATLAB Simulink 
sim2xl.m - Code to transfer Simulink readings to an Excel Workbook

--------- Project Description ---------
In a scenario where overhead transmission lines are used to transmit electricity, power theft can easily occur by connecting an external wire to a low voltage transmission line. In that case, we can find out the exact location of the power theft and report it to the respective authorities.

To achieve this we are assuming a scenario where a low voltage transmission line is transmitting power to a smart neighborhood, where the power usage can be monitored. We aim to take the current and voltage readings from these smart homes/consumers to analyse how much power is being consumed.

At the same time, current and voltage readings are being taken from the secondary side of the tranformer, which tells us how much power is being transmitted from the source.

Using all these readings, we can check if the amount of power that is being transmitted matches the amount of power that is being used by the consumers.
After factoring in transmission losses, if the amount of energy being transmitted exceeds the amount of energy that is being used, we can conclude that a theft has occured.

The exact location of power theft can be calculated by using all the readings and the length of the transmission line. The same can be plotted on a map and sent to the concerned authorities for taking action.


We have created a simulation of this whole process, which stores the readings in an excel sheet. The main code extracts data from the excel, performs calculations and created a map.html which is sent to the authorities by email.
