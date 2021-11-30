# MICDANA (MIC DAta ANAlyser) 

### Automate the data analysis of your MICs
MICDANA is a custom GUI designed to automate the tedious process of cleaning, manipulation and analysis of data from an excel file that 
the software SoftMax Pro v7.0. exports after the data collection in our experiment.

<p align="center">
<img src="https://user-images.githubusercontent.com/51713466/144073450-2db53114-eb01-4d35-bfd4-1c21ccc0b1cb.png" width="620.3" height="306.7" />
</p>

### Data preparation

1. To be able to clean the data, the information should be written in `SoftMax Pro v7.0` as follows:

> * One bacteria per plate.
> * Two antimicrobials per plate.    
> * Concentration in decreasing order, starting from the left side with the highest concentration in triplicates.

<p align="center">
<img src="https://user-images.githubusercontent.com/51713466/144073727-ca385711-a754-4ce7-ba33-07a80544f00d.png" width="588.7" height="426.7" />
</p>

2. Once we saved the data in the correct format, we need to export it as follows:

<p align="center">
<img src="https://user-images.githubusercontent.com/51713466/144073897-10faf00d-251d-4432-a3f0-721d5e2dbe6e.png"/>
</p>

3. After export the data to an excel file, it is necessary to convert the data to a ***.csv*** file to be able to read the information and to start running the script.

## Requierements
#### Software requirements:
> * SoftMax Pro v7.0  
> * Python v3.8.2
> * Pandas v1.0.5
> * Numpy 1.10.0
> * Matplotlib 3.2.2
> * XlsxWriter 1.2.9
> 
#### Hardware requirements:
> * Spectramax i3x

## Run the program
To be able to run the program, open the command line and type:

 `python MICDANA.py`
 
 
 
