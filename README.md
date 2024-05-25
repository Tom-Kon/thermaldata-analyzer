The code presented in this repository can mainly be used to analyze thermal data. Think of differential scanning calorimetry, themogravimetric analysis or dynamic mechanical analysis. 
It requires analyses to be run in TRIOS software beforehand. The code consists of a ShinyR app that focuses on user-friendlieness. The target audience is people that have no experience whatsoever with code.
Inside of the app, a tutorial can be found explaining how to generate the files required for the code, as well as some more information how to use the app. 
Aside of the heavy focus on user-friendlieness, the code also contains internal checks to ensure that the user input is correct and internally consistent. It can analyze different data structures (see the "examples" folder). 

general_data_analyzer_integrated.R is the code. 

paper.md is the paper submitted to the journal of open source software. It contains several figures also present in the repository. 

The code can be verified by running the files in the "examples" folder. The same folder also contains an example of an expected output excel. 
