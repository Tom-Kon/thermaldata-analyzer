<h1>README</h1>

<p>The code presented in this repository can be used to analyze thermal data. Think of differential scanning calorimetry, themogravimetric analysis or dynamic mechanical analysis. These techniques are relevant for material science, pharmaceutical research, and the food industry to name a few. </p>
<p>The code consists of a ShinyR app that focuses on user-friendlieness. The target audience are people that have no experience whatsoever with code. The app requires analyses to be run in TRIOS software beforehand. TRIOS software, produced by TA instruments, is one of the most commonly used pieces of software to analyze thermal data. Inside of the app, a tutorial can be found explaining how to generate the files required for the code in TRIOS. The tutorail also contains some more information on how to use the app. 
Aside of the heavy focus on user-friendlieness, the code also contains internal checks to ensure that the user input is correct and internally consistent. It can analyze different data structures (see the "examples" folder). </p>

The repository contains:
<ol>
  <li>general_data_analyzer_integrated.R: this is the code (app). </li>
  <li>paper.md: this is the paper submitted to the journal of open source software. The figures it contains are in the "figures" folder. </li>
  <li>The code can be verified by running the files in the "examples" folder. The same folder also contains an example of an expected output excel. </li>  
</ol>

