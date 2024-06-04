<h1>README</h1>

<p>The code presented in this repository can be used to analyze thermal data. Think of data generated by differential scanning calorimetry (DSC), themogravimetric analysis (TGA) or dynamic mechanical analysis (DMA). These techniques are relevant for material science, pharmaceutical research, and the food industry, to name a few. </p>
<p>The code consists of a ShinyR app that focuses on user-friendlieness. The target audience are people that have no experience whatsoever with code. The app requires analyses to be run in TRIOS software beforehand. TRIOS software, produced by TA instruments, is one of the most commonly used pieces of software to analyze thermal data. Inside of the app, a tutorial can be found explaining how to generate the files required for the code in TRIOS. The tutorial also contains some more information on how to use the app. 
Aside of the heavy focus on user-friendlieness, the code also contains internal checks to ensure that the user input is correct and internally consistent. It is capable of analyzing different data structures (see the "examples" folder), meaning that any of the data generated in TRIOS, following a few minor constraints, will work. </p>

The repository contains:
<ol>
  <li>general_data_analyzer_integrated.R: this is the code (app). </li>
  <li>paper.md: this is the paper submitted to the journal of open source software. The figures it contains are in the "figures" folder. </li>
  <li>The code can be verified by running the files in the "examples" folder. Different examples are presented in different subfolders, all having different datastuctures. The subfolder of each particular example also contains an expected output excel. </li>  
</ol>

In case you've landed here but do not know how GitHub works, here is the online version of the code: https://posit.cloud/content/7489585. In the panel on the right, select the file with the .R extension. Click "Run App" in the top of the main screen. You'll be able to access the app's tutorial in this way and try out the online version. 
