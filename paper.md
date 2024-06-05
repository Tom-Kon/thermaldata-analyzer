---
title: 'The thermal data analyzer'
tags:
  - Rstudio
  - ShinyR
  - Thermal data
  - Differential scanning calorimetry
  - TRIOS
  - Dynamic mechanical analysis
  - Thermogravimetric analysis
authors:
  - name: Tom Konings
    orcid: 0000-0003-1256-6557
    affiliation: 1
  - name: Julia Bandera
    orcid: 0009-0000-1104-7232
    affiliation: 2
affiliations:
 - name: Department of pharmaceutical and pharmacological sciences, Herestraat 49 box 921, University of Leuven (KU Leuven), 3000 Leuven, Belgium 
   index: 1
 - name: Faculty of Bioscience Engineering, Kasteelpark Arenberg 20, University of Leuven (KU Leuven), 3001 Leuven, Belgium
   index: 2
date: Submitted to the journal of open source software on 4th June 2024
bibliography: paper.bib
---

# Summary
Thermal and rheological analysis is used in many scientific fields and industries, including the food- and pharmaceutical industries [@Seifi:2020; @Panzade:2024]. Common techniques include dynamic mechanical analysis [@Jones:2012], thermogravimetric analysis, (modulated) differential scanning calorimetry [@Liu:2022; @Jurczak:2020], and so forth. Such analyses give the researcher access to essential material properties such as the viscosity, the glass transition, and other events such as melting, just to name a few. One important supplier of instruments capable of such analyses is TA instruments. Their most recent software, TRIOS [@TRIOS:2024], facilitates instrument control and data analysis with user-friendly features. However, it has limitations in terms of customizable and automatic data analysis, requiring supplementary software packages.

# Statement of need
TRIOS does allow users to save analysis templates and report templates for rapid analysis of similar samples6. It also permits exporting data to Excel, but this process is less automated and customizable [@TRIOS:2024]. Consequently, issues arise when analyzing similar samples in scientific research contexts, particularly when dealing with replicates (runs) as TRIOS treats them as entirely separate files. This complicates the automatic export of data from different runs to Excel. Furthermore, difficulties are exacerbated when researchers need to calculate descriptive statistics such as mean, standard deviation, and relative standard deviation, despite these being straightforward operations.

The abovementioned limitation can, of course, be addressed by a program. However, researchers from a pharmaceutical background may not possess sufficient coding expertise to develop software even for simple analyses. This is why the thermal data analyzer was created. It consists of a user interface built using ShinyR (with code written in R). The program relies on users generating multiple Word documents containing their analysis data from several replicates. These documents can be automatically generated within the TRIOS software, and the ShinyR app includes a tutorial (Figure 2) on how to accomplish this. The entire procedure can be automated, with all outputs customizable to suit the user's requirements (Figure 1). Moreover, the app includes several internal checks to verify data entry accuracy, issuing clear error statements if any mistakes are detected. The app generates Excel documents where means, standard deviations, and relative standard deviations are organized into formatted tables. It can also be used for extraction of raw data from Word documents for export into Excel files. User-friendliness, particularly for people without coding knowledge, is the most important aspect of the code, resulting in the user interfaces presented in Figures 1 and 2.

The app mainly focuses on analyzing differential scanning calorimetry (DSC) data, but this is visible only in the nomenclature of the output. However, it can be used for any data obtained in TRIOS, including data from different instruments. Data generated using older TA instruments can still be opened in TRIOS, meaning that the software presented in this paper is also compatible with these instruments.  The app is capable of processing nearly any data structure as input, with limitations outlined in the tutorial section (Figure 2) of the app.

<br>

<p align="center">
  <img src="https://github.com/Tom-Kon/thermaldata-analyzer/blob/main/Figures/figure%201%20main%20menu.png" width='60%'>
  <br>
  <em>Figure 1: The user interface for input of data. On the following page, users are prompted to provide inputs such as the files to analyze and the desired name for the output Excel file. Additionally, the button to initiate the analysis is located on  page 2.</em>
</p>

<br>
<br>

<p align="center">
  <img src="https://github.com/Tom-Kon/thermaldata-analyzer/blob/main/Figures/figure%202%20tutorial.png" width='60%'>
  <br>
  <em>Figure 2: The tutorial, also present in the app, gives detailed instructions regarding the input data and guidelines for creating necessary documents in TRIOS. Additionally, it outlines the program's limitations and offers a brief overview of the code's functionality.</em>
</p>

<br>



# Acknowledgements
The software presented in this paper was developed as part of a project funded by the Flemish fund for scientific research (project 1SH0S24N). The authors also acknowledge Prof. Guy Van den Mooter’s help in reviewing the manuscript. 


# References



