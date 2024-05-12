---
title: 'The thermal data analyzer'
tags:
  - Rstudio
  - ShinyR
  - Thermal data
  - Differential scanning calorimetry
  - TRIOS
authors:
  - name: Tom Konings
    orcid: 0000-0003-1256-6557
    affiliation: 1
  - name: Julia Bandera
    orcid: 
    affiliation: 2
affiliations:
 - name: Department of pharmaceutical and pharmacological sciences, Herestraat 49 box 921, Catholic University of Leuven, 3000 Leuven, Belgium 
   index: 1
 - name:
   index: 2
date: 
bibliography: paper.bib
---

# Summary
Thermal analysis is used in many scientific fields and industries, including the food- and pharmaceutical industries. Common techniques include dynamic mechanical analysis, rheology studies, differential scanning calorimetry, and so forth. Such analyses give the researcher access to essential material properties such as the viscosity, the glass transition, and other events such as melting, just to name a few. One of the most important suppliers of instruments capable of such analyses is TA instruments. The most recent software for controlling all of these instruments as well as performing data analysis is TRIOS. TRIOS is user-friendly, but has some limiations regarding customizable and automatic data analysis that call for an additional software package. 

# Statement of need
TRIOS does allow the user to save analysis templates and report templates for rapid analysis of similar samples. It also allows for exporting data to Excel, but this process is less automated and less customizable. This results in problems when analyzing similar samples in the context of scientific research. This is especially true when replicates (runs) are involved, since TRIOS treats those as completely separate files. This makes automatically exporting data from different runs to Excel convoluted. This is exacerbated when researcher wishes to calculate descriptive statistics such as the mean, standard deviation and relative standard deviation of a sample, even though these are very simple operations. 

The abovementioned limitation can of course be solved by a program, but researchers that come from a pharmaceutical background might also not have enough coding know-how to write a piece of software that can do simple analyses for them. That is why the thermal data analyzer was developed. It consists of a user-friendly user interface, made using ShinyR (the code is written in R). It relies on the user generating several word documents containing their analysis data from several replicates. These documents can be generated automatically in the TRIOS software, and the ShinyR app contains a tutorial on how to do this. This whole process can be automized, all outputs are customizable to fit the user's needs, and it doesn't require any knowledge of code. Moreover, the app contains several internal checks to ensure that the user didn't make any mistakes while putting in their data, which result in clear error statements if something went wrong. The app outputs Excel documents, where means, standard deviations and relative standard deviations are grouped into formatted tables. It can also be used to extract the raw data from word documents and export it to the Excel files. 

The app mainly focuses on analyzing differential scanning calorimetry (DSC) data, but this is visible only in the nomenclature of the output. It can be used for any data obtained in TRIOS, including data from different instruments. Data generated using older TA instruments can still be opened in TRIOS, meaning that the software presented in this paper is also compatible with these machines. The app is able to take almost any data structure as input, facing few constraints that are stated in the tutorial section of the app. 

# Mathematics
Formulas used in the code when at least 3 replicates are present are the standard formulas for calculating the mean ($\overline{x}$), standard deviation (s) and relative standard deviation (RSD) of a dataset with n observations ($x_i$). 

$$\overline{x} = \sum_{i=1}^{i=n} \ x_i * \frac{1}{n}$$


$$ s = \sqrt{\sum_{i=1}^{i=n} \ (x_i - \overline{x})^2 * \frac{1}{n-1}}$$


$$ RSD = \frac{s}{\overline{x}} * 100 \\% $$

In case only duplicates were performed, the spread and relative spread are calculated instead of the standard deviation and relative standard deviation. 

$$ spread = |x_1 - x_2| $$

$$ relative \ spread = \frac{spread}{\overline{x}} * 100 \\% $$

# Acknowledgements
The software presented in this paper was developed as part of a project funded by the Flemish fund for scientific research (project 1SH0S24N). 


# References
