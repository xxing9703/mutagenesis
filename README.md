# Mutagenesis batch tool
This is a batch tool for creating mutagenesis DNA sequences and display them in excel with customer prefered format.
Provided a parent DNA sequence, customers requested to do multiple mutagenesis tasks. Each task can be a modification on a single or multiple sites, coded as "G118R/F121Y/R263K" for example, meaning that simultaneously turn G into R on site 118, turn F into Y on site and turn R into K on site 263. the DNA sequences after mutations will be exported to an excel file with mutated sites highlighted to ensure that it is done correctly.

This code will first verify that each requested mutagenesis command is valid and "codon2.txt" is used for validation purpose only (e.g, for G118R, site 118, it's 'G'	= GGT, GGA, GGG or GGC).  Then, mutagenesis is performed based on 'codon1.txt' codon table.

An example of the exported excel format is shown below.

![image](https://user-images.githubusercontent.com/16364863/143309385-6e797d1d-13a4-4af6-b731-9c44b5ef5bc1.png)

## Required inputs and example files:
1. parent DNA sequence (parent2.txt)
2. codon1.txt in txt (for mutagenesis)
3. codon2.txt in txt (for validation)
4. task command list in txt (tab2.txt)

## Output
an excel file (tab2cl.xlsx)

## running the code
Simply run script "code.m" in matlab. Make changes of the input files for a different task.
