#Mutation conservation across organisms and chemical modifier screens in ALS yeast 

##Mutation Conservation 

####Summary of current data
Genes: 260

For latest data and figure see the folder named conservation_latest_data

Throughout all methods pathogenic mutations are on average more conserved than the average amino acid in the protein (Jensen Shannon Divergence p: 7.78E-03, Shannon Entropy p: 1.73E-05, Sum of Pairs p: 1.68E-13), and more conserved than benign mutations (Jensen Shannon Divergence p: 1.05E-26, Shannon Entropy p: 1.55E-17, Sum of Pairs p: 2.98E-27)

The average number of mutations per gene is 11.59459459, and 19% are benign. On average 5.24609375 mutations are fully conserved which make up about 3.4% of the fully conserved amino acids on the proteins.

Notably, the difference in conservation between pathogenic mutations and the average amino acid seems to be more significant when we have more mutations (pearson correlation: -0.361300), and the relationship appears to follow a power trend. 


####(folder: yeast_replaceable_genes)
This project started as a follow up to a paper from the Marcotte lab in which they replaced 414 essential yeast genes with their human counterparts and found that in about half of these strains the genes were "replaceable", meaning that the yeast was able to survive with the human version of the gene. The idea was then to find which of these genes were associated with Mendelian diseases and/or Monogenic diseases and based on sequence alignments between yeast, worm, fly, zebrafish, mouse and human, figure out if disease-causing mutations affect conserved or variable amino acid positions (more background info here: http://mtc.science/humanization-of-yeast-genes) 

As a way to test the pipeline I tried to figure out if disease causing mutations affect more conserved amino acids in NPC1, which is the gene that causes Niemann Pick Type C. I got the amino acid sequences from entrez protein and performed an alignment with CLUSTAL W2. Then I got a list of mutations in the NPC1 gene from ClinVar. I got a total of 51 amino acid changes, 6 of which were benign/likely benign. I then searched for the amino acid in which the mutations occur in the sequence alignment and saw whether the amino acid was conserved or not. To see if the amino acids were conserved I used 3 different methods: Jensen Shannon Distribution, Shannon Entropy and Sum of Pairs (more information and results here: http://mtc.science/mutation-conservation-across-organisms-in-npc1)

Later I performed the analysis with the genes that were found to be replaceable and non replaceable to see if this was the case with these genes as well or not and if there was a difference between the two groups. Halfway through this I figured out a way to automate more of the process with a python script which you can see here: https://github.com/materechm/plabData/blob/master/yeast_replaceable_genes/test.py (note - on the script the red coloring based on the amino acid properties DOES NOT work) You can see results here: http://mtc.science/humanization-of-yeast-genes-part-2

I am now in the process of performing this analysis in a set of 260 genes that cause mendelian diseases that involve all organelles. Specific gene data is under sequences, then that is divided into replaceable, non replaceable, and other mendelian. When you go to a specific gene you can see a summary of the mutation conservation for the gene, protein sequences, protein alignment, and the conservation scores. You can see the consolidation of the data here: https://github.com/materechm/plabData/blob/master/yeast_replaceable_genes/sequences/all_other_mendelian/mendelian_conservation_summary.xlsx and a blog post sumarizing the data for the first 130 genes here: http://mtc.science/mutation-conservation-in-mendelian-genes

####Links to external tools used 
- Homologene: http://www.ncbi.nlm.nih.gov/homologene
- Oma Browser: http://omabrowser.org/oma/home/
- ClustalW2 (retired): http://www.ebi.ac.uk/Tools/msa/clustalw2/
- ClustalOmega: http://www.ebi.ac.uk/Tools/msa/clustalo/
- ClinVar: http://www.ncbi.nlm.nih.gov/clinvar/
- Conservation Scoring (developed by Mona Singh's group at Princeton) http://compbio.cs.princeton.edu/conservation/

##Screens 
####(folder: yeast) 
The folder currently contains images of the yeast were you can see the gene induction (because it is tagged to YFP). There are two folders, each for a different time point (2.5h, 23h). There is a third folder that contains the plate reader data for the first screen at two time points (24h and 48h). We tested 7 different 384 well plates on 3 yeast trains with a duplicate of each (and an extra plate across the 3 strains but no duplicate). Plates labeled A and D are TDP43 WT, B and E are TDP43 M337V and C and F are FUS WT. 

More information here: http://mtc.science/on-yeast-and-als-part-3


