#Mutation conservation across organisms and chemical modifier screens in ALS yeast 

##Mutation Conservation 
####(folder: yeast_replaceable_genes)
This project started as a follow up to a paper from the Marcotte lab in which they replaced 414 essential yeast genes with their human counterparts and found that in about half of these strains the genes were "replaceable", meaning that the yeast was able to survive with the human version of the gene. The idea was then to find which of these genes were associated with Mendelian diseases and/or Monogenic diseases and based on sequence alignments between yeast, worm, fly, zebrafish, mouse and human, figure out if disease-causing mutations affect conserved or variable amino acid positions (more background info here: http://mtc.science/humanization-of-yeast-genes) 

As a way to test the pipeline I tried to figure out if disease causing mutations affect more conserved amino acids in NPC1, which is the gene that causes Niemann Pick Type C. I got the amino acid sequences from entrez protein and performed an alignment with CLUSTAL W2. Then I got a list of mutations in the NPC1 gene from ClinVar. I got a total of 51 amino acid changes, 6 of which were benign/likely benign. I then searched for the amino acid in which the mutations occur in the sequence alignment and saw whether the amino acid was conserved or not. To see if the amino acids were conserved I used 3 different methods: Jensen Shannon Distribution, Shannon Entropy and Sum of Pairs (more information and results here: http://mtc.science/mutation-conservation-across-organisms-in-npc1)

Later I performed the analysis with the genes that were found to be replaceable and non replaceable to see if this was the case with these genes as well or not and if there was a difference between the two groups. Halfway through this I figured out a way to automate more of the process with a python script which you can see here: https://github.com/materechm/plabData/blob/master/yeast_replaceable_genes/test.py (note - on the script the red coloring based on the amino acid properties DOES NOT work) You can see results here: http://mtc.science/humanization-of-yeast-genes-part-2

I am now in the process of performing this analysis in a set of 260 genes that cause mendelian diseases that involve all organelles. Specific gene data is under sequences, then that is divided into replaceable, non replaceable, and other mendelian. When you go to a specific gene you can see a summary of the mutation conservation for the gene, protein sequences, protein alignment, and the conservation scores. You can see the consolidation of the data here: https://github.com/materechm/plabData/blob/master/yeast_replaceable_genes/sequences/all_other_mendelian/mendelian_conservation_summary.xlsx

##Screens 
####(folder: yeast) 
The folder currently only contains images of the yeast were you can see the gene induction (because it is tagged to YFP). There are two folders, each for a different time point (2.5h, 23h). More information here: http://mtc.science/on-yeast-and-als-part-3


