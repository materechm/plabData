import xlsxwriter
import csv
import numpy
from pyper import *
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf
import math

humanSequence = "MELCGLGLPRPPMLLALLLATLLAAMLALLTQVALVVQVAEAARAPSVSAKPGPALWPLPLSVKMTPNLLHLAPENFYISHSPNSTAGPSCTLLEEAFRRYHGYIFGFYKWHHEPAEFQAKTQVQQLLVSIT-LQSECDAFPNIS-SDESYTLLVKEPVAVLKANRVWGALRGLETFSQLVYQDSYG-TFTINESTIIDSPRFSHRGILIDTSRHYLPVKIILKTLDAMAFNKFNVLHWHIVDDQSFPYQSITFPELSNKGSYS-LSHVYTPNDVRMVIEYARLRGIRVLPEFDTPGHTLSWGKGQKDLLTPCYSRQNKLDSFGP--INPTLNTTYSFLTTFFKEISEVFPDQFIHLGGDEVE---FKCWESNPKIQDFMRQKGFGTDFKKLESFYIQKVLDIIA--TINKGSIVWQEVFDDKAKLAPGTIVEVWKD---SAYPEELSRVTASGFPVILSAPWYLDLISYGQDWR-----------KYYKVEPLDFGGTQKQKQLFIGGEACLWGEYVDATNLTPRLWPRASAVGERLWSSKDVRD-MDDAYDRLTRHRCRMVERGIAAQP-LYAGYCNHENM---------"
mouseSequence =  "-------MPQSP-----------RSAPGLLLLQALVSLVSLALVAP---ARLQPALWPFPRSVQMFPRLLYISAEDFSIDHSPNSTAGPSCSLLQEAFRRYYNYVFGFYKRHHGPARFRAEPQLQKLLVSIT-LESECESFPSLS-SDETYSLLVQEPVAVLKANSVWGALRGLETFSQLVYQDSFG-TFTINESSIADSPRFPHRGILIDTSRHFLPVKTILKTLDAMAFNKFNVLHWHIVDDQSFPYQSTTFPELSNKGSYS-LSHVYTPNDVRMVLEYARLRGIRVIPEFDTPGHTQSWGKGQKNLLTPCYNQKTKTQVFGP--VDPTVNTTYAFFNTFFKEISSVFPDQFIHLGGDEVE---FQCWASNPNIQGFMKRKGFGSDFRRLESFYIKKILEIIS--SLKKNSIVWQEVFDDKVELQPGTVVEVWKS---EHYSYELKQVTGSGFPAILSAPWYLDLISYGQDWK-----------NYYKVEPLNFEGSEKQKQLVIGGEACLWGEFVDATNLTPRLWPRASAVGERLWSPKTVTD-LENAYKRLAVHRCRMVSRGIAAQP-LYTGYCNYENKI--------"
zebrafishSequence = "---MLCLLKFTP----------------LFLVVAVCHGWLFGDLFEKQKELDEISLWPLPQKYQSSAVAFKLSAASFQIVHAKQSTAGPSCSLLENAFRRYFEYMFGELKRQEKSRKKAFDSDLSELQVWITSADPECDGYPSLR-TDESYSLSVDETSAVLKAANVWGALRGLETFSQLVYEDDYG-VRNINKTDISDFPRFAHRGILLDSSRHFLPLKVILANLEAMAMNKFNVFHWHIVDDPSFPFMSRTFPELSQKGAYHPFTHVYTPSDVKMVIEFARMRGIRVVAEFDTPGHTQSWGNGIKDLLTPCYSGSSPSGSFGP--VNPILNSSYEFMAQLFKEISTVFPDAYIHLGGDEVD---FSCWKSNPDIQKFMNQQGFGTDYSKLESFYIQRLLDIVA--ATKKGYMVWQEVFDNGVKLKDDTVVEVWKG---NDMKEELQNVTGAGFTTILSAPWYLDYISYGQDWQ-----------RYYKVEPLDFTGTDAQKKLVIGGEACLWGEYVDATNLTPRLWPRASAVAERLWSDASVTD-VGNAYTRLAQHRCRMVRRGIPAEP-LFVGHCRHEYKGL-------"
drosophilaSequence = ""
celegansSequence = "---MRLLIP-------------------ILIFALITTAVTWFYGRDDPDRWSVGGVWPLPKKIVYGSKNRTITYDKIGIDLGDK----KDCDILLSMADNYMNKWLFPFPVEMKTG------GTEDFIITVT-VKDECPSGPPVHGASEEYLLRVSLTEAVINAQTVWGALRAMESLSHLVFYDHKSQEYQIRTVEIFDKPRFPVRGIMIDSSRHFLSVNVIKRQLEIMSMNKLNVLHWHLVDSESFPYTSVKFPELHGVGAYS-PRHVYSREDIADVIAFARLRGIRVIPEFDLPGHTSSWR-GRKGFLTECFDEKG-VETFLPNLVDPMNEANFDFISEFLEEVTETFPDQFLHLGGDEVSDYIVECWERNKKIRKFMEEKGFGNDTVLLENYFFEKLYKIVENLKLKRKPIFWQEVFDNNIP-DPNAVIHIWKGNTHEEIYEQVKNITSQNFPVIVSACWYLNYIKYGADWRDEIRGTAPSNSRYYYCDPTNFNGTVAQKELVWGGIAAIWGELVDNTNIEARLWPRASAAAERLWSPAEKTQRAEDAWPRMHELRCRLVSRGYRIQPNNNPDYCPFEFDEPPATKTEL"
yeastSequence = ""
gene = "HEXB"
file_path = "/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/mendelian_conservation_summary.xlsx"
fileName = "/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/HEXB/HEXBmutationConservationData.xlsx"
jsdivergencePath = "/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/HEXB/jsDivergenceHEXB.csv"
sentropyPath = "/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/HEXB/sEntropyHEXB.csv"
sumofpairsPath = "/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/HEXB/sumOfPairsHEXB.csv"
pathogenic_mutations = "S62L,Y266C,R284*,P417L,Y456S,D459Y,V493G,P504S,R505Q"
pathogenic_mutations = map(str, pathogenic_mutations.split(","))
benign_mutations = "L62S,L72F,K121R,I207V,A543T"
benign_mutations = map(str, benign_mutations.split(","))
polarAA = ["N", "Q", "S", "T", "K", "R", "H", "D", "E"]
nonPolarAA = ["A", "V", "L", "I", "P", "Y", "F", "M", "W", "C"]
hBondingAA = ["C", "W", "N", "Q", "S", "T", "Y", "K", "R", "H", "D", "E"]
sulfurContainingAA = ["C", "M"]
acidicAA = ["D", "E"]
basicAA = ["K", "R"]
ionizableAA = ["D", "E", "H", "C", "Y", "K", "R"]
aromaticAA = ["F", "W", "Y"]
aliphaticAA = ["G", "A", "V", "L", "I", "P"]
aa_symbols = {"Gly": "G", "Ala": "A", "Leu": "L", "Met": "M", "Phe": "F", "Trp": "W", "Lys": "K", "Gln": "Q", "Glu": "E", "Ser": "S", "Pro": "P", "Val": "V",  "Ile": "I", "Cys": "C", "Tyr": "Y", "His": "H", "Arg": "R", "Asn": "N", "Asp": "D", "Thr": "T", "del": "d" }
fullyConservedMutations = 0
fullyConserved = 0
fullyConservedMouse = 0
fullyConservedZebrafish = 0
fullyConservedDrosophila = 0
fullyConservedCelegans = 0
fullyConservedYeast = 0
jsDivergenceScores = []
pathogenicJS = []
benignJS = []
mutationsJS = []
sEntropyScores = []
pathogenicSE = []
benignSE = []
mutationsSE = []
sumOfPairsScores = []
pathogenicSOP = []
benignSOP = []
mutationsSOP = []

def findOccurences(s, ch):
    return [i for i, letter in enumerate(s) if letter == ch]

empty_positions = findOccurences(humanSequence, "-")

def clean_sequences(humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence, empty_positions):
  i = 0
  for x in empty_positions:
    x = x-i
    humanSequence = humanSequence[:x] + humanSequence[(x+1):]
    if mouseSequence != "":
        mouseSequence = mouseSequence[:x] + mouseSequence[(x+1):]
    if zebrafishSequence != "":
        zebrafishSequence = zebrafishSequence[:x] + zebrafishSequence[(x+1):]
    if drosophilaSequence != "":
        drosophilaSequence = drosophilaSequence[:x] + drosophilaSequence[(x+1):]
    if celegansSequence != "":
        celegansSequence = celegansSequence[:x] + celegansSequence[(x+1):]
    if yeastSequence != "":
        yeastSequence = yeastSequence[:x] + yeastSequence[(x+1):]
    i += 1
  return humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence

humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence = clean_sequences(humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence, empty_positions)

print(str(len(humanSequence))) + " lenght human sequence"

#def process_mutations(path, pathogenic_mutations, benign_mutations, gene, aa_symbols):
  #read_file = open(path, 'rb')
  #reader = csv.reader(read_file, delimiter='\t')
  #fieldnames = reader.next()
  #for row in reader:
      #if not row:
          #continue
          #geneSymbol = row[6]
          #print geneSymbol
          #significance = row[7]
          #id = row[10]
          #index = id.index(":p.")+3
          #protein_change_raw = id[index:]
          #protein_change = ""
          #if geneSymbol == gene:
              #if protein_change_raw[len(protein_change_raw-1)] != "=":
                  #if protein_change_raw[len(protein_change_raw-1) == "*"]:
                      #cAA = protein_change_raw[-1:]
                      #oAA = protein_change_raw[:3]
                      #protein_change.append(aa_symbols[oAA])
                      #protein_change.append(protein_change_raw[3:-1])
                      #protein_change.append(aa_symbols[cAA])
                  #cAA = protein_change_raw[-3:]
                  #oAA = protein_change_raw[:3]
                  #if oAA and cAA in aa_symbols:
                      #protein_change.append(aa_symbols[oAA])
                      #protein_change.append(protein_change_raw[3:-3])
                      #protein_change.append(aa_symbols[cAA])
                  #if significance == "Likely benign" or significance == "Benign":
                      #benign_mutations.append(protein_change)
                  #if significance == "Likely pathogenic" or significance == "Pathogenic":
                      #pathogenic_mutations.append(protein_change)
  #return pathogenic_mutations, benign_mutations
  #read_file.close()

#pathogenic_mutations, benign_mutations = process_mutations("/Users/mtchavez/plabData/yeast_replaceable_genes/clinvar.tsv", pathogenic_mutations, benign_mutations, "ABCB7", aa_symbols)

def process_jsdivergence(jsdivergencePath, jsDivergenceScores, empty_positions):
  read_file = open(jsdivergencePath, 'rU')
  reader = csv.reader(read_file, delimiter=',')
  for row in reader:
      if not row:
          continue
      score = row[1]
      jsDivergenceScores.append(score)
  read_file.close()
  i = 0
  for x in empty_positions:
      x = x-i
      jsDivergenceScores = jsDivergenceScores[:x] + jsDivergenceScores[(x+1):]
      i += 1
  return jsDivergenceScores

jsDivergenceScores = process_jsdivergence(jsdivergencePath, jsDivergenceScores, empty_positions)

def process_sentropy(sentropyPath, sEntropyScores, empty_positions):
  read_file = open(sentropyPath, 'rU')
  reader = csv.reader(read_file, delimiter=',')
  for row in reader:
      if not row:
          continue
      score = row[1]
      sEntropyScores.append(score)
  read_file.close()
  i = 0
  for x in empty_positions:
      x = x-i
      sEntropyScores = sEntropyScores[:x] + sEntropyScores[(x+1):]
      i += 1
  return sEntropyScores

sEntropyScores = process_sentropy(sentropyPath, sEntropyScores, empty_positions)

def process_sumofpairs(sumofpairsPath, sumOfPairsScores, empty_positions):
  read_file = open(sumofpairsPath, 'rU')
  reader = csv.reader(read_file, delimiter=',')
  for row in reader:
      if not row:
          continue
      score = row[1]
      sumOfPairsScores.append(score)
  read_file.close()
  i = 0
  for x in empty_positions:
      x = x-i
      sumOfPairsScores = sumOfPairsScores[:x] + sumOfPairsScores[(x+1):]
      i += 1
  return sumOfPairsScores

sumOfPairsScores = process_sumofpairs(sumofpairsPath, sumOfPairsScores, empty_positions)


def create_spreadsheet(humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence, pathogenic_mutations, benign_mutations, gene, fileName, polarAA, nonPolarAA, hBondingAA, sulfurContainingAA, acidicAA, basicAA, ionizableAA, aromaticAA, aliphaticAA, fullyConservedMutations, fullyConserved, fullyConservedMouse, fullyConservedZebrafish, fullyConservedDrosophila, fullyConservedCelegans, fullyConservedYeast, jsDivergenceScores, pathogenicJS, benignJS, mutationsJS, sEntropyScores, pathogenicSE, benignSE, mutationsSE, sumOfPairsScores, pathogenicSOP, benignSOP, mutationsSOP, file_path):
   workbook = xlsxwriter.Workbook(fileName)
   mutationData = workbook.add_worksheet(gene + "mutationConservationData")
   jsDivergence = workbook.add_worksheet("JS Divergence")
   sEntropy = workbook.add_worksheet("Shannon Entropy")
   sumOfPairs = workbook.add_worksheet("Sum of Pairs")
   number = 1
   for x in jsDivergenceScores:
        jsDivergence.write(number-1, 0, number)
        jsDivergence.write(number-1, 1, x)
        number +=1
   number = 1
   for x in sEntropyScores:
        sEntropy.write(number-1, 0, number)
        sEntropy.write(number-1, 1, x)
        number +=1
   number = 1
   for x in sumOfPairsScores:
        sumOfPairs.write(number-1, 0, number)
        sumOfPairs.write(number-1, 1, x)
        number +=1
   mutationData.write(1, 1, "Mouse")
   mutationData.write(1, 2, "Zebrafish")
   mutationData.write(1, 3, "Drosophila")
   mutationData.write(1, 4, "C. Elegans")
   mutationData.write(1, 5, "Yeast")
   mutationData.write(1, 6, "JS Divergence")
   mutationData.write(1, 7, "S Entropy")
   mutationData.write(1, 8, "Sum of Pairs")
   mutationData.merge_range('K2:L2', 'Summary')
   mutationData.write(2, 10, "% fully conserved mice")
   mutationData.write(3, 10, "% fully conserved zebrafish")
   mutationData.write(4, 10, "% fully conserved drosophila")
   mutationData.write(5, 10, "% fully conserved c elegans")
   mutationData.write(6, 10, "% fully conserved yeast")
   mutationData.write(7, 10, "% fully conserved mutations")
   mutationData.write(8, 10, "# fully conserved mutations (fcm)")
   mutationData.write(9, 10, "# fully conserved amino acids (fca)")
   mutationData.write(10, 10, "fcm/faa (percentage)")
   mutationData.write(11, 10, "# pathogenic mutations analyzed")
   mutationData.write(12, 10, "# benign mutations analyzed")
   mutationData.write(13, 10, "# total mutations analyzed")
   mutationData.write(14, 10, "JS divergence average")
   mutationData.write(15, 10, "JS divergence average pathogenic")
   mutationData.write(16, 10, "JS divergence average benign")
   mutationData.write(17, 10, "Shannon entropy average")
   mutationData.write(18, 10, "Shannon entropy average pathogenic")
   mutationData.write(19, 10, "Shannon entropy average benign")
   mutationData.write(20, 10, "Sum of pairs average")
   mutationData.write(21, 10, "Sum of pairs average pathogenic")
   mutationData.write(22, 10, "Sum of pairs average benign")
   mutationData.merge_range('K25:L25', 'Standard deviation')
   mutationData.write(25, 10, "JS divergence SD")
   mutationData.write(26, 10, "JS divergence pathogenic SD")
   mutationData.write(27, 10, "JS divergence benign SD")
   mutationData.write(28, 10, "Shannon entropy SD")
   mutationData.write(29, 10, "Shannon entropy pathogenic SD")
   mutationData.write(30, 10, "Shannon entropy benign SD")
   mutationData.write(31, 10, "Sum of pairs SD")
   mutationData.write(32, 10, "Sum of pairs pathogenic SD")
   mutationData.write(33, 10, "Sum of pairs benign SD")
   mutationData.merge_range('K36:L36', 'p Values')
   redformat = workbook.add_format()
   redformat.set_bg_color('red')
   yellowformat = workbook.add_format()
   yellowformat.set_bg_color('yellow')
   greenformat = workbook.add_format()
   greenformat.set_bg_color('green')
   grayformat = workbook.add_format()
   grayformat.set_bg_color('gray')
   row = 2
   for mutation in benign_mutations:
       oAA = mutation[0]
       pos = int(mutation[1:-1])
       if mouseSequence != "":
           mAA = mouseSequence[pos-1]
       else:
           mAA = "-"
       if zebrafishSequence != "":
           zAA = zebrafishSequence[pos-1]
       else:
           zAA = "-"
       if drosophilaSequence != "":
           dAA = drosophilaSequence[pos-1]
       else:
           dAA = "-"
       if celegansSequence != "":
           cAA = celegansSequence[pos-1]
       else:
           cAA = "-"
       if yeastSequence != "":
           yAA = yeastSequence[pos-1]
       else:
           yAA = "-"
       mutationData.write(row, 6, jsDivergenceScores[pos-1])
       mutationsJS.append(jsDivergenceScores[pos-1])
       mutationData.write(row, 7, sEntropyScores[pos-1])
       mutationsSE.append(sEntropyScores[pos-1])
       mutationData.write(row, 8, sumOfPairsScores[pos-1])
       mutationsSOP.append(sumOfPairsScores[pos-1])
       if oAA != humanSequence[pos-1]:
           mutationData.write(row, 0, mutation, redformat)
       else:
           mutationData.write(row, 0, mutation, greenformat)
           benignSOP.append(sumOfPairsScores[pos-1])
           benignSE.append(sEntropyScores[pos-1])
           benignJS.append(jsDivergenceScores[pos-1])
       if oAA == mAA:
           mutationData.write(row, 1, mAA, greenformat)
           fullyConservedMouse += 1
       elif (oAA and mAA in polarAA) or (oAA and mAA in nonPolarAA) or (oAA and mAA in hBondingAA) or (oAA and mAA in sulfurContainingAA) or (oAA and mAA in acidicAA) or (oAA and mAA in basicAA) or (oAA and mAA in ionizableAA) or (oAA and mAA in aromaticAA) or (oAA and mAA in aliphaticAA):
           mutationData.write(row, 1, mAA, yellowformat)
       elif mAA == '-':
           mutationData.write(row, 1, mAA)
       else:
           mutationData.write(row, 1, mAA, redformat)
       if oAA == zAA:
           mutationData.write(row, 2, zAA, greenformat)
           fullyConservedZebrafish += 1
       elif (oAA and zAA in polarAA) or (oAA and zAA in nonPolarAA) or (oAA and zAA in hBondingAA) or (oAA and zAA in sulfurContainingAA) or (oAA and zAA in acidicAA) or (oAA and zAA in basicAA) or (oAA and zAA in ionizableAA) or (oAA and zAA in aromaticAA) or (oAA and zAA in aliphaticAA):
           mutationData.write(row, 2, zAA, yellowformat)
       elif zAA == '-':
           mutationData.write(row, 2, zAA)
       else:
           mutationData.write(row, 2, zAA, redformat)
       if oAA == dAA:
           mutationData.write(row, 3, dAA, greenformat)
           fullyConservedDrosophila += 1
       elif (oAA and dAA in polarAA) or (oAA and dAA in nonPolarAA) or (oAA and dAA in hBondingAA) or (oAA and dAA in sulfurContainingAA) or (oAA and dAA in acidicAA) or (oAA and dAA in basicAA) or (oAA and dAA in ionizableAA) or (oAA and dAA in aromaticAA) or (oAA and dAA in aliphaticAA):
           mutationData.write(row, 3, dAA, yellowformat)
       elif dAA == '-':
           mutationData.write(row, 3, dAA)
       else:
           mutationData.write(row, 3, dAA, redformat)
       if oAA == cAA:
           mutationData.write(row, 4, cAA, greenformat)
           fullyConservedCelegans += 1
       elif (oAA and cAA in polarAA) or (oAA and cAA in nonPolarAA) or (oAA and cAA in hBondingAA) or (oAA and cAA in sulfurContainingAA) or (oAA and cAA in acidicAA) or (oAA and cAA in basicAA) or (oAA and cAA in ionizableAA) or (oAA and cAA in aromaticAA) or (oAA and cAA in aliphaticAA):
           mutationData.write(row, 4, cAA, yellowformat)
       elif cAA == '-':
           mutationData.write(row, 4, cAA)
       else:
           mutationData.write(row, 4, cAA, redformat)
       if oAA == yAA:
           mutationData.write(row, 5, yAA, greenformat)
           fullyConservedYeast += 1
       elif (oAA and yAA in polarAA) or (oAA and yAA in nonPolarAA) or (oAA and yAA in hBondingAA) or (oAA and yAA in sulfurContainingAA) or (oAA and yAA in acidicAA) or (oAA and yAA in basicAA) or (oAA and yAA in ionizableAA) or (oAA and yAA in aromaticAA) or (oAA and yAA in aliphaticAA):
           mutationData.write(row, 5, yAA, yellowformat)
       elif yAA == '-':
           mutationData.write(row, 5, yAA)
       else:
           mutationData.write(row, 5, yAA, redformat)
       if oAA == mAA == zAA == dAA == cAA == yAA:
           fullyConservedMutations += 1
       row += 1
   for mutation in pathogenic_mutations:
       oAA = mutation[0]
       pos = int(mutation[1:-1])
       if mouseSequence != "":
           mAA = mouseSequence[pos-1]
       else:
           mAA = "-"
       if zebrafishSequence != "":
           zAA = zebrafishSequence[pos-1]
       else:
           zAA = "-"
       if drosophilaSequence != "":
           dAA = drosophilaSequence[pos-1]
       else:
           dAA = "-"
       if celegansSequence != "":
           cAA = celegansSequence[pos-1]
       else:
           cAA = "-"
       if yeastSequence != "":
           yAA = yeastSequence[pos-1]
       else:
           yAA = "-"
       mutationData.write(row, 6, jsDivergenceScores[pos-1])
       mutationsJS.append(jsDivergenceScores[pos-1])
       mutationData.write(row, 7, sEntropyScores[pos-1])
       mutationsSE.append(sEntropyScores[pos-1])
       mutationData.write(row, 8, sumOfPairsScores[pos-1])
       mutationsSOP.append(sumOfPairsScores[pos-1])
       if oAA != humanSequence[pos-1]:
           mutationData.write(row, 0, mutation, redformat)
       else:
           mutationData.write(row, 0, mutation)
           pathogenicSOP.append(sumOfPairsScores[pos-1])
           pathogenicSE.append(sEntropyScores[pos-1])
           pathogenicJS.append(jsDivergenceScores[pos-1])
       if oAA == mAA:
           mutationData.write(row, 1, mAA, greenformat)
           fullyConservedMouse += 1
       elif (oAA and mAA in polarAA) or (oAA and mAA in nonPolarAA) or (oAA and mAA in hBondingAA) or (oAA and mAA in sulfurContainingAA) or (oAA and mAA in acidicAA) or (oAA and mAA in basicAA) or (oAA and mAA in ionizableAA) or (oAA and mAA in aromaticAA) or (oAA and mAA in aliphaticAA):
           mutationData.write(row, 1, mAA, yellowformat)
       elif mAA == '-':
           mutationData.write(row, 1, mAA)
       else:
           mutationData.write(row, 1, mAA, redformat)
       if oAA == zAA:
           mutationData.write(row, 2, zAA, greenformat)
           fullyConservedZebrafish += 1
       elif (oAA and zAA in polarAA) or (oAA and zAA in nonPolarAA) or (oAA and zAA in hBondingAA) or (oAA and zAA in sulfurContainingAA) or (oAA and zAA in acidicAA) or (oAA and zAA in basicAA) or (oAA and zAA in ionizableAA) or (oAA and zAA in aromaticAA) or (oAA and zAA in aliphaticAA):
           mutationData.write(row, 2, zAA, yellowformat)
       elif zAA == '-':
           mutationData.write(row, 2, zAA)
       else:
           mutationData.write(row, 2, zAA, redformat)
       if oAA == dAA:
           mutationData.write(row, 3, dAA, greenformat)
           fullyConservedDrosophila += 1
       elif (oAA and dAA in polarAA) or (oAA and dAA in nonPolarAA) or (oAA and dAA in hBondingAA) or (oAA and dAA in sulfurContainingAA) or (oAA and dAA in acidicAA) or (oAA and dAA in basicAA) or (oAA and dAA in ionizableAA) or (oAA and dAA in aromaticAA) or (oAA and dAA in aliphaticAA):
           mutationData.write(row, 3, dAA, yellowformat)
       elif dAA == '-':
           mutationData.write(row, 3, dAA)
       else:
           mutationData.write(row, 3, dAA, redformat)
       if oAA == cAA:
           mutationData.write(row, 4, cAA, greenformat)
           fullyConservedCelegans += 1
       elif (oAA and cAA in polarAA) or (oAA and cAA in nonPolarAA) or (oAA and cAA in hBondingAA) or (oAA and cAA in sulfurContainingAA) or (oAA and cAA in acidicAA) or (oAA and cAA in basicAA) or (oAA and cAA in ionizableAA) or (oAA and cAA in aromaticAA) or (oAA and cAA in aliphaticAA):
           mutationData.write(row, 4, cAA, yellowformat)
       elif cAA == '-':
           mutationData.write(row, 4, cAA)
       else:
           mutationData.write(row, 4, cAA, redformat)
       if oAA == yAA:
           mutationData.write(row, 5, yAA, greenformat)
           fullyConservedYeast += 1
       elif (oAA and yAA in polarAA) or (oAA and yAA in nonPolarAA) or (oAA and yAA in hBondingAA) or (oAA and yAA in sulfurContainingAA) or (oAA and yAA in acidicAA) or (oAA and yAA in basicAA) or (oAA and yAA in ionizableAA) or (oAA and yAA in aromaticAA) or (oAA and yAA in aliphaticAA):
           mutationData.write(row, 5, yAA, yellowformat)
       elif yAA == '-':
           mutationData.write(row, 5, yAA)
       else:
           mutationData.write(row, 5, yAA, redformat)
       mutationaa = []
       if humanSequence:
           mutationaa.append(oAA)
       if mouseSequence:
           mutationaa.append(mAA)
       if zebrafishSequence:
           mutationaa.append(zAA)
       if drosophilaSequence:
           mutationaa.append(dAA)
       if celegansSequence:
           mutationaa.append(cAA)
       if yeastSequence:
           mutationaa.append(yAA)
       if mutationaa.count(mutationaa[0]) == len(mutationaa):
           fullyConservedMutations += 1
       row += 1
   mutationData.write(row, 0, "Summary", grayformat)
   mutationData.write(row, 1, fullyConservedMouse, grayformat)
   mutationData.write(row, 2, fullyConservedZebrafish, grayformat)
   mutationData.write(row, 3, fullyConservedDrosophila, grayformat)
   mutationData.write(row, 4, fullyConservedCelegans, grayformat)
   mutationData.write(row, 5, fullyConservedYeast, grayformat)
   jsDivergenceScores = map(float, jsDivergenceScores)
   sEntropyScores = map(float, sEntropyScores)
   sumOfPairsScores = map(float, sumOfPairsScores)
   mutationsJS = map(float, mutationsJS)
   mutationsSE = map(float, mutationsSE)
   mutationsSOP = map(float, mutationsSOP)
   pathogenicJS = map(float, pathogenicJS)
   benignJS = map(float, benignJS)
   pathogenicSE = map(float, pathogenicSE)
   benignSE = map(float, benignSE)
   pathogenicSOP = map(float, pathogenicSOP)
   benignSOP = map(float, benignSOP)
   if len(mutationsJS) > 0:
       mutationData.write(row, 6, numpy.mean(mutationsJS), grayformat)
   else:
       mutationData.write(row, 6, "-", grayformat)
   if len(mutationsSE) > 0:
       mutationData.write(row, 7, numpy.mean(mutationsSE), grayformat)
   else:
       mutationData.write(row, 7, "-", grayformat)
   if len(mutationsSOP) > 0:
       mutationData.write(row, 8, numpy.mean(mutationsSOP), grayformat)
   else:
       mutationData.write(row, 8, "-", grayformat)
   totalMutations = len(pathogenic_mutations) + len(benign_mutations)
   if totalMutations > 0:
       mutationData.write(2, 11, str((fullyConservedMouse/float(totalMutations))*100) + "%")
       mutationData.write(3, 11, str((fullyConservedZebrafish/float(totalMutations))*100) + "%")
       mutationData.write(4, 11, str((fullyConservedDrosophila/float(totalMutations))*100) +"%")
       mutationData.write(5, 11, str((fullyConservedCelegans/float(totalMutations))*100) + "%")
       mutationData.write(6, 11, str((fullyConservedYeast/float(totalMutations))*100) + "%")
       mutationData.write(7, 11, str((fullyConservedMutations/float(totalMutations))*100) + "%")
   else:
       mutationData.write(2, 11, "-")
       mutationData.write(3, 11, "-")
       mutationData.write(4, 11, "-")
       mutationData.write(5, 11, "-")
       mutationData.write(6, 11, "-")
       mutationData.write(7, 11, "-")
   mutationData.write(8, 11, fullyConservedMutations)
   redfont = workbook.add_format()
   redfont.set_font_color('red')
   orangefont = workbook.add_format()
   orangefont.set_font_color('orange')
   greenfont = workbook.add_format()
   greenfont.set_font_color('green')
   r = R()
   r.m1 = numpy.mean(jsDivergenceScores)
   r.m2 = numpy.mean(pathogenicJS)
   r.sd1 = numpy.std(jsDivergenceScores)
   r.sd2 = numpy.std(pathogenicJS)
   r.num1 = len(jsDivergenceScores)
   r.num2 = len(pathogenicJS)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p1 = r.p
   mutationData.write(36, 10, "JS div. path. and JS div.")
   if math.isnan(r.p) == False:
       if r.m1 < r.m2:
           if r.p < 0.05:
               mutationData.write(36, 11, r.p, greenfont)
           else:
               mutationData.write(36, 11, r.p, redfont)
       if r.m2 < r.m1:
           if r.p > 0.05:
               mutationData.write(36, 11, r.p, orangefont)
           else:
               mutationData.write(36, 11, r.p, redfont)
   else:
       mutationData.write(36, 11, "-")
   r.m1 = numpy.mean(benignJS)
   r.sd1 = numpy.std(benignJS)
   r.num1 = len(benignJS)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p2 = r.p
   mutationData.write(37, 10, "JS div. path. and JS div. benign")
   if math.isnan(r.p) == False:
     if r.m1 < r.m2:
        if r.p < 0.05:
             mutationData.write(37, 11, r.p, greenfont)
        else:
             mutationData.write(37, 11, r.p, redfont)
     if r.m2 < r.m1:
         if r.p > 0.05:
             mutationData.write(37, 11, r.p, orangefont)
         else:
             mutationData.write(37, 11, r.p, redfont)
   else:
     mutationData.write(37, 11, "-")
   r.m1 = numpy.mean(sEntropyScores)
   r.m2 = numpy.mean(pathogenicSE)
   r.sd1 = numpy.std(sEntropyScores)
   r.sd2 = numpy.std(pathogenicSE)
   r.num1 = len(sEntropyScores)
   r.num2 = len(pathogenicSE)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p3 = r.p
   mutationData.write(38, 10, "S. ent. path. and S. ent.")
   if math.isnan(r.p) == False:
       if r.m1 < r.m2:
           if r.p < 0.05:
               mutationData.write(38, 11, r.p, greenfont)
           else:
               mutationData.write(38, 11, r.p, redfont)
       if r.m2 < r.m1:
           if r.p > 0.05:
               mutationData.write(38, 11, r.p, orangefont)
           else:
               mutationData.write(38, 11, r.p, redfont)
   else:
       mutationData.write(38, 11, "-")
   r.m1 = numpy.mean(benignSE)
   r.sd1 = numpy.std(benignSE)
   r.num1 = len(benignSE)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p4 = r.p
   mutationData.write(39, 10, "S. ent. path. and S. ent. benign")
   if math.isnan(r.p) == False:
       if r.m1 < r.m2:
           if r.p < 0.05:
               mutationData.write(39, 11, r.p, greenfont)
           else:
               mutationData.write(39, 11, r.p, redfont)
       if r.m2 < r.m1:
           if r.p > 0.05:
               mutationData.write(39, 11, r.p, orangefont)
           else:
               mutationData.write(39, 11, r.p, redfont)
   else:
       mutationData.write(39, 11, "-")
   r.m1 = numpy.mean(sumOfPairsScores)
   r.m2 = numpy.mean(pathogenicSOP)
   r.sd1 = numpy.std(sumOfPairsScores)
   r.sd2 = numpy.std(pathogenicSOP)
   r.num1 = len(sumOfPairsScores)
   r.num2 = len(pathogenicSOP)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p5 = r.p
   mutationData.write(40, 10, "S. of pairs path. and S. of pairs")
   if math.isnan(r.p) == False:
       if r.m1 < r.m2:
           if r.p < 0.05:
               mutationData.write(40, 11, r.p, greenfont)
           else:
               mutationData.write(40, 11, r.p, redfont)
       if r.m2 < r.m1:
           if r.p > 0.05:
               mutationData.write(40, 11, r.p, orangefont)
           else:
               mutationData.write(40, 11, r.p, redfont)
   else:
       mutationData.write(40, 11, "-")
   r.m1 = numpy.mean(benignSOP)
   r.sd1 = numpy.std(benignSOP)
   r.num1 = len(benignSOP)
   r('se <- sqrt(sd1*sd1/num1+sd2*sd2/num2)')
   r('t <- (m1-m2)/se')
   r('p <- pt(-abs(t),df=pmin(num1,num2)-1)')
   p6 = r.p
   mutationData.write(41, 10, "S. of pairs path. and S. of pairs benign")
   if math.isnan(r.p) == False:
       if r.m1 < r.m2:
           if r.p < 0.05:
               mutationData.write(41, 11, r.p, greenfont)
           else:
               mutationData.write(41, 11, r.p, redfont)
       if r.m2 < r.m1:
           if r.p > 0.05:
               mutationData.write(41, 11, r.p, orangefont)
           else:
               mutationData.write(41, 11, r.p, redfont)
   else:
       mutationData.write(41, 11, "-")
   sequence_length = len(humanSequence)
   pos = 0
   aalistpos = []
   while pos < sequence_length:
       if humanSequence:
           aalistpos.append(humanSequence[pos])
       if mouseSequence:
           aalistpos.append(mouseSequence[pos])
       if zebrafishSequence:
           aalistpos.append(zebrafishSequence[pos])
       if drosophilaSequence:
           aalistpos.append(drosophilaSequence[pos])
       if celegansSequence:
           aalistpos.append(celegansSequence[pos])
       if yeastSequence:
           aalistpos.append(yeastSequence[pos])
       if aalistpos.count(aalistpos[0]) == len(aalistpos):
           fullyConserved +=1
       pos += 1
       aalistpos = []
   mutationData.write(9, 11, fullyConserved)
   if fullyConservedMutations > 0:
       mutationData.write(10, 11, str(fullyConservedMutations/float(fullyConserved)*100) +"%")
   else:
       mutationData.write(10, 11, "0%")
   mutationData.write(11, 11, len(pathogenic_mutations))
   mutationData.write(12, 11, len(benign_mutations))
   mutationData.write(13, 11, totalMutations)
   jsDivergenceScores = map(float, jsDivergenceScores)
   pathogenicJS = map(float, pathogenicJS)
   benignJS = map(float, benignJS)
   sEntropyScores = map(float, sEntropyScores)
   pathogenicSE = map(float, pathogenicSE)
   benignSE = map(float, benignSE)
   sumOfPairsScores = map(float, sumOfPairsScores)
   pathogenicSOP = map(float, pathogenicSOP)
   benignSOP = map(float, benignSOP)
   mutationData.write(14, 11, numpy.mean(jsDivergenceScores))
   if len(pathogenicJS) > 0:
       mutationData.write(15, 11, numpy.mean(pathogenicJS))
   else:
       mutationData.write(15, 11, "-")
   if len(benignJS) > 0:
       mutationData.write(16, 11, numpy.mean(benignJS))
   else:
       mutationData.write(16, 11, "-")
   mutationData.write(17, 11, numpy.mean(sEntropyScores))
   if len(pathogenicSE) > 0:
       mutationData.write(18, 11, numpy.mean(pathogenicSE))
   else:
       mutationData.write(18, 11, "-")
   if len(benignSE) > 0:
       mutationData.write(19, 11, numpy.mean(benignSE))
   else:
       mutationData.write(19, 11, "-")
   mutationData.write(20, 11, numpy.mean(sumOfPairsScores))
   if len(pathogenicSOP) > 0:
       mutationData.write(21, 11, numpy.mean(pathogenicSOP))
   else:
       mutationData.write(21, 11, "-")
   if len(benignSOP) > 0:
       mutationData.write(22, 11, numpy.mean(benignSOP))
   else:
       mutationData.write(22, 11, "-")
   if len(jsDivergenceScores) > 1:
       mutationData.write(25, 11, numpy.std(jsDivergenceScores))
   elif len(jsDivergenceScores) == 1:
       mutationData.write(25, 11, 0)
   else:
       mutationData.write(25, 11, "-")
   if len(pathogenicJS) > 1:
       mutationData.write(26, 11, numpy.std(pathogenicJS))
   elif len(pathogenicJS) == 1:
       mutationData.write(26, 11, 0)
   else:
       mutationData.write(26, 11, "-")
   if len(benignJS) > 1:
       mutationData.write(27, 11, numpy.std(benignJS))
   elif len(benignJS) == 1:
       mutationData.write(27, 11, 0)
   else:
       mutationData.write(27, 11, "-")
   if len(sEntropyScores) > 1:
       mutationData.write(28, 11, numpy.std(sEntropyScores))
   elif len(sEntropyScores) == 1:
       mutationData.write(28, 11, 0)
   else:
       mutationData.write(28, 11, "-")
   if len(pathogenicSE) > 1:
       mutationData.write(29, 11, numpy.std(pathogenicSE))
   elif len(pathogenicSE) == 1:
       mutationData.write(29, 11, 0)
   else:
       mutationData.write(29, 11, "-")
   if len(benignSE) > 1:
       mutationData.write(30, 11, numpy.std(benignSE))
   elif len(benignSE) == 1:
       mutationData.write(30, 11, 0)
   else:
       mutationData.write(30, 11, "-")
   if len(sumOfPairsScores) > 1:
       mutationData.write(31, 11, numpy.std(sumOfPairsScores))
   elif len(sumOfPairsScores) == 1:
       mutationData.write(31, 11, 0)
   else:
       mutationData.write(31, 11, "-")
   if len(pathogenicSOP) > 1:
       mutationData.write(32, 11, numpy.std(pathogenicSOP))
   elif len(pathogenicSOP) == 1:
       mutationData.write(32, 11, 0)
   else:
       mutationData.write(32, 11, "-")
   if len(benignSOP) > 1:
       mutationData.write(33, 11, numpy.std(benignSOP))
   elif len(benignSOP) == 1:
       mutationData.write(33, 11, 0)
   else:
       mutationData.write(33, 11, "-")
   workbook.close()
   START_ROW = 13 # 0 based (subtract 1 from excel row number)
   col_gene = 0
   rb = open_workbook(file_path)
   r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
   wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
   w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy
   redFont = easyxf('font: color red')
   orangeFont = easyxf('font: color orange')
   greenFont = easyxf('font: color green')
   style_percent = easyxf(num_format_str='0.00%')
   for row_index in range(START_ROW, r_sheet.nrows):
       current_gene = r_sheet.cell(row_index, col_gene).value
       if current_gene == gene:
            w_sheet.write(row_index, 1, numpy.mean(jsDivergenceScores))
            if numpy.mean(pathogenicJS) < numpy.mean(jsDivergenceScores):
                w_sheet.write(row_index, 2, numpy.mean(pathogenicJS), redFont)
            elif numpy.mean(pathogenicJS) < numpy.mean(benignJS):
                w_sheet.write(row_index, 2, numpy.mean(pathogenicJS), orangeFont)
            else:
                w_sheet.write(row_index, 2, numpy.mean(pathogenicJS))
            w_sheet.write(row_index, 3, numpy.mean(benignJS))
            w_sheet.write(row_index, 4, numpy.mean(sEntropyScores))
            if numpy.mean(pathogenicSE) < numpy.mean(sEntropyScores):
                w_sheet.write(row_index, 5, numpy.mean(pathogenicSE), redFont)
            elif numpy.mean(pathogenicSE) < numpy.mean(benignSE):
                w_sheet.write(row_index, 5, numpy.mean(pathogenicSE), orangeFont)
            else:
                w_sheet.write(row_index, 5, numpy.mean(pathogenicSE))
            w_sheet.write(row_index, 6, numpy.mean(benignSE))
            w_sheet.write(row_index, 7, numpy.mean(sumOfPairsScores))
            if numpy.mean(pathogenicSOP) < numpy.mean(sumOfPairsScores):
                w_sheet.write(row_index, 8, numpy.mean(pathogenicSOP), redFont)
            elif numpy.mean(pathogenicSOP) < numpy.mean(benignSOP):
                w_sheet.write(row_index, 8, numpy.mean(pathogenicSOP), orangeFont)
            else:
                w_sheet.write(row_index, 8, numpy.mean(pathogenicSOP))
            w_sheet.write(row_index, 9, numpy.mean(benignSOP))
            w_sheet.write(row_index, 10, len(mutationsJS))
            w_sheet.write(row_index, 11, len(benignJS))
            if len(mutationsJS) > 0:
                w_sheet.write(row_index, 12, len(benignJS)/float(len(mutationsJS)), style_percent)
                w_sheet.write(row_index, 14, fullyConservedMutations/float(len(mutationsJS)), style_percent)
            else:
                w_sheet.write(row_index, 12, "-")
                w_sheet.write(row_index, 14, "-")
            w_sheet.write(row_index, 13, fullyConservedMutations)
            w_sheet.write(row_index, 15, fullyConservedMutations/float(fullyConserved), style_percent)
            species = ""
            if humanSequence:
                species+="h"
            if mouseSequence:
                species+="m"
            if zebrafishSequence:
                species+="z"
            if drosophilaSequence:
                species+="d"
            if celegansSequence:
                species+="c"
            if yeastSequence:
                species+="y"
            w_sheet.write(row_index, 16, species)
            if len(jsDivergenceScores) > 0:
                w_sheet.write(row_index, 17, numpy.std(jsDivergenceScores))
            else:
                w_sheet.write(row_index, 17, "-")
            if len(pathogenicJS) > 0:
                w_sheet.write(row_index, 18, numpy.std(pathogenicJS))
            else:
                w_sheet.write(row_index, 18, "-")
            if len(benignJS) > 0:
                w_sheet.write(row_index, 19, numpy.std(benignJS))
            else:
                w_sheet.write(row_index, 19, "-")
            if len(sEntropyScores) > 0:
                w_sheet.write(row_index, 20, numpy.std(sEntropyScores))
            else:
                w_sheet.write(row_index, 20, "-")
            if len(pathogenicSE) > 0:
                w_sheet.write(row_index, 21, numpy.std(pathogenicSE))
            else:
                w_sheet.write(row_index, 21, "-")
            if len(benignSE) > 0:
                w_sheet.write(row_index, 22, numpy.std(benignSE))
            else:
                w_sheet.write(row_index, 22, "-")
            if len(sumOfPairsScores) > 0:
                w_sheet.write(row_index, 23, numpy.std(sumOfPairsScores))
            else:
                w_sheet.write(row_index, 23, "-")
            if len(pathogenicSOP) > 0:
                w_sheet.write(row_index, 24, numpy.std(pathogenicSOP))
            else:
                w_sheet.write(row_index, 24, "-")
            if len(benignSOP) > 0:
                w_sheet.write(row_index, 25, numpy.std(benignSOP))
            else:
                w_sheet.write(row_index, 25, "-")
            if math.isnan(p1) == False:
                if numpy.mean(jsDivergenceScores) < numpy.mean(pathogenicJS):
                    if p1 < 0.05:
                        w_sheet.write(row_index, 26, p1, greenFont)
                    else:
                        w_sheet.write(row_index, 26, p1, redFont)
                if numpy.mean(pathogenicJS) < numpy.mean(jsDivergenceScores):
                    if p1 > 0.05:
                        w_sheet.write(row_index, 26, p1, orangeFont)
                    else:
                        w_sheet.write(row_index, 26, p1, redFont)
            else:
                w_sheet.write(row_index, 26, "-")
            if math.isnan(p2) == False:
                if numpy.mean(benignJS) < numpy.mean(pathogenicJS):
                    if p2 < 0.05:
                        w_sheet.write(row_index, 27, p2, greenFont)
                    else:
                        w_sheet.write(row_index, 27, p2, redFont)
                if numpy.mean(pathogenicJS) < numpy.mean(benignJS):
                    if p1 > 0.05:
                        w_sheet.write(row_index, 27, p2, orangeFont)
                    else:
                        w_sheet.write(row_index, 27, p2, redFont)
            else:
                w_sheet.write(row_index, 27, "-")
            if math.isnan(p3) == False:
                if numpy.mean(sEntropyScores) < numpy.mean(pathogenicSE):
                    if p3 < 0.05:
                        w_sheet.write(row_index, 28, p3, greenFont)
                    else:
                        w_sheet.write(row_index, 28, p3, redFont)
                if numpy.mean(pathogenicSE) < numpy.mean(sEntropyScores):
                    if p1 > 0.05:
                        w_sheet.write(row_index, 28, p3, orangeFont)
                    else:
                        w_sheet.write(row_index, 28, p3, redFont)
            else:
                w_sheet.write(row_index, 28, "-")
            if math.isnan(p4) == False:
                if numpy.mean(benignSE) < numpy.mean(pathogenicSE):
                    if p4 < 0.05:
                        w_sheet.write(row_index, 29, p4, greenFont)
                    else:
                        w_sheet.write(row_index, 29, p4, redFont)
                if numpy.mean(pathogenicSE) < numpy.mean(benignSE):
                    if p4 > 0.05:
                        w_sheet.write(row_index, 29, p4, orangeFont)
                    else:
                        w_sheet.write(row_index, 29, p4, redFont)
            else:
                w_sheet.write(row_index, 29, "-")
            if math.isnan(p5) == False:
                if numpy.mean(sumOfPairsScores) < numpy.mean(pathogenicSOP):
                    if p5 < 0.05:
                        w_sheet.write(row_index, 30, p5, greenFont)
                    else:
                        w_sheet.write(row_index, 30, p5, redFont)
                if numpy.mean(pathogenicSOP) < numpy.mean(sumOfPairsScores):
                    if p5 > 0.05:
                        w_sheet.write(row_index, 30, p5, orangeFont)
                    else:
                        w_sheet.write(row_index, 30, p5, redFont)
            else:
                w_sheet.write(row_index, 30, "-")
            if math.isnan(p6) == False:
                if numpy.mean(benignSOP) < numpy.mean(pathogenicSOP):
                    if p6 < 0.05:
                        w_sheet.write(row_index, 31, p6, greenFont)
                    else:
                        w_sheet.write(row_index, 31, p6, redFont)
                if numpy.mean(pathogenicSOP) < numpy.mean(benignSOP):
                    if p6 > 0.05:
                        w_sheet.write(row_index, 31, p6, orangeFont)
                    else:
                        w_sheet.write(row_index, 31, p6, redFont)
            else:
                w_sheet.write(row_index, 31, "-")
   wb.save("/Users/mtchavez/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/test.xslx")

create_spreadsheet(humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence, pathogenic_mutations, benign_mutations, gene, fileName, polarAA, nonPolarAA, hBondingAA, sulfurContainingAA, acidicAA, basicAA, ionizableAA, aromaticAA, aliphaticAA, fullyConservedMutations, fullyConserved, fullyConservedMouse, fullyConservedZebrafish, fullyConservedDrosophila, fullyConservedCelegans, fullyConservedYeast, jsDivergenceScores, pathogenicJS, benignJS, mutationsJS, sEntropyScores, pathogenicSE, benignSE, mutationsSE, sumOfPairsScores, pathogenicSOP, benignSOP, mutationsSOP, file_path)
