import xlsxwriter
import csv
import numpy
from pyper import *
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf
import math

humanSequence = "--------MTARGLALGLLLLLL-CPAQVFSQSCVWYGECGIAY--GDKRYNCEYSGP-----PKPLPKDGYDLVQELCPGFFF-GNVSLCCDVRQLQTLKDNLQLPLQFLSRCPSCFYNLLNLFCELTCSPRQSQFLNVTATE-----DYVDPVTNQTKTNVKELQYYVGQSFANAMYNACRDVEAPSSNDKALGLLCGK-DADACNATNWIEYMFNKD--NGQAPFTITPVFSD-F---PVHGMEPMNNATKGCDESVDEVTAPCSCQDCSIVCGPKP-QPPPP---PAPWTILGLDAMYVIMWITYMAFLLVFFGAFFAVWCYRKRYFVSEYTPIDS-NIAFSVNASDKGEAS---CCDPVSAAFEGCLRRLFTRWGSFCVRNPGCVIFFSLVFITACSS-GLVFVRVTTNPVDLWSAPSSQARLEKEYFDQHFGPFFRTEQLIIRAPLTDKHIYQPYPSGADVPFGPPLDIQILHQVLDLQIAIENITAS-YDNETVTLQDICLAPLSPYN-----TNCTILSVLNYFQNSHSVLDHKKG--------DDFFV-------YADYHTHFLYCVRAPASLNDTSLLHDPCLGTFGGPVFPWLVLGGYDDQ--------NYNNATALVITFPVNNYYNDTEKLQRAQAWEKEFINFVKNYKNPNL----TISFTAERSIEDELNRESDSDVFTVVISYAIMFLYISLALGHMKSCR----RLLVDSKVSLGIAGILIVLSSVACSLGVFSYIGLPLTLIVIEVIPFLVLAVGVDNIFILVQAYQRD-ERLQGETLD---QQLGRVLGEVAPSMFLSSFSETVAFFLGALSVMPAVHTFSLFAGLAVFIDFLLQITCFVSLLGLDIKRQEKNRLDIFCCVR-----GAEDGTSVQASESCLFRFFKNSYSPLLLKDWMRPIVIAIFVGVLSFSIAVLNKVDIGLDQSLSMPDDSYMVDYFKSISQYLHAGPPVYFVLEEGHDYTSSKGQNMVCGGM-GCNNDSLVQQIFNAAQLDNYTRIGFAPSSWIDDYFDWVKPQS-SCCRV-DNITDQFCNA-----SVVDPACVRCRPL---------TPEGKQRPQGGDFMRFLPMFLSDNPNPKCGKGGHAAYSSAVNILLGH-GTRVGATYFMTYHTVL--QTSADFIDALKKARLIASNVTETMGIN----------GSAYRVFPYSVFYVFYEQYLTIIDDTIFNLGVSLGAIFLVTMVLLGCELWSAVIMCATIAMVLVNMFGVMWLWGISLNAVSLVNLVMSCGISVEFCSHITRAFTVSMKGSRV---ERAEEALAHMGSSVFSGITLTKFGGIVVLAFAKSQIFQIFYFRMYLAMVLLGATHGLIFLPVLLSYIGPSVNKAKSCATEERYKG-------TERER----LLN-F-------------------------------------------------------------"
mouseSequence = "--------MGAHHPALGLLLLLL-CPAQVFSQSCVWYGECGIAT--GDKRYNCKYSGP-----PKPLPKDGYDLVQELCPGLFF-DNVSLCCDIQQLQTLKSNLQLPLQFLSRCPSCFYNLMTLFCELTCSPHQSQFLNVTATE-----DYFDPKTQENKTNVKELEYFVGQSFANAMYNACRDVEAPSSNEKALGLLCGR-DARACNATNWIEYMFNKD--NGQAPFTIIPVFSD-L---SILGMEPMRNATKGCNESVDEVTGPCSCQDCSIVCGPKP-QPPPP---PMPWRIWGLDAMYVIMWVTYVAFLFVFFGALLAVWCHRRRYFVSEYTPIDS-NIAFSVNSSDKGEAS---CCDPLGAAFDDCLRRMFTKWGAFCVRNPTCIIFFSLAFITVCSS-GLVFVQVTTNPVELWSAPHSQARLEKEYFDKHFGPFFRTEQLIIQAPNTSVHIYEPYPAGADVPFGPPLNKEILHQVLDLQIAIESITAS-YNNETVTLQDICVAPLSPYN-----KNCTIMSVLNYFQNSHAVLDSQVG--------DDFYI-------YADYHTHFLYCVRAPASLNDTSLLHGPCLGTFGGPVFPWLVLGGYDDQ--------NYNNATALVITFPVNNYYNDTERLQRAWAWEKEFISFVKNYKNPNL----TISFTAERSIEDELNRESNSDVFTVIISYVVMFLYISLALGHIQSCS----RLLVDSKISLGIAGILIVLSSVACSLGIFSYMGMPLTLIVIEVIPFLVLAVGVDNIFILVQTYQRD-ERLQEETLD---QQLGRILGEVAPTMFLSSFSETSAFFFGALSSMPAVHTFSLFAGMAVLIDFLLQITCFVSLLGLDIKRQEKNHLDILCCVR-----GADDGQGSHASESYLFRFFKNYFAPLLLKDWLRPIVVAVFVGVLSFSVAVVNKVDIGLDQSLSMPNDSYVIDYFKSLAQYLHSGPPVYFVLEEGYNYSSRKGQNMVCGGM-GCDNDSLVQQIFNAAELDTYTRVGFAPSSWIDDYFDWVSPQS-SCCRL-YNVTHQFCNA-----SVMDPTCVRCRPL---------TPEGKQRPQGKEFMKFLPMFLSDNPNPKCGKGGHAAYGSAVNI-VGD-DTYIGATYFMTYHTIL--KTSADYTDAMKKARLIASNITETMRSK----------GSDYRVFPYSVFYVFYEQYLTIIDDTIFNLSVSLGSIFLVTLVVLGCELWSAVIMCITIAMILVNMFGVMWLWGISLNAVSLVNLVMSCGISVEFCSHITRAFTMSTKGSRV---SRAEEALAHMGSSVFSGITLTKFGGIVVLAFAKSQIFEIFYFRMYLAMVLLGATHGLIFLPVLLSYIGPSVNKAKRHTTYERYRG-------TERER----LLN-F-------------------------------------------------------------"
zebrafishSequence = "-----MLLLGRNHIFRLLLATVL-LSHWVHGQHCIWYGECGNSPNVPEKKLNCNYTGP-----AVPLPNEGQELLQELCPRLVY-ADNRVCCDTQQLNTLKSNIQIPLQYLSRCPACFFNFMTLFCELTCSPRQSQFISVKD--------F---LKEKNQTSVGNVTYYITQTFADAMFNACRDVQAPSSNIKALGLLCGR-DASVCTPQIWIQYMFSIS--NGQVPFGIEPIFTD-V---PVQGMTPMNNRTFNCSQSLDDGSEPCSCQDCSEVCGPTP-VPPPI---PPPWIILGLDAMSFIMWCSYIAFLLIFFGVVLGAWCYRRSVVTSEYGPILDSNQPHSLN--SDDEAS---CCETVGERFENSLRLVFSRWGSLCVRQPLTIILSSLVLICICSA-GLSYMRITTNPVELWSAPSSRARQEKNYFDQHFGPFFRTEQLIITTPWTEEGGFST-ITGDIIPFSPILNLSLLHQVLDLQLEIENLIAE-YKGENVTLKDICVSPLSPYN-----DNCTILSVLNYFQNSHEVLDHEFQ--------DEFFL-------YNDYHTHLLYCASSPTSLDDTSRLHDPCMGTFGGPVFPWLVLGGYEDS--------AYNNATALVITFPVTNYLNDTEKLGKALAWEKEFIRFMKNYENPNL----TVSFSSERSIEDEIDRESNSDVSTIVISYIIMFVYISVALGRINSCR----TLLVDSKISLGIAGILIVLSSVACSLGIFSYIGIPLTLIVIEVIPFLVLAVGVDNIFIIVQTYQRD-ERMPEEELH---QQIGRILGDVAPSMFLSSFSETVAFFLGALSTMPAVRTFSLFAGLAIFIDFLLQISCFVSLLGLDIKRQEANRMDILCCVK-----LSDGQ--EEKSEGWLFRFFKKIYAPFILKDWVRPLVVAVFVGMLSFSIAVVNKVEIGLEQTLSMPDDSYVLNYFGNLSKYLHTGPPVYFVVEDGHDYKTFEGQNAVCGGV-GCNNDSLVQQIYTASLMSNYTRISNVPSSWLDDYFDWVKPQS-TCCRY-YNSTGAFCNA-----SVVDKSCVHCRPM---------TSSGKQRPNGTEFMHFLPMFLSDNPNIKCGKGGHAAYGTAVDL-KDN-NTDVGATYFMSYHTIL--KNSSDFINAMKMARELTDNITQTLSTH----------DKSYKVFPYSVFYVFYEQYLTIVDDTALNLGVSLSAIFIVTAVLLGFELWSAVLVCFTIAMILINMFGVMWLWSISLNAVSLVNLVMSCGISVEFCSHIVRAFSISTRSSRV---ERAEEALAHMGSSVFSGITLTKFGGILILALSKSQIFQIFYFRMYLAIVLLGAAHGLIFLPVLLSYAGPSVNKAKVLAAHNRFVG-------TERER----LIY---------------------------------------------------------------"
drosophilaSequence = "MSPRSPLRISPFGVHILIAAVLF-TLIQSSKQDCVWYGVCNTND--FSHSQNCPYNGT-----AKEMATDGLELLKKRCGFLLENSENKFCCDKNQVELLNKNVELAGNILDRCPSCMENLVRHICQFTCSPKQAEFMHVVATQ-----KN-----KKGDEYISSVDLHISTEYINKTYKSCSQVSVPQTGQLAFDLMCGAYSASRCNPTKWFNFMGDAT--NPYVPFQITYIQHE-PK-SNSNNFTPLNVTTVPCNQAVSSKLPACSCSDCDLSCPQGP-PEPPR---PEPFKIVGLDAYFVIMAAVFLVGVLVFLM---GSFLFTQGSSMDDNFQVDGNDVS--DEMPYSENDS---YFEKLGAHTETFLETFFTKWGTYFASNPGLTLIAGASLVVILGY-GINFIEITTDPVKLWASPNSKSRLEREFFDTKFSPFYRLEQIIIKAVNLPQIVH--NTSNGPYTFGPVFDREFLTKVLDLQEGIKEIN-----ANGTQLKDICYAPLSDDGSEIDVSQCVVQSIWGYFGDDRERLDDHDE--------DNGFN--------VTYLDALYDCISNP----------YLCLAPYGGPVDPAIALGGFLPPGDQLTGSTKFELANAIILTFLVKNHHNK-TDLENALTWEKKFVEFMTNYTKNNMSQYMDIAFTSERSIEDELNRESQSDVLTILVSYLIMFMYIAISLGHVKEFK----RVFIDSKITLGIGGVIIVLASVVSSVGVFGYIGLPATLIIVEVIPFLVLAVGVDNIFILVQTHQRD-QRKPNETLE---QQVGRILGKVGPSMLLTSLSESFCFFLGGLSDMPAVRAFALYAGVALIIDFLLQITCFVSLFTLDTKRREENRMDICCFIK-----GKKP-DSITSNEGLLYKFFSSVYVPFLMKKIVRASVMVIFFAWLCFSIAIAPRIDIGLDQELAMPQDSFVLHYFQSLNENLNIGPPVYFVLKGDLAYTNSSDQNLVCAGQ-YCNDDSVLTQIYLASRHSNQTYIARPASSWIDDYFDWAAAAS-SCCKY-RKDSGDFCPH-------QDTSCLRCNI----------TKNSLLRPEEKEFVKYLPFFLKDNPDDTCAKAGHAAYGGAVRYSNSHERLNIEASYFMAYHTIL--KSSADYFLALESARKISANITQMLQGRLMSNGVPMASALTVEVFPYSVFYVFYEQYLTMWSDTLQSMGISVLSIFVVTFVLMGFDVHSALVVVITITMIVVNLGGLMYYWNISLNAVSLVNLVMAVGISVEFCSHLVHSFATSKSVSQI---DRAADSLSKMGSSIFSGITLTKFAGILVLAFAKSQIFQVFYFRMYLGIVVIGAAHGLIFLPVLLSYIGAPVSNARLRYHSQA--A-------AEHET----ALAGIL------------------------------------------------------------"
celegansSequence = "----------MK--QLLIFCLLFGSIFHHGDAGCIMRGLCQKHTE--NAYGPCVTNDTNVEPTAFDKTHPAYEKMVEFCPHLLT-GDNKLCCTPSQAEGLTKQIAQARHILGRCPSCFDNFAKLWCEFTCSPNQQDFVSISEMKPIEKKEGFTPEYQPAEAYVNTVEYRLSTDFAEGMFSSCKDVTFG--GQPALRVMCTS---TPCTLTNWLEFIGTQNLD-LNIPIHTKFLLYDPIKTPPSDRSTYMNVNFTGCDKSARVGWPACSTSECNKEEYANLIDLDDGKTSGQTCNVHGIACLNIFVMLAFIGSLAVLLCVG---F--VFTSYDEDYTNLRQTQ--------SGEESPKRNRIKRTGAWIHNFMENNARDIGMMAGRNPKSHFFIGCAVLIFCLP-GMIYHKESTNVVDMWSSPRSRARQEEMVFNANFGRPQRYQQIMLLSHRD--------FQSSGKLYGPVFHKDIFEELFDILNAIKNISTQDSDGRTITLDDVCYRPMGPG------YDCLIMSPTNYFQGNKEHLDMKSNKEETVSEDDDAFDYFSSEATTDEWMNHMAACIDQPMSQKTKS--GLSCMGTYGGPSAPNMVFGKNST---------NHQAANSIMMTILVTQRT--EPEIQKAELWEKEFLKFCKEYREKSP--KVIFSFMAERSITDEIENDAKDEIVTVVIALAFLIGYVTFSLGRYFVCENQLWSILVHSRICLGMLSVIINLLSSFCSWGIFSMFGIHPVKNALVVQFFVVTLLGVCRTFMVVKYYAQQRVSMPYMSPDQCPEIVGMVMAGTMPAMFSSSLGCAFSFFIGGFTDLPAIRTFCLYAGLAVLIDVVLHCTIFLALFVWDTQRELNGKPEFFFPYQIKDLLGAYLIGRQRATDTFMTQFFHFQVAPFLMHRMTRIITGIIFIASFITTVILSSKISVGFDQSMAFTEKSYISTHFRYLDKFFDVGPPVFFTVDGELDWHRPDVQNKFCTFP-GCSDTSFGNIMNYAVGHTEQTYLSGEMYNWIDNYLEWISRKS-PCCKVYVHDPNTFCSTNRNKSALDDKACRTCMDFDYVANSYPKSSIMYHRPSIEVFYRHLRHFLEDTPNSECVFGGRASFKDAISFTS---RGRIQASQFMTFHKKLSISNSSDFIKAMDTARMVSRRLERSI-------------DDTAHVFAYSKIFPFYEQYSTIMPILTTQLFITVVGVFGIICVTLGIDVKGAACAVICQVSNYFHIVAFMYIFNIPVNALSATNLVMSSGILIEFSVNVLKGYACSLRQRAK---DRAESTVGSIGPIILSGPVVTMAGSTMFLSGAHLQIITVYFFKLFLITIVSSAVHALIILPILLAFGGSRGHGSSETSTNDNDEQHDACVLSPTAESHISNVEEGILNRPSLLDASHILDPLLKAEGGIDKAIDIITIDRSYPSTPSSLPCTSRMPRAHIEPDLRSL"
yeastSequence = "----------MNVLWIIALVGQL-MRLVQGTATCAMYGNCGKKSV-FGNELPCPVPRSF-E--PPVLSDETSKLLVEVCGEEWK-EVRYACCTKDQVVALRDNLQKAQPLISSCPACLKNFNNLFCHFTCAADQGRFVNITKVE-----KS-----KEDKDIVAELDVFMNSSWASEFYDSCKNIKFSATNGYAMDLIGGG----AKNYSQFLKFLGDAKPMLGGSPFQINYKYDL-A--NEEKEWQEFNDEVYACD----DAQYKCACSDCQESCPHLK-PLKDGVCKVGPLPCFSLSVLIFYTICALFAF----------MWYYLCKRKKNGAMIVDDDIVPE-SGSLDESETNVFESFNNETNFFNGKLANLFTKVGQFSVENPYKILITTVFSIFVFSFIIFQYATLETDPINLWVSKNSEKFKEKEYFDDNFGPFYRTEQIFVVNET-----------------GPVLSYETLHWWFDVENFITEEL---QSSENIGYQDLCFRPTED-------STCVIESFTQYFQGALPNKDS--------------------------WKRELQECGKFP----------VNCLPTFQQPLKTNLLFSDD-----------DILNAHAFVVTLLLTNHT------QSANRWEERLEEYLLDLKVPEG---LRISFNTEISLEKELNN--NNDISTVAISYLMMFLYATWALRRKDGKT----------RLLLGISGLLIVLASIVCAAGFLTLFGLKSTLIIAEVIPFLILAIGIDNIFLITHEYDRNCEQKPEYSID---QKIISAIGRMSPSILMSLLCQTGCFLIAAFVTMPAVHNFAIYSTVSVIFNGVLQLTAYVSILSLYEKRSNYKQIT-----------G-----NEETKESF----LKTFYFKMLTQ---KRLIIIIFSAWFFTSLVFLPEIQFGLDQTLAVPQDSYLVDYFKDVYSFLNVGPPVYMVVKN-LDLTKRQNQQKICGKFTTCERDSLANVLEQ---ERHRSTITEPLANWLDDYFMFLNPQNDQCCRL-KKGTDEVCPP-----SFPSRRCETCFQQGSW------NYNMSGFPEGKDFMEYLSIWIN-APSDPCPLGGRAPYSTALVYN----ETSVSASVFRTAHHPL--RSQKDFIQAYSDGVRISSSF------------------PELDMFAYSPFYIFFVQYQTLGPLTLKLIGSAIILIFFISSVFLQ-NIRSSFLLALVVTMIIVDIGALMALLGISLNAVSLVNLIICVGLGVEFCVHIVRSFTVVPSETKKDANSRVLYSLNTIGESVIKGITLTKFIGVCVLAFAQSKIFDVFYFRMWFTLIIVAALHALLFLPALLSLFGGESYRDDSIEAED--------------------------------------------------------------------------------------"
gene = "NPC1"
file_path = "/media/data/plabData/mendelian_conservation/sequences/all_other_mendelian/mendelian_conservation_summary.xlsx"
fileName = "/media/data/plabData/mendelian_conservation/sequences/all_other_mendelian/NPC1/NPC1Omega_ConservationData.xlsx"
jsdivergencePath = "/media/data/plabData/mendelian_conservation/sequences/all_other_mendelian/NPC1/jsDivergenceNPC1.csv"
sentropyPath = "/media/data/plabData/mendelian_conservation/sequences/all_other_mendelian/NPC1/sEntropyNPC1.csv"
sumofpairsPath = "/media/data/plabData/mendelian_conservation/sequences/all_other_mendelian/NPC1/sumOfPairsNPC1.csv"
pathogenic_mutations = "F1278L, F1278S, N1277S, R1274Q, R1274W, E1273G, R1272H, R1272C, E1271A, E1271Q, T1270K, R1266Q, E1265V, E1265K, A1258V, V1255I, I1251T, I1242V, G1240R, G1236E, L1230S, I1223V, L1213V, V1212L, F1207S, F1207L, T1205K, L1204V, S1200G, M1194V, H1193V, H1193Y, A1190G, A1190V, A1190S, E1189G, A1187V, A1187T, R1186H, R1186C, E1185K, V1184M, R1183H, R1183S, R1183C, M1179V, T1176M, V1165L, V1165M, N1156S, N1156D, V1155I, A1151T, S1148R, G1146S, M1142T, M1138L, M1127V, V1125F, W1122R, L1121I, V1115F, V1115I, M1114I, F1110L, A1108V, A1108T, T1098I, T1098A, I1095T, V1078I, R1077G, G1073S, T1068I, T1068S, E1067K, T1066N, T1066A, I1061T, R1059Q, L1055M, V1044M, T1036M, A1035V, V1033I, H1029L, S1020T, Y1019C, A1018T, H1016R, G1015V, C1011S, P1007A, S1004L, G993E, G992R, G992W, Q991H, E985D, P984L, R978H, R978C, V977I, C976G, P974S, N968S, N961S, N961D, R958Q, S954L, Q953L, K951E, D944N, S940L, P939A, A938T, Y932S, A927V, A926V, A926T, N916S, M912T, M912L, M912V, G911S, G910S, M907V, G904R, S902C, D898N, L893M, A885V, Y882C, Y882H, K877T, D874V, M872V, M866V, L864F, I858T, I858V, V852F, A851T, I850M, L846V, V845F, V843L, F842S, F842L, I841V, I837T, I837V, M834T, D832E, D832G, L830R, L830V, S826Y, K822R, R819H, R819C, E814D, E814K, A812V, V810F, V810I, V810L, E805K, G803S, R794Q, K792E, Q790R, I787V, L783F, S781G, V780M, T777N, Q775P, L773F, L773I, I770V, A767G, A764V, F763S, T759A, H758Y, V757M, A756D, V753M, G749V, L748I, A745E, T743I, S738L, F736L, S734I, V727I, D721N, G717E, R714H, R714C, D712N, R711G, Y709F, V706A, F703I, L695A, F703I, L695V, I690M, I678V, V668L, S666N, V664L, V664M, I663N, L662M, A659V, A659T, I658T, S652L, D651A, D651Y, L648P, R646H, R646C, M642I, L637V, S636A, I630T, Y628H, V624A, V624I, T623N, D618H, R615L, R615C, S608N, T604I, N598S, N593T, N589S, E586K, L577F, K576r, E575Q, Y571C, Y570C, N568H, P566S, A558T, N554K, N554S, Y550C, F537S, T536M, S527R, T526A, T526S, D525V, S522T, A521S, A521T, P520S, A519S, R518W, V517I, C516F, T511M, D508N, D502E, G500E, H497Y, V494M, S491N, V484L, T480A, T477M, P474L, P471L, L469V, V462L, S456F, E451K, I450M, I450V, A449V, H441Y, Q438E, P434S, A427T, S425L, D416G, R411Q, Q407H, T405M, R404Q, F403S, P401S, Q397R, D396E, R389H, V378F, N376D, T375A, V373I, R372Q, S365L, A363V, I361V, C352F, V347I, R341P, R341C, F339C, R337Q, C334S, A329T, V327I, A321V, S316G, N308H, D306N, Y297C, Y297F, R296L, R296Q, R296W, C292Y, W291C, V290A, V290L, L281P, L281F, A278V, A278E, M277V, I271V, A267V, A267G, G264V, P259T, P257S, P255R, P255S, P254Q, P254R, K250T, P249L, V246I, V246F, I245T, S244T, P237S, A236T, T235I, D232G, V231A, N222S, N222D, P219S, M217L, H215R, H215Y, P213S, D211Y, T206S, I205V, A201T, K196E, F194L, M193V, N188S, N185D, A183V, A183T, A183S, D182N, A181T, D180Y, A172V, S167L, R161W, M156T, M156V, Q150R, G149E, G149R, V148L, T137M, V133G, V130I, L121F, R116Q, F101C, S95F, R78Q, R78W, V77I, C75W, N70S, F68L, C63R, Q60H, D57N, G55E, E43K, N41S, D37G, Y35H, G32R, E30V"
pathogenic_mutations = map(str, pathogenic_mutations.split(", "))
benign_mutations = ""
#benign_mutations = map(str, benign_mutations.split(", "))
polarAA = ["N", "Q", "S", "T", "K", "R", "H", "D", "E"]
nonPolarAA = ["A", "V", "L", "I", "P", "Y", "F", "M", "W", "C"]
hBondingAA = ["C", "W", "N", "Q", "S", "T", "Y", "K", "R", "H", "D", "E"]
sulfurContainingAA = ["C", "M"]
acidicAA = ["D", "E"]
basicAA = ["K", "R"]
ionizableAA = ["D", "E", "H", "C", "Y", "K", "R"]
aromaticAA = ["F", "W", "Y"]
aliphaticAA = ["G", "A", "V", "L", "I", "P"]
aa_symbols = {"Gly": "G", "Ala": "A", "Leu": "L", "Met": "M", "Phe": "F", "Trp": "W", "Lys": "K", "Gln": "Q", "Glu": "E", "Ser": "S", "Pro": "P", "Val": "V",  "Ile": "I", "Cys": "C", "NPC1": "Y", "His": "H", "Arg": "R", "Asn": "N", "Asp": "D", "Thr": "T", "del": "d" }
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

#pathogenic_mutations, benign_mutations = process_mutations("/home/jamie/plabData/yeast_replaceable_genes/clinvar.tsv", pathogenic_mutations, benign_mutations, "ABCB7", aa_symbols)

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
   mutationData = workbook.add_worksheet(gene)
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
       	   print(pos)
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
   wb.save("/media/data/plabData/yeast_replaceable_genes/sequences/all_other_mendelian/test.xslx")

create_spreadsheet(humanSequence, mouseSequence, zebrafishSequence, drosophilaSequence, celegansSequence, yeastSequence, pathogenic_mutations, benign_mutations, gene, fileName, polarAA, nonPolarAA, hBondingAA, sulfurContainingAA, acidicAA, basicAA, ionizableAA, aromaticAA, aliphaticAA, fullyConservedMutations, fullyConserved, fullyConservedMouse, fullyConservedZebrafish, fullyConservedDrosophila, fullyConservedCelegans, fullyConservedYeast, jsDivergenceScores, pathogenicJS, benignJS, mutationsJS, sEntropyScores, pathogenicSE, benignSE, mutationsSE, sumOfPairsScores, pathogenicSOP, benignSOP, mutationsSOP, file_path)
