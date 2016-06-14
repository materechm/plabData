import xlsxwriter
import csv
import numpy
from pyper import *
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf
import math
from scipy import stats
import itertools
import pprint

class ClusterScore:
  def __init__(self, conservation_file, cluster_file, output_file):
    self.conservation_file = conservation_file
    self.cluster_file = cluster_file
    self.output_file = output_file
    self.genes = [] #list of genes
    self.JSA = dict() #Maps from geneName::String to JSAverageScore::string
    self.JSP = dict() #Maps from geneName::String to JSPathogenicScore::string
    self.SEA = dict() #Maps from geneName::String to SEAverageScore::string
    self.SEP = dict() #Maps from geneName::String to SEPathogenicScore::string
    self.SOPA = dict() #Maps from geneName::String to SOPAverageScore::string
    self.SOPP = dict() #Maps from geneName::String to SOPPathogenicScore::string
    self.clusters = dict() #Maps from functionName::String to[Genes::string]
    self.jsa_data = 0.7255788165
    self.jsp_data = 0.7664621872
    self.sea_data = 0.584436638
    self.sep_data = 0.6858068153
    self.sopa_data = 2.5983080016
    self.sopp_data = 3.501998831
    self.jsa_sd_data = 0.1237569577
    self.jsp_sd_data = 0.2286742533
    self.sea_sd_data = 0.1277891905
    self.sep_sd_data = 0.1604910682
    self.sopa_sd_data = 0.8802541608
    self.sopp_sd_data = 1.2637220562

  def load_conservation_score(self):
    read_file = open(self.conservation_file)
    reader = csv.reader(read_file, delimiter = '\t')
    fieldnames = reader.next()
    for row in reader:
        if not row:
            continue
        gene = row[0]
        jsAverage = row[1]
        jsPathogenic = row[2]
        seAverage = row[4]
        sePathogenic = row[5]
        sopAverage = row[7]
        sopPathogenic = row[8]
        self.genes.append(gene)
        self.JSA.setdefault(gene, jsAverage)
        self.JSP.setdefault(gene, jsPathogenic)
        self.SEA.setdefault(gene, seAverage)
        self.SEP.setdefault(gene, sePathogenic)
        self.SOPA.setdefault(gene, sopAverage)
        self.SOPP.setdefault(gene, sopPathogenic)
    print "test (dict fetc):" + self.JSA['OCRL']
    read_file.close()

  def load_clusters(self):
    read_file = open(self.cluster_file)
    reader = csv.reader(read_file, delimiter = '\t')
    fieldnames = reader.next()
    for row in reader:
        if not row:
            continue
        clusterName = row[0]
        genes = row[3]
        genes = set(genes.split(', '))
        self.clusters.setdefault(clusterName, genes)
    read_file.close()

  def write_results(self):
    write_file = open(self.output_file, 'w')
    writer = csv.writer(write_file, delimiter='\t')
    writer.writerow(['Gene Group', 'JSD A', 'JSDP A', 'SE A', 'SEP A', 'SoP A', 'SoPP A', 'JSD SD', 'JSDP SD', 'SE SD', 'SEP SD', 'SoP SD', 'SoPP SD', 'JS & JSP p', 'SE & SEP p', 'SoP & SoPP p', 'JS & JS p', 'JSP & JSP p', 'SE & SE p', 'SEP & SEP p', 'SoP & SoP p', 'SoPP & SoPP p' ])
    for cluster in self.clusters:
        genes = self.clusters.get(cluster)
        jsa = []
        jsp = []
        sea = []
        sep = []
        sopa = []
        sopp = []
        for gene in genes:
            jsAverage = self.JSA.get(gene)
            if gene in self.JSP:
                jsPathogenic = self.JSP.get(gene)
            if gene in self.SEA:
                seAverage = self.SEA.get(gene)
            if gene in self.SEP:
                sePathogenic = self.SEP.get(gene)
            if gene in self.SOPA:
                sopAverage = self.SOPA.get(gene)
            if gene in self.SOPP:
                sopPathogenic = self.SOPP.get(gene)
            if jsAverage == "-":
                continue
            else:
                print jsAverage
                if jsAverage:
                    jsAverage = float(jsAverage)
                    jsa.append(jsAverage)
            if jsPathogenic == "-":
                continue
            else:
                jsPathogenic = float(jsPathogenic)
                jsp.append(jsPathogenic)
            if seAverage == "-":
                continue
            else:
                sea.append(float(seAverage))
            if sePathogenic == "-":
                continue
            else:
                sep.append(float(sePathogenic))
            if sopAverage == "-":
                continue
            else:
                sopa.append(float(sopAverage))
            if sopPathogenic == "-":
                continue
            else:
                sopp.append(float(sopPathogenic))
        r = R()
        #jsa = numpy.asarray(jsa)
        print jsa
        #jsa = numpy.mean(jsa)
        r.jsam = numpy.mean(jsa)
        r.jspm = numpy.mean(jsp)
        r.seam = numpy.mean(sea)
        r.sepm = numpy.mean(sep)
        r.sopam = numpy.mean(sopa)
        r.soppm = numpy.mean(sopp)
        r.jsasd = numpy.std(jsa)
        r.jspsd = numpy.std(jsp)
        r.seasd = numpy.std(sea)
        r.sepsd = numpy.std(sep)
        r.sopasd = numpy.std(sopa)
        r.soppsd = numpy.std(sopp)
        r.jsal = len(jsa)
        r.jspl = len(jsp)
        r.seal = len(sea)
        r.sepl = len(sep)
        r.sopal = len(sopa)
        r.soppl = len(sopp)
        r.jsam_data = self.jsa_data
        r.jspm_data = self.jsp_data
        r.seam_data = self.sea_data
        r.sepm_data = self.sep_data
        r.sopam_data = self.sopa_data
        r.soppm_data = self.sopp_data
        r.jsasd_data = self.jsa_sd_data
        r.jspsd_data = self.jsp_sd_data
        r.seasd_data = self.sea_sd_data
        r.sepsd_data = self.sep_sd_data
        r.sopasd_data = self.sopa_sd_data
        r.soppsd_data = self.sopp_sd_data
        r.jsal_data = 259
        r.jspl_data = 235
        r.seal_data = 259
        r.sepl_data = 235
        r.sopal_data = 259
        r.soppl_data = 235
        r('se <- sqrt(jsasd*jsasd/jsal+jspsd*jspsd/jspl)')
        r('t <- (jsam-jspm)/se')
        r('p1 <- pt(-abs(t),df=pmin(jsal,jspl)-1)')
        r('se <- sqrt(seasd*seasd/seal+sepsd*sepsd/sepl)')
        r('t <- (seam-sepm)/se')
        r('p2 <- pt(-abs(t),df=pmin(seal,sepl)-1)')
        r('se <- sqrt(sopasd*sopasd/sopal+soppsd*soppsd/soppl)')
        r('t <- (sopam-soppm)/se')
        r('p3 <- pt(-abs(t),df=pmin(sopal,soppl)-1)')
        r('se <- sqrt(jsasd*jsasd/jsal+jsasd_data*jsasd_data/jspl)')
        r('t <- (jsam-jsam_data)/se')
        r('p4 <- pt(-abs(t),df=pmin(jsal,jsal_data)-1)')
        r('se <- sqrt(jspsd_data*jspsd_data/jspl_data+jspsd*jspsd/jspl)')
        r('t <- (jspm_data-jspm)/se')
        r('p5 <- pt(-abs(t),df=pmin(jspl_data,jspl)-1)')
        r('se <- sqrt(seasd*seasd/seal+seasd_data*seasd_data/seal_data)')
        r('t <- (seam-seam_data)/se')
        r('p6 <- pt(-abs(t),df=pmin(seal,seal_data)-1)')
        r('se <- sqrt(sepsd_data*sepsd_data/sepl_data+sepsd*sepsd/sepl)')
        r('t <- (sepm_data-sepm)/se')
        r('p7 <- pt(-abs(t),df=pmin(sepl_data,sepl)-1)')
        r('se <- sqrt(sopasd*sopasd/sopal+sopasd_data*sopasd_data/sopal_data)')
        r('t <- (sopam-sopam_data)/se')
        r('p8 <- pt(-abs(t),df=pmin(sopal,sopal_data)-1)')
        r('se <- sqrt(soppsd_data*soppsd_data/soppl_data+soppsd*soppsd/soppl)')
        r('t <- (soppm_data-soppm)/se')
        r('p9 <- pt(-abs(t),df=pmin(soppl_data,soppl)-1)')
        writer.writerow([cluster, r.jsam, r.jspm, r.seam, r.sepm, r.sopam, r.soppm, r.jsasd, r.jspsd, r.seasd, r.sepsd, r.sopasd, r.soppsd, r.p1, r.p2, r.p3, r.p4, r.p5, r.p6, r.p7, r.p8, r.p9])
    write_file.close()


def main():
  conservation_file = '/home/jamie/conservation.csv'
  cluster_file = '/home/jamie/clusters.csv'
  output_file = '/home/jamie/output.csv'
  cm = ClusterScore(conservation_file, cluster_file, output_file)
  cm.load_conservation_score()
  cm.load_clusters()
  cm.write_results()

main()
