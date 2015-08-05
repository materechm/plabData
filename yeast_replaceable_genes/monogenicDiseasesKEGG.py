#does not work at all 

class MonogenicDiseaseManager:
    def __init__(self, fileName, path):
        self.fileName = fileName
        self.path = path
        self.diseases2genes = {} # Maps from diseaseName::String to {geneName::string}


    def extract(self, raw_string, start_marker, end_marker):
      start = raw_string.index(start_marker) + len(start_marker)
      end = raw_string.index(end_marker, start)
      return raw_string[start:end]

    def load_data(self):
        read_file = open(self.fileName)
        fieldnames = read_file.readline()
        reported_genes = set()
        named = False
        genes = False
        pos = 4
        for row in read_file:
          if not row:
              continue
          while named == False:
            if row.startswith("NAME"):
              diseaseName = row[4:]
              named = True
            if named == True:
              if row.startswith("GENE"):
                for x in row:
                  pos +=1
                  if x == "(":
                    reported_genes.append(row[4:pos])
                    pos = 4
                while row.startswith("CARCINOGEN") == False:
                  row = read_file.readline()
                  for x in row:
                    pos +=1
                    if x == "(":
                      reported_genes.append(row[4:pos])
                      pos = 4
          print diseaseName, reported_genes
          group = self.diseases2genes.setdefault(diseaseName, set())
          group.update(reported_genes)
        read_file.close()


    def write_data_to_csv(self):
        write_file = open(self.path, 'w')
        for key in self.diseases2genes:
            if sum(1 for x in self.diseases2genes[key]) == 1:
                print key, self.diseases2genes[key]
                write_file.write('%s, ' %key)
                for x in self.diseases2genes[key]:
                    write_file.write('%s' %x)
                    write_file.write("\n")
        write_file.close()

def main():
    fileName = "/Users/mtchavez/plab/yeast_replaceable_genes/KEGGDisease.txt" #location of OMIM morbidmap in computer
    path = "/Users/mtchavez/plab/yeast_replaceable_genes/monogenicDiseasesKEGG.txt" #location where you want to store file
    mdm = MonogenicDiseaseManager(fileName, path)
    mdm.load_data()
    mdm.write_data_to_csv()

main()
