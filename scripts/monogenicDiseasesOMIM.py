import csv

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
        for row in read_file:
            if not row:
              continue
            data = row.split("|", 1)[0].rstrip()
            data = data.rsplit(",", 1)
            diseaseName = data[0]
            reported_genes = self.extract(row, "|", "|")
            reported_genes = set(reported_genes.split(','))
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
    fileName = "/Users/mtchavez/plab/yeast_replaceable_genes/OMIM/morbidmap.txt" #location of OMIM morbidmap in computer
    path = "/Users/mtchavez/plab/yeast_replaceable_genes/monogenicDiseasesOMIM.txt" #location where you want to store file
    mdm = MonogenicDiseaseManager(fileName, path)
    mdm.load_data()
    mdm.write_data_to_csv()

main()
