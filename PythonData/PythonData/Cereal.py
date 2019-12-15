import oS
import CSV

#where is the file?
cerealFile = 'C:\Users\Micha\PythonData\03-Python_3_Activities_01-Stu_CerealCleaner_Resources_cereal.csv'

# open the CSV
with open(Stu_CerealCleaner_Resources_cereal, newline='') as cerealCSV:
    reader = csv.reader(cerealFile, delimiter=',')

   # get the header
    header = next(cerealCSV).split('/')
  # for i in range(len(header)):
  #   print(f'{i}: {header[i]}')
  # look at each row
    print()
    print(f'{“-”*50}')
    print(f'{“name”:>40}|{“fiber”:10}')
    print(f'{“-”*50}')
    for row in reader:
        name = row[0]
        fiber = float(row[7])
        if fiber > 5.0:
        print(f'{name:>40}|{fiber:7.2f} g')
    print(f'{“-”*50}')
    print()