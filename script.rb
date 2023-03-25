require 'rubyXL'
require 'pg'


# Ouvre le fichier Excel
workbook = RubyXL::Parser.parse('orders.xlsx')

# Sélectionne la feuille 0 
worksheet = workbook[0]

# Stocke les noms de column dans un tableau
names_column = []
for i in 0..(worksheet[0].size - 1)
  names_column << worksheet[0][i].value
end

# test names column
# for i in names_column
#    puts "#{i}"
#    print ""
# end

# Parcours et stocke les données dans un tab de hachages
data = []
for i in 1..(worksheet.sheet_data.size-1)
  row = worksheet[i]
  hachage_data = {}
  for j in 0..(row.size - 1)
    hachage_data[names_column[j].to_sym] = row[j]&.value
  end
  data << hachage_data
end
puts data

# test data
#for i in data
#    puts "#{i}"
#    print ""
#end

# connection db
connection = PG.connect(
  host: 'localhost',
  port: 5432,
  dbname: 'due',
  user: 'due',
  password: 'root'
)

# Insértion dans la db
for hachage_data in data
  column = hachage_data.keys.map(&:to_s).join(', ')
  valeurs = hachage_data.values.map { |v| "'#{v}'" }.join(', ')
  connection.exec("INSERT INTO items (#{column}) VALUES (#{valeurs})")
end





