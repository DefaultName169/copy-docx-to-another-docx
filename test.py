import re
path = open('path.txt', 'r' , encoding='utf-8')

# print(path.readlines())

lines = path.readlines()

pathname = re.sub('^__path__ = |\n', '' , lines[0])
# print(pathname)
output_name = lines[1]
output_name = re.sub(r'^name_of_output = |\n','',output_name)
# print(output_name)

lastlevel = -1
folder = pathname
for i in range(2, len(lines)):
    if lines[i] == '\n' :
        continue
    else :
        level = len(re.findall('\t|    ', lines[i]))
        if level <= lastlevel : 
            for j in range(0, lastlevel - level + 1):
                pathname = re.sub('(.*)/.*', r'\1' , pathname)

        name = re.sub('\n|\t|    ', '' , lines[i])
        names = name.split(' : ')
        name_header = names[-1]
        folder_name = names[0]

        isFolder = not re.search('.docx|.png|.jpg', folder_name)
        ##############################
            #code cu

        ###########################
        pathname = pathname + '/' + folder_name
        print(pathname)
        lastlevel = level