import os

# working with paths
# print(os.path.join('var', 'www', 'python', 'automate'))
# print(os.listdir(os.getcwd()))

# os.path.exists(path)
# os.path.isfile(path)
# os.path.isdir(path)


# USING PLAIN TEXT FILES
# with open(os.path.join(os.getcwd(), 'hello.txt'), 'a') as myfile:
#     myfile.write("Hello from python again!")
#
# with open(os.path.join(os.getcwd(), 'hello.txt'), 'r') as helloFile:
#     helloContent = helloFile.read()
#     print(helloContent)
    # helloLines = helloFile.readlines()
    # for line in helloLines:
    #     print(line)


# USING BINARY FILES WITH SHELVE
# import shelve

# shelfFile = shelve.open('mydata')
# cats = ['zophie', 'pooka', 'simon']
# shelfFile['cats'] = cats
# shelfFile.close()

# shelfFile = shelve.open('mydata')
# print(list(shelfFile.keys()))
# print(list(shelfFile.values()))
# print(shelfFile['cats'])
# shelfFile.close()


# PPRINT MODULE
# import pprint
# cats = [{'name': 'Zophie', 'desc': 'chubby'}, {'name': 'Pooka', 'desc': 'fluffy'}]
# print(pprint.pformat(cats))
# fileObj = open('myCats.py', 'w')
# fileObj.write('cats = ' + pprint.pformat(cats) + '\n')
# fileObj.close()
