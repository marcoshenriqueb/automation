import re

phoneRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d\d')
mo = phoneRegex.search("My number is 415-555-4242")
print("Phone number found: " + mo.group())


phoneRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d\d')
mo = phoneRegex.findall("My number is 415-555-4242, 415-555-4243 or 415-555-4244")
for p in mo:
    print("Another number found: " + p)

namesRegex = re.compile(r'Agent \w+')
sub = namesRegex.sub('CENSORED', 'Agent Alice gave the secret documents to Agent Bob.')
print(sub)
