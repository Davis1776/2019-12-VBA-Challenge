names = []
# for _ in range(5):
#    name = input("Please enter the name of someone you know. ")
#    names.append(name)
names = [input("Please enter the name of someone you know. ").lower() for in range(5) ]

# @TODO: Use a list comprehension to create a list of lowercased names
# lowercased = [name for names in names.append]
lowercased = names

# @TODO: Use a list comprehension to create a list of titlecased names
# https://www.tutorialspoint.com/python/string_title.htm
titlecased = [name.title() for name in lowercased]

invitations = [
    f"Dear {name}, please come to the wedding this Saturday!" for name in titlecased]

for invitation in invitations:
    print(invitation)
