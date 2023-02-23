import numpy as np

a = np.array(["Akshay","Naman","Raja","Manav"],dtype=str)

print("Akshay" in a)

set1 = {"Ram","Ajay","Enjoy"}
set2 = {"Enjoy"}

print(set1 - set2)
c = "Enjoy"

b = str(np.array([c],dtype = str)[0])
print(b)
print(type(b))

print(c)
print(type(c))