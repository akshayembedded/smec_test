import openpyxl

path ="Question 1.xlsx"
wb_obj = openpyxl.load_workbook(path) 
# Get workbook active sheet object 
# from the active attribute 
sheet_obj = wb_obj.active
def bestiteration(rl,rr,ct=2,cb=7):
    s=0
    best=0
    itsum=[]
    for i in range(rl,rr):
        s=0
        for j in range(ct,cb):
            cell_obj = sheet_obj.cell(row = i, column = j) 
            s+=cell_obj.value
        itsum.append(s)
    m=max(itsum)
    n=min(itsum)
    if(itsum.count(m)==1):
        print("Best iteration - ",m,"\nOccured at ",itsum.index(m)+1)
    else:
        print("Best iterations -",m)
        j=1
        for i in itsum:
            if i==m:
                print("Iteration no: ",j)
            j+=1
    
    if(itsum.count(n)==1):
        print("Worst iteration - ",n,"\nOccured at ",itsum.index(n)+1)
    else:
        print("worst iterations -",n)
        j=1
        for i in itsum:
            if i==n:
                print("Iteration no: ",j)
            j+=1

a=input("1. Route 1\n2. Route 2\n3. Route 3\n4. Route 4\nSelect Route : ")

if(a=="1"):
    bestiteration(4,9)

elif(a=="2"):
    bestiteration(14,19)
elif(a =="3"):
    bestiteration(24,29)
elif(a=="4"):
    bestiteration(34,39)
else:
    print("Wrong input")
