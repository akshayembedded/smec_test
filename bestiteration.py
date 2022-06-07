import openpyxl

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
    if(itsum.count(m)==1):
        print("Best iteration - ",m,"\nOccured at ",itsum.index(m)+1)
    else:
        print("Best iterations -",m)
        j=1
        for i in itsum:
            if i==m:
                print("Iteration no: ",j)
            j+=1


path ="Question 1.xlsx"
wb_obj = openpyxl.load_workbook(path) 
# Get workbook active sheet object 
# from the active attribute 
sheet_obj = wb_obj.active
i=1
print("Route 1")
bestiteration(4,9)
print("Route 2")
bestiteration(14,19)
print("Route 3")
bestiteration(24,29)
print("Route 4")
bestiteration(34,39)