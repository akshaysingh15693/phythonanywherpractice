def checkind(abc,wrd):
    a= len(abc)
    b=a-1
    i=0
    while i<=b:
        if abc[i]==wrd:
            print("index is ",i)
            c=i-a
            print("Reverse index is",c)
            i=i+1
        else:
            i=i+1
            