def highlight_duplicate(arr):
    flag = []
    index = []
    for x in arr:
        flag.append(arr.count(x))
    for x in range(len(flag)):
        if flag[x] > 1:
            index.append(x)
    return(index) 