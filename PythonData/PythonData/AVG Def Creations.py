# sum(nums)/len(nums)

def average(lst):
    total = sum(lst)
    N = len(lst)
    avg = total/N
    return avg

nums = [1,2,3,4,5]
print(average(nums))